#For details on what the script does and how to run it, check: https://www.michev.info/Blog/Post/3555

#Set the authentication details
#oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
#oooooooooo                    REPLACE THIS                    oooooooooo
#oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
$tenantID = "tenant.onmicrosoft.com" #your tenantID or tenant root domain
$appID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #the GUID of your app. For best result, use app with Organization.Read.All and User.ReadWrite.All scope granted
$client_secret = "verylongsecurestring" #client secret for the app
#oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
#oooooooooo      NEVER STORE CREDENTIALS IN PLAIN TEXT!!!      oooooooooo
#oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo

$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $client_secret
    grant_type    = "client_credentials"
}

#Get a token
$authenticationResult = Invoke-WebRequest -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body $body -ErrorAction Stop
$token = ($authenticationResult.Content | ConvertFrom-Json).access_token
$authHeader = @{'Authorization'="Bearer $token"}

#Import the list of users, or generate it dynamically as needed
#$users = Import-Csv .\Users-to-disable.csv

$users = @()
$uri = "https://graph.microsoft.com/v1.0/users?`$filter=assignedLicenses/`$count ne 0&`$count=true&`$select=displayName,mail,userPrincipalName,id,userType,assignedLicenses,licenseAssignmentStates&`$top=999"
#Use licenseAssignmentStates for the assignedByGroup property.
Write-Verbose "Obtaining the list of licensed users"

#needs ConsistencyLevel = eventual
$authHeader["ContentType"] = "application/x-www-form-urlencoded"
$authHeader["ConsistencyLevel"] = "eventual"

do {
    $result = Invoke-WebRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop -Headers $authHeader
    $uri = $($result | ConvertFrom-Json).'@odata.nextLink'
    #If we are getting multiple pages, best add some delay to avoid throttling
    Start-Sleep -Milliseconds 500
    $users += ($result | ConvertFrom-Json).Value
} while ($uri)

#Get a list of all SKUs within the tenant # requires Organization.Read.All at minimum
Write-Verbose "Obtaining the list of SKUs"
$SKUs = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/subscribedSkus/" -Verbose:$VerbosePreference -ErrorAction Stop -Headers $authHeader
$SKUs = $SKUs.Content | ConvertFrom-Json | Select -ExpandProperty value
 
#Provide a list of plans to enable
$plansToEnable = @("MICROSOFTBOOKINGS","b737dad2-2f6c-4c65-90e3-ca563267e8b9")

#Loop over each entry
$count = 1; $PercentComplete = 0;
foreach ($user in $users) {
    #Simple progress indicator
    $ActivityMessage = "Retrieving data for user $($user.displayName). Please wait..."
    $StatusMessage = ("Processing user {0} of {1}: {2}" -f $count, @($users).count, $user.UserPrincipalName)
    $PercentComplete = ($count / @($users).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    Write-Verbose "Processing licenses for user $($user.UserPrincipalName)"
    $lic = $user.assignedLicenses | ConvertTo-Json -Depth 5 | ConvertFrom-Json

    #Go over each assigned license and make toggle services from the $plansToEnable list
    $userLicenses = @();$license = $null
    foreach ($license in $user.assignedLicenses) {
        $SKU = $SKUs | ? {$_.SkuId -eq $license.SkuId}
        Write-Verbose "Processing license $($license.SkuId) ($($SKU.skuPartNumber))"

        #Check if the license is assigned via Group, and if so, skip. Otherwise we will end up assigning the SKU directly...
        if (!($user.licenseAssignmentStates | ? {$_.SkuId -eq $license.skuId}).assignedByGroup) { 
            foreach ($planToEnable in $plansToEnable) {
                if ($planToEnable -notmatch "^[{(]?[0-9A-F]{8}[-]?([0-9A-F]{4}[-]?){3}[0-9A-F]{12}[)}]?$") { $planToEnable = ($SKU.ServicePlans | ? {$_.ServicePlanName -eq "$planToEnable"}).ServicePlanId }
                if (($planToEnable -in $SKU.ServicePlans.ServicePlanId) -and ($planToEnable -in $license.DisabledPlans)) {
                    $license.DisabledPlans = @($license.DisabledPlans | ? {$_ -ne $planToEnable} | sort -Unique)
                    $planToEnableName = ($Sku.servicePlans | ? {$_.ServicePlanId -eq "$planToEnable"}).servicePlanName #move out of the loop...
                    Write-Verbose "Toggled plan $planToEnable ($($planToEnableName)) from license $($license.SkuId) ($($SKU.skuPartNumber))"
                }
        }}
        else { Write-Verbose "License $($license.SkuId) ($($SKU.skuPartNumber)) is assigned via group, no changes will be made." }
        
        $userLicenses += $license           
    }

    #Check if changes are needed
    if (!(Compare-Object $lic $userLicenses -Property disabledPlans,skuId)) { Write-Verbose "No licensing changes needed for user $($user.UserPrincipalName)."; continue }
    Write-Verbose "Updating license assignments for user $($user.UserPrincipalName)."

    #Update license assignment
    try {
        $uri = "https://graph.microsoft.com/v1.0/users/$($user.UserPrincipalName)/assignLicense"
        $body = @{
            "addLicenses" = @($userLicenses)
            "removeLicenses" = @()
        }
        Invoke-WebRequest -Headers $authHeader -Uri $uri -Body ($body | ConvertTo-Json -Depth 5) -Method Post -ErrorAction Stop -Verbose -ContentType 'application/json'
    }
    catch {
        $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
        $errResp = $streamReader.ReadToEnd() | ConvertFrom-Json
        $streamReader.Close()
        $errResp | fl * -Force; continue #catch-all for any unhandled errors
    }
    #Simple anti-throttling control
    Start-Sleep -Milliseconds 200
}

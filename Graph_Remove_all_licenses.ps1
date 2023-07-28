#Set the authentication details
$tenantID = "tenant.onmicrosoft.com" #your tenantID or tenant root domain
$appID = "12345678-1234-1234-1234-1234567890AB" #the GUID of your app. For best result, use app with User.ReadWrite.All scope granted
$client_secret = "XXXXXXXXXXXXXXXXXXX" #client secret for the app

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
$users = Import-Csv .\Users-to-disable.csv

#Loop over each entry
$count = 1; $PercentComplete = 0;
foreach ($user in $users) {
    #Simple progress indicator
    $ActivityMessage = "Retrieving data for user $($user.UserPrincipalName). Please wait..."
    $StatusMessage = ("Processing user {0} of {1}: {2}" -f $count, @($users).count, $user.UserPrincipalName)
    $PercentComplete = ($count / @($users).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    Write-Verbose "Processing licenses for user $($user.UserPrincipalName)"
    #Prepare the query
    $uri = "https://graph.microsoft.com/v1.0/users/$($user.UserPrincipalName)?`$select=id,userPrincipalName,assignedLicenses"
    try { $user = Invoke-WebRequest -Headers $authHeader -Uri $uri -ErrorAction Stop | select -ExpandProperty Content | ConvertFrom-Json }
    catch { Write-Verbose "User $($user.UserPrincipalName) not found, skipping..." ; continue }

    #Check if the user has any licenses applied, skip to the next user if not
    if (!$user.assignedLicenses) { Write-Verbose "No Licenses found for user $($user.UserPrincipalName), skipping..." ; continue }

    #Loop over each assigned license
    foreach ($SKU in $user.assignedLicenses) {
        Write-Verbose "Removing license $($SKU.SkuId) from user $($user.UserPrincipalName)"
        #prepare query
        $uri = "https://graph.microsoft.com/v1.0/users/$($user.UserPrincipalName)/assignLicense"
        $body = @{
            "addLicenses" = @()
            "removeLicenses" = @($SKU.skuId)
        }

        #try to remove the license
        try {
            Invoke-WebRequest -Headers $authHeader -Uri $uri -Body ($body | ConvertTo-Json) -Method Post -ErrorAction Stop -Verbose -ContentType 'application/json' | Out-Null
        }
        catch {
            $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
            $errResp = $streamReader.ReadToEnd() | ConvertFrom-Json
            $streamReader.Close()
            if ($errResp.error.message -eq "User license is inherited from a group membership and it cannot be removed directly from the user.") {
                Write-Verbose "License $($SKU.skuId) is assigned via the group-based licensing feature, either remove the user from the group or unassign the group license, as needed."
                continue
            }
            else {$errResp | fl * -Force; continue} #catch-all for any unhandled errors
    }}
    #Simple anti-throttling control
    Start-Sleep -Milliseconds 200
}
#Requires -Version 3.0
#Requires -Modules @{ ModuleName="Microsoft.Graph.Users"; ModuleVersion="1.19.0" }
#Requires -Modules @{ ModuleName="Microsoft.Graph.Identity.DirectoryManagement"; ModuleVersion="1.19.0" }

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/3555/bulk-enable-specific-services-via-the-graph-api

[CmdletBinding()] #Make sure we can use -Verbose
param([string[]]$UserList)

#region Authentication
#Connect to Graph PowerShell and make sure we run with User.ReadWrite.All permissions (to manage license assignments) and Directory.Read.All (to get all the required details)
if (!(Get-MgContext) -or !((Get-MgContext).Scopes.Contains("User.ReadWrite.All")) -or !((Get-MgContext).Scopes.Contains("Directory.Read.All"))) {
    Write-Verbose "Not connected to the Microsoft Graph or the required permissions are missing!"
    Connect-MgGraph -Scopes Directory.Read.All,User.ReadWrite.All -ErrorAction Stop | Out-Null
}

#Double-check required permissions
if (!((Get-MgContext).Scopes.Contains("User.ReadWrite.All")) -or !((Get-MgContext).Scopes.Contains("Directory.Read.All"))) { Write-Error "The required permissions are missing, please re-consent!"; return }
#endregion Authentication


#region Users
$Users = @()

#If a list of users was provided via the -UserList parameter, only run against a set of users
if ($UserList) {
    Write-Verbose "Running the script against the provided list of users..."
    foreach ($user in $UserList) {
        try {
            $ures = Get-MgUser -UserId $user -ErrorAction Stop -Property displayName,mail,userPrincipalName,id,userType,assignedLicenses,licenseAssignmentStates
            $Users += $ures
        }
        catch {
            Write-Verbose "No match found for provided user entry $user, skipping..."
            continue
        }
    }
}
else {
    #Get the list of all user objects within the tenant.
    Write-Verbose "Running the script against all users in the tenant..."

    $Users = Get-MgUser -All -Filter 'assignedLicenses/$count ne 0' -ConsistencyLevel eventual -CountVariable count -Property displayName,mail,userPrincipalName,id,userType,assignedLicenses,licenseAssignmentStates
}
#endregion Users

#region SKUs
#Get a list of all SKUs within the tenant # requires Organization.Read.All at minimum
$SKUs = Get-MgSubscribedSku
#endregion SKUs

#Provide a list of plans to enable
$plansToEnable = @("MICROSOFTBOOKINGS","b737dad2-2f6c-4c65-90e3-ca563267e8b9")

#Loop over each entry
$out = @(); $count = 1; $PercentComplete = 0;
foreach ($user in $users) {
    #Simple progress indicator
    $ActivityMessage = "Retrieving data for user $($user.displayName). Please wait..."
    $StatusMessage = ("Processing user {0} of {1}: {2}" -f $count, @($users).count, $user.UserPrincipalName)
    $PercentComplete = ($count / @($users).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    Write-Verbose "Processing licenses for user $($user.UserPrincipalName)"
    $lic = $user.assignedLicenses | select disabledPlans,skuId

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

    #Update license assignment
    try {
        Write-Verbose "Updating license assignments for user $($user.UserPrincipalName)."
        Set-MgUserLicense -UserId $user.UserPrincipalName -AddLicenses $userLicenses -RemoveLicenses @() -ErrorAction Stop | Out-Null
        $outtemp = New-Object psobject -Property ([ordered]@{"User" = $user.UserPrincipalName;"Original licenses" = $lic;"Updated licenses" = ($userLicenses | select disabledPlans,SkuId) })
        $out += $outtemp
    }
    catch {
        $_ | fl * -Force; continue #catch-all for any unhandled errors
    }
    #Simple anti-throttling control
    Start-Sleep -Milliseconds 200
}

if ($out) { $out | Out-Default }
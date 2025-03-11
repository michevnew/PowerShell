#Requires -Version 3.0
#The script requires the following permissions:
#    User.Read.All (required for /memberOf)
#    AdministrativeUnit.Read.All (required, used to fetch AU membership)
#    Directory.Read.All (optional, if you don't care about least privilege and don't want to add multiple permissions)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/6437/reporting-on-users-administrative-unit-membership-in-entra-id

param([string[]]$UserList)

#Connect to the Graph with required permissions
Write-Verbose "Connecting to Graph API..."
try {
    Connect-MGgraph -Scopes "User.Read.All","AdministrativeUnit.Read.All" -NoWelcome -ErrorAction Stop -Verbose:$false
}
catch { throw $_ }

#region Users
$Users = @()

#If a list of users was provided via the -UserList parameter, only run against a set of users
if ($UserList) {
    Write-Verbose "Running the script against the provided list of users..."
    foreach ($user in $UserList) {
        try {
            #Make sure the user entry is valid
            $ures = Get-MgUser -UserId $user -ErrorAction Stop

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

    $Users = Get-MgUser -All:$true -ErrorAction Stop
}
#endregion Users

#region AUs
#Cycle over each user and fetch AU membership
$output = [System.Collections.Generic.List[Object]]::new() #output variable
$count = 1; $PercentComplete = 0;
foreach ($u in $Users) {
    #Progress message
    $ActivityMessage = "Retrieving data for user $($u.userPrincipalName). Please wait..."
    $StatusMessage = ("Processing user object {0} of {1}: {2}" -f $count, @($Users).count, $u.id)
    $PercentComplete = ($count / @($Users).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #Get the list of AUs for the user
    Write-Verbose "Fetching AU membership for user $($u.userPrincipalName)..."
    $uAUs = Get-MgUserMemberOfAsAdministrativeUnit -UserId $u.id -ErrorAction Stop

    #If no AUs returned for the user, still write to output
    if (!$uAUs) {
        #prepare the output
        $uInfo = [PSCustomObject][ordered]@{
            "Id" = $u.id
            "UPN" = $u.userPrincipalName
            "AU" = "N/A"
            "AUName" = $null
            "AUType" = $null
            "AURule" = $null
            "AUHiddenMembership" = $null
            "RMAU" = $null
        }

        $output.Add($uInfo)
        continue
    }

    #For each AU returned, output the relevant details
    foreach ($AU in $uAUs) {
        #prepare the output
        $uInfo = [PSCustomObject][ordered]@{
            "Id" = $u.id
            "UPN" = $u.userPrincipalName
            "AU" = $AU.Id
            "AUName" = $AU.displayName
            "AUType" = (&{if ($AU.membershipType -eq "Dynamic") { "Dynamic" } else { "Static" }})
            "AURule" = (&{if ($AU.membershipType -eq "Dynamic") { $AU.membershipRule } else { $null }})
            "AUHiddenMembership" = (&{if ($AU.Visibility) { "Yes" } else { "No" }})
            "RMAU" = (&{if ($AU.IsMemberManagementRestricted) { "Yes" } else { "No" }})
        }

        $output.Add($uInfo)
    }
}
#endregion AUs

#Finally, export to CSV
$output | select * | Export-CSV -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_AUMembership.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
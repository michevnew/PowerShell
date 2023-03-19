#Requires -Version 3.0
#Make sure to fill in all the required variables before running the script
#Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Directory.Read.All

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/4235/reporting-on-users-group-membership-in-azure-ad
[CmdletBinding()] #Make sure we can use -Verbose
param([string[]]$UserList,[switch]$TransitiveMembership=$false)

#region Authentication
try { Connect-MgGraph -Scopes Directory.Read.All | Out-Null }
catch { Write-Error $_ -ErrorAction Stop; return }
#endregion Authentication


#region Users
$Users = @()

#If a list of users was provided via the -UserList parameter, only run against a set of users
if ($UserList) {
    Write-Verbose "Running the script against the provided list of users..."
    foreach ($user in $UserList) {
        try {
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

    $Users = Get-MgUser -All -ErrorAction Stop
}
#endregion Users

#region GroupMembership
#Cycle over each user and fetch group membership
$output = [System.Collections.Generic.List[Object]]::new() #output variable
$i=0; $count = 1; $PercentComplete = 0;
foreach ($u in $Users) {
    #Progress message
    $ActivityMessage = "Retrieving data for user $($u.userPrincipalName). Please wait..."
    $StatusMessage = ("Processing user object {0} of {1}: {2}" -f $count, @($Users).count, $u.id)
    $PercentComplete = ($count / @($Users).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #Simple anti-throttling control
    Start-Sleep -Milliseconds 100

    #Prepare the query depending on the type of membership we are interested in
    if ($TransitiveMembership) { $QueryType = "Get-MgUserTransitiveMemberOf" } else { $QueryType = "Get-MgUserMemberOf" }

    #Obtain the list of groups, taking into account the desired query type and pagination
    Write-Verbose "Processing single user entry $($u.userPrincipalName) with $QueryType query..."
    $uGroups = @()

    $uGroups = Invoke-Expression "$QueryType -UserId $($u.id) -All -Property id,displayName,mailEnabled,securityEnabled,membershipRule,mail,isAssignableToRole,groupTypes" -Verbose

    #If no group objects returned for the user, still write to output
    if (!$uGroups) {
        #prepare the output
        $uInfo = [PSCustomObject][ordered]@{
            "Id" = $u.id
            "UPN" = $u.userPrincipalName
            "Group" = "N/A"
            "GroupName" = $null
            "Mail" = $null
            "RoleAssignable" = $null
            "GroupType" = $null
            "MembershipType" = $null
            "GroupRule" = $null
        }

        $output.Add($uInfo)
        continue
    }

    #For each group returned, output the relevant details
    foreach ($Group in $uGroups) {
        if ($Group.AdditionalProperties.'@odata.type' -ne "#microsoft.graph.group") { continue } #Filter out non-group objects
        
        #prepare the output
        $uInfo = [PSCustomObject][ordered]@{
            "Id" = $u.id
            "UPN" = $u.userPrincipalName
            "Group" = $Group.Id
            "GroupName" = $Group.AdditionalProperties.displayName
            "Mail" = (&{if ($Group.AdditionalProperties.mail) { $Group.AdditionalProperties.mail } else { "N/A" }})
            "RoleAssignable" = (&{if ($Group.AdditionalProperties.isAssignableToRole) { $true } else { $false }})
            "GroupType" = (&{
                if ($Group.AdditionalProperties.groupTypes -eq "Unified" -and $Group.AdditionalProperties.securityEnabled) { "Microsoft 365 (security-enabled)" }
                elseif ($Group.AdditionalProperties.groupTypes -eq "Unified" -and !$Group.AdditionalProperties.securityEnabled) { "Microsoft 365" }
                elseif (!($Group.AdditionalProperties.groupTypes -eq "Unified") -and $Group.AdditionalProperties.securityEnabled -and $Group.mailEnabled) { "Mail-enabled Security" }
                elseif (!($Group.AdditionalProperties.groupTypes -eq "Unified") -and $Group.AdditionalProperties.securityEnabled) { "Azure AD Security" }
                elseif (!($Group.AdditionalProperties.groupTypes -eq "Unified") -and $Group.AdditionalProperties.mailEnabled) { "Distribution" }
                else { "N/A" }
            }) #triple-check this
            "MembershipType" = (&{if ($Group.AdditionalProperties.membershipRule) { "Dynamic" } else { "Assigned" }})
            "GroupRule" = (&{if ($Group.AdditionalProperties.membershipRule) { $Group.AdditionalProperties.membershipRule } else { "N/A" }})
        }

        $output.Add($uInfo)
    }
}
#endregion GroupMembership

#Finally, export to CSV
$output | select * #| Export-CSV -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_AADGroupMembership.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
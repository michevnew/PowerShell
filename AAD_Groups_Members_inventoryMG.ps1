#Requires -Version 3.0
#Requires -Modules @{ ModuleName="Microsoft.Graph.Groups"; ModuleVersion="1.19.0" }
#Requires -Modules @{ ModuleName="Microsoft.Graph.Users"; ModuleVersion="1.19.0" }

[CmdletBinding()] #Make sure we can use -Verbose
param([string[]]$GroupList,[switch]$TransitiveMembership=$false)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/4357/report-on-azure-ad-group-members-via-the-graph-api

#region Authentication
try { Connect-MgGraph -Scopes Directory.Read.All | Out-Null }
catch { Write-Error $_ -ErrorAction Stop; return }
#endregion Authentication

if (!(Get-MgContext).Scopes.Contains("Directory.Read.All")) { Write-Error "The required permissions are missing, please re-consent!" -ErrorAction Stop }

Select-MgProfile beta #needed to include SP objects in the output

#region Groups
$Groups = @()

#If a list of groups was provided via the -GroupList parameter, only run against a set of groups
if ($GroupList) {
    Write-Verbose "Running the script against the provided list of groups..."
    foreach ($group in $GroupList) {
        try {
            $gres = Get-MgGroup -GroupId $group -Property id,displayName,groupTypes,securityEnabled,mailEnabled,membershipRule,isAssignableToRole,mail,assignedLicenses -Expand 'owners($select=id,userPrincipalName)' -ErrorAction Stop
            $Groups += $gres
        }
        catch {
            Write-Verbose "No match found for provided group entry $group, skipping..."
            continue
        }
    }
    Write-Verbose "The following list of groups will be used: $($Groups.displayName -join ",")"
}
else {
    #Get the list of all user objects within the tenant. The SDK should handle pagination?
    Write-Verbose "Running the script against all groups in the tenant..."

    $Groups = Get-MgGroup -All -Property id,displayName,groupTypes,securityEnabled,mailEnabled,membershipRule,isAssignableToRole,mail,assignedLicenses -Expand 'owners($select=id,userPrincipalName)' -ErrorAction Stop
}
#endregion Groups

#region GroupMembership
#Cycle over each group and fetch group membership
$output = [System.Collections.Generic.List[Object]]::new() #output variable for expanded CSV (one line per member)
$output2 = [System.Collections.Generic.List[Object]]::new() #output variable for summary CSV (one line per group)
$count = 1; $PercentComplete = 0;
foreach ($g in $Groups) {
    #Progress message
    $ActivityMessage = "Retrieving data for group $($g.displayName). Please wait..."
    $StatusMessage = ("Processing group object {0} of {1}: {2}" -f $count, @($Groups).count, $g.id)
    $PercentComplete = ($count / @($Groups).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #Simple anti-throttling control
    Start-Sleep -Milliseconds 200

    #Set generic group properties to avoid re-evaluating them
    $g | Add-Member -MemberType NoteProperty -Name GroupType -Value (&{
        if ($g.groupTypes -eq "Unified" -and $g.securityEnabled) { "Microsoft 365 (security-enabled)" }
        elseif ($g.groupTypes -eq "Unified" -and !$g.securityEnabled) { "Microsoft 365" }
        elseif (!($g.groupTypes -eq "Unified") -and $g.securityEnabled -and $g.mailEnabled) { "Mail-enabled Security" }
        elseif (!($g.groupTypes -eq "Unified") -and $g.securityEnabled) { "Azure AD Security" }
        elseif (!($g.groupTypes -eq "Unified") -and $g.mailEnabled) { "Distribution" }
        else { "N/A" }
    })

    #Prepare the query depending on the type of membership we are interested in
    if ($TransitiveMembership) { $QueryType = "Get-MgGroupTransitiveMember" } else { $QueryType = "Get-MgGroupMember" }

    #Obtain the list of members, taking into account the desired query type and pagination
    Write-Verbose "Processing single group entry $($g.displayName) with $QueryType query..."
    $gMembers = @()

    #We use /beta here, as /v1.0 does not return service principal objects yet
    $gMembers = Invoke-Expression "$QueryType -GroupId $($g.id) -All -Property id,displayName,mailEnabled,securityEnabled,membershipRule,mail,isAssignableToRole,groupTypes,userPrincipalName,userType,deviceId"

    #prepare the output for the expanded CSV
    $uInfo = [PSCustomObject][ordered]@{
        "Id" = $g.id
        "DisplayName" = $g.displayName
        "GroupType" = $g.groupType
        "Owners" = (&{if ($g.owners) { $($g.Owners.AdditionalProperties.userPrincipalName -join ",") } else { "N/A" }})
        "PrimarySmtpAddress" = (&{if ($g.mail) { $g.mail } else { "N/A" }})
        "RoleAssignable" = (&{if ($g.isAssignableToRole) { $true } else { $false }})
        "AssignedLicenses" = (&{if ($g.assignedLicenses) { ($g.assignedLicenses.skuId -join ",") } else { $false }})
        "MembershipType" = (&{if ($g.membershipRule) { "Dynamic" } else { "Assigned" }})
        "MembershipRule" = (&{if ($g.membershipRule) { $g.membershipRule } else { "N/A" }})
        "MemberId" = $null
        "MemberDisplayName" = $null
        "MemberType" = $null
        "MemberMail" = $null
        "MemberIdentifier" = $null
    }
    if (!$gMembers) { $output.Add($uInfo) } #add the "empty" value

    #For each member returned, include the relevant details
    $j = 0;$usermembers = @();$groupmembers = @();$devicemembers = @();$contactmembers = @();$SPmembers = @();
    foreach ($Member in $gMembers) {
        $j++ #cheap member count that accounts for unhandled member types
        $uInfo.MemberId = $Member.id
        $uInfo.MemberDisplayName = $Member.AdditionalProperties.displayName
        if ($Member.AdditionalProperties.userType -eq "Guest") { $uInfo.MemberType = "Guest" }
        else { $uInfo.MemberType = $Member.AdditionalProperties.'@odata.type'.Split(".")[-1] }
        $uInfo.MemberMail = (&{if ($Member.AdditionalProperties.mail) { $Member.AdditionalProperties.mail } else { "N/A" }})

        #add to the lists used by the summary CSV file
        switch ($Member.AdditionalProperties.'@odata.type'.Split(".")[-1]) {
            "user" { $uInfo.MemberIdentifier = $Member.AdditionalProperties.userPrincipalName; $usermembers += $Member.AdditionalProperties.userPrincipalName }
            "group" { $uInfo.MemberIdentifier = $Member.id; $groupmembers += $Member.id }
            "device" { $uInfo.MemberIdentifier = $Member.AdditionalProperties.deviceId; $devicemembers += $Member.AdditionalProperties.deviceId }
            "orgContact" { $uInfo.MemberIdentifier = $Member.AdditionalProperties.mail; $contactmembers += $Member.AdditionalProperties.mail }
            "servicePrincipal" { $uInfo.MemberIdentifier = $Member.id; $SPmembers += $Member.id }
            default { Write-Verbose "Unhandled scenario" }
        }

        $output.Add($uInfo.psobject.Copy()) #!
    }

    #prepare the output for summary CSV
    $uInfo2 = [PSCustomObject][ordered]@{
        "Id" = $g.id
        "DisplayName" = $g.displayName
        "GroupType" = $g.GroupType
        "Owners" = (&{if ($g.owners) { $($g.Owners.AdditionalProperties.userPrincipalName -join ",") } else { "N/A" }})
        "HasNestedGroups" = &{If ($groupmembers) { $groupmembers.Count } else {$false} }
        "PrimarySmtpAddress" = (&{if ($g.mail) { $g.mail } else { "N/A" }})
        "RoleAssignable" = (&{if ($g.isAssignableToRole) { $true } else { $false }})
        "MembershipType" = (&{if ($g.membershipRule) { "Dynamic" } else { "Assigned" }})
        "MembershipRule" = (&{if ($g.membershipRule) { $g.membershipRule } else { "N/A" }})
        "AssignedLicenses" = (&{if ($g.assignedLicenses) { ($g.assignedLicenses.skuId -join ",") } else { $false }})
        "MemberCountTotal" = $j
        "UserMemberCount" = $usermembers.count
        "GroupMemberCount" = $groupmembers.count
        "DeviceMemberCount" = $devicemembers.count
        "ContactMemberCount" = $contactmembers.count
        "SPMemberCount" = $SPmembers.count
        "UserMembers" = &{If ($usermembers) { $usermembers -join ","}}
        "GroupMembers" = &{If ($groupmembers) { $groupmembers -join ","}}
        "DeviceMembers" = &{If ($devicemembers) { $devicemembers -join ","}}
        "ContactMembers" = &{If ($contactmembers) { $contactmembers -join ","}}
        "SPmembers" = &{If ($SPmembers) { $SPmembers -join ","}}
    }

    $output2.Add($uInfo2)
}
#endregion GroupMembership

#Finally, export to CSV
Write-Verbose "Writing output to CSV files..."
$output | select * | Export-CSV -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_AADGroupMembersExpanded.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
$output2 | select * | Export-CSV -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_AADGroupMembers.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
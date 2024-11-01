#Requires -Version 3.0
#The script requires the following permissions:
#    Directory.Read.All (required)
#    RoleManagement.Read.Directory (optional, needed to retrieve PIM eligible role assignments)
#    PrivilegedEligibilitySchedule.Read.AzureADGroup (optional, needed to retrieve Privileged Access Group assignments)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/5958/reporting-on-ent…ts-including-pim

[CmdletBinding()] #Make sure we can use -Verbose
Param([switch]$IncludePIMEligibleAssignments, #Indicate whether to include PIM elibigle role assignments in the output.
      [switch]$IncludePAGAssignments #Indicate whether to include Privileged Access Group assignments in the output.
)

#region Authentication
#Determine the required scopes, based on the parameters passed to the script
$RequiredScopes = switch ($PSBoundParameters.Keys) {
    "IncludePIMEligibleAssignments" { "RoleManagement.Read.Directory" }
    "IncludePAGAssignments" { "PrivilegedEligibilitySchedule.Read.AzureADGroup" } #Also requires Global reader or Privileged role administrator for the current user!
    Default { "Directory.Read.All" }
}

#Connect to the Graph API
Write-Verbose "Connecting to Graph API..."
try {
    if ($IncludePIMEligibleAssignments -or $IncludePAGAssignments) { Import-Module Microsoft.Graph.Identity.Governance -Verbose:$false -ErrorAction Stop }
    Connect-MgGraph -Scopes $RequiredScopes -Verbose:$false -ErrorAction Stop -NoWelcome
}
catch { throw $_ }

#Check if we have all the required permissions
$CurrentScopes = (Get-MgContext).Scopes
if ($RequiredScopes | ? {$_ -notin $CurrentScopes }) { Write-Error "The access token does not have the required permissions, rerun the script and consent to the missing scopes!" -ErrorAction Stop }
#endregion Authentication

#region Roles
Write-Verbose "Collecting role assignments..."
#Use the Get-MgRoleManagementDirectoryRoleAssignment cmdlet to collect a list of all role assignments. We cannot expand multiple properties, so we do two passes here.
$roles = Get-MgRoleManagementDirectoryRoleAssignment -All -ExpandProperty Principal -Verbose:$false -ErrorAction Stop #$expand=* is BROKEN
$roles1 = Get-MgRoleManagementDirectoryRoleAssignment -All -ExpandProperty roleDefinition -Verbose:$false -ErrorAction Stop #fix to also fetch the roleDefinition
foreach ($role in $roles) { Add-Member -InputObject $role -MemberType NoteProperty -Name roleDefinition1 -Value ($roles1 | ? {$_.id -eq $role.id}).roleDefinition } #and another fix needed as PowerShell populates empty roleDefinition property...

#Use the Get-MgRoleManagementDirectoryRoleEligibilitySchedule cmdlet to collect a list of all PIM eligible role assignments.
if ($IncludePIMEligibleAssignments) {
    Write-Verbose "Collecting PIM eligible role assignments..."

    $roles += (Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty roleDefinition,principal -Verbose:$false -ErrorAction Stop | select id,principalId,directoryScopeId,roleDefinitionId,status,principal,@{n="roleDefinition1";e={$_.roleDefinition}})
    #Use the Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance cmdlet to collect a list of all PIM activated role assignments.
    $roleactivations = Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -All -Filter "AssignmentType eq 'Activated'" -Verbose:$false -ErrorAction Stop

    #If an eligible role is assigned, it will appear as Permanent in the output of Get-MgRoleManagementDirectoryRoleAssignment, so we need some clean up
    foreach ($roleactivation in $roleactivations) {
        #This should give us the role assignments that are eligible AND activated
        $roles | ? {($_.id -eq $roleactivation.RoleAssignmentOriginId) -and ($_.Id -eq $roleactivation.Id)} | % { Add-Member -InputObject $_ -MemberType NoteProperty -Name "Duplicate" -Value $true }
    }
}

if (!$roles) { Write-Verbose "No valid role assignments found, verify the required permissions have been granted?"; return }

$rtemp = $roles | ? {!$_.Duplicate}
Write-Verbose "A total of $($rtemp.count) role assignments were found"
Write-Verbose "$(($rtemp | ? {$_.directoryScopeId -eq "/"}).Count) are tenant-wide and $(($rtemp | ? {$_.directoryScopeId -ne "/"}).Count) are AU-scoped."
Write-Verbose "$(($rtemp | ? {!$_.status}).Count) roles are permanently assigned, you might want to address that!"
#endregion Roles

#region PAG
if ($IncludePAGAssignments) {
    #Get the set of roles with Group principal
    $Proles = $roles | ? {$_.Principal.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group'} #not necessarily PIM-managed group (can be "old" PAG, in both cases the role-assignable flag should be set)
    if (!$Proles) { Write-Verbose "No role assignments with Group principal found, skipping PAG collection" }

    foreach ($role in $Proles) {
        Write-Verbose "Collecting Privileged Access Group members for $($role.PrincipalId) ..."
        #Get the list of permanent/active members, easily done via the Get-MgGroupTransitiveMember cmdlet (with the added benefit of expanding nested groups)
        $dMembers = @{};$dMembersId = @()
        foreach ($member in (Get-MgGroupTransitiveMember -GroupId $role.PrincipalId -Property id,displayName,userPrincipalName -Verbose:$false -ErrorAction Stop)) {
            $dMembers[$member.Id] = $member.AdditionalProperties.userPrincipalName
            if ($member.AdditionalProperties.userPrincipalName) { $dMembersId += $member.AdditionalProperties.userPrincipalName }
            else { $dMembersId += "$($member.AdditionalProperties.displayName) ($($member.Id))" }
        }
        $role | Add-Member -MemberType NoteProperty -Name "Active group members" -Value $dMembers
        $role | Add-Member -MemberType NoteProperty -Name "Active group members IDs" -Value ($dMembersId -join ";")

        #Get the list of eligible members, done via the Get-MgBetaIdentityGovernancePrivilegedAccessGroupEligibilitySchedule cmdlet. #NOT expanding groups here
        #If a member is both eligible and active, it will appear in both lists!
        $eMembers = @{};$eMembersId = @()
        foreach ($member in (Get-MgIdentityGovernancePrivilegedAccessGroupEligibilitySchedule -Filter "groupId eq '$($role.principalId)'" -ExpandProperty principal -Verbose:$false -ErrorAction Stop)) {
            $eMembers[$member.principal.Id] = $member.principal.AdditionalProperties.userPrincipalName
            $eMembersId += if ($member.principal.AdditionalProperties.userPrincipalName) { $member.principal.AdditionalProperties.userPrincipalName } else { "$($member.principal.AdditionalProperties.displayName) ($($member.principal.Id))" }
        }
        $role | Add-Member -MemberType NoteProperty -Name "Eligible group members" -Value $eMembers
        $role | Add-Member -MemberType NoteProperty -Name "Eligible group members IDs" -Value ($eMembersId -join ";")
    }
}
#endregion PAG

#region Output
#prepare the script output
Write-Verbose "Preparing the output..."
$report = @()
foreach ($role in $roles) {
    #Get rid of the duplicate entries
    if ($role.Duplicate) { continue }

    if (!$role.status) { #if the role is permanently assigned, we don't need to check the role activations
        $role | Add-Member -MemberType NoteProperty -Name "Start time" -Value "Permanent"
        $role | Add-Member -MemberType NoteProperty -Name "End time" -Value "Permanent"
        $role | Add-Member -MemberType NoteProperty -Name "AssignmentType" -Value "Permanent"
    } else { #otherwise, we need to check the role activations
        if ($role.principal.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group') {
            $activeRole = $roleactivations | ? {($_.roleDefinitionId -eq $role.roleDefinitionId) -and ($role."Active group members".ContainsKey($_.principalId)) -and ($_.MemberType -eq "Group")}
            $role | Add-Member -MemberType NoteProperty -Name "Activated for" -Value (($activeRole | % { $($role."Active group members"[$_.principalId]) }) -join ";")
        }
        else {
            $activeRole = $roleactivations | ? {($_.roleDefinitionId -eq $role.roleDefinitionId) -and ($_.PrincipalId -eq $role.PrincipalId)}
            $role | Add-Member -MemberType NoteProperty -Name "Activated for" -Value $null
        }

        $role | Add-Member -MemberType NoteProperty -Name "Start time" -Value (&{if ($activeRole.startDateTime) {(Get-Date($activeRole.startDateTime | select -Unique | Sort-Object | select -First 1) -format g)} else {$null}})
        $role | Add-Member -MemberType NoteProperty -Name "End time" -Value (&{if ($activeRole.endDateTime) {(Get-Date($activeRole.endDateTime | select -Unique | Sort-Object -Descending | select -First 1) -format g)} else {$null}})
        $role | Add-Member -MemberType NoteProperty -Name "AssignmentType" -Value (&{if ($activeRole.startDateTime) {"Eligible (Active)"} else {"Eligible"}})
    }

    $reportLine=[ordered]@{
        "Principal" = switch ($role.principal.AdditionalProperties.'@odata.type') {
            '#microsoft.graph.user' {$role.principal.AdditionalProperties.userPrincipalName}
            '#microsoft.graph.servicePrincipal' {$role.principal.AdditionalProperties.appId}
            '#microsoft.graph.group' {$role.principalid}
        }
        "PrincipalDisplayName" = $role.principal.AdditionalProperties.displayName
        "PrincipalType" = $role.principal.AdditionalProperties.'@odata.type'.Split(".")[-1]
        "AssignedRole" = $role.roleDefinition1.displayName
        "AssignedRoleScope" = $role.directoryScopeId
        "AssignmentType" = $role.AssignmentType
        "AssignmentStartDate" = $role.'Start time'
        "AssignmentEndDate" = $role.'End time'
        "ActiveGroupMembers" = $role.'Active group members IDs'
        "EligibleGroupMembers" = $role.'Eligible group members IDs'
        "GroupEligibleAssignmentActivatedFor" = $role.'Activated for' #Permanently assigned group members will not show here, that's the expected behavior!
        "IsBuiltIn" = $role.roleDefinition1.isBuiltIn
        "RoleTemplate" = $role.roleDefinition1.templateId
        #"AllowedActions" = $role.roleDefinition1.RolePermissions.allowedResourceActions -join ";"
    }
    $report += @([pscustomobject]$reportLine)
}
#endregion Output

#format and export
$report | Sort-Object PrincipalDisplayName | Export-CSV -nti -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_AzureADRoleInventory.csv"
Write-Verbose "Output exported to $($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_AzureADRoleInventory.csv"

#LIST all PAG
#Connect-MgGraph -Scopes PrivilegedAccess.Read.AzureADGroup
#Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/privilegedAccess/aadGroups/resources"
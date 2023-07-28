#Requires -Version 3.0
[CmdletBinding(SupportsShouldProcess)] #Make sure we can use -WhatIf and -Verbose
Param([switch]$IncludePIMEligibleAssignments) #Indicate whether to include PIM elibigle role assignments in the output. NOTE: Currently the RoleManagement.Read.Directory scope seems to be required!

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/3958/generate-a-report-of-azure-ad-role-assignments-via-the-graph-api-or-powershell

#region Authentication
try { Connect-MgGraph -Scopes RoleManagement.Read.Directory,Directory.Read.All} #RoleManagement.Read.Directory needed for PIM eligible assignments, Directory.Read.All for roles and "translating" GUIDs
catch { Write-Error $_ -ErrorAction Stop; return }
#endregion Authentication

if (!(Get-MgContext).Scopes.Contains("Directory.Read.All")) { Write-Error "The required permissions are missing, please re-consent!" -ErrorAction Stop }

#region Roles
Write-Verbose "Collecting role assignments..."
#Use the Get-MgRoleManagementDirectoryRoleAssignment cmdlet to collect a list of all role assignments.
$roles = Get-MgRoleManagementDirectoryRoleAssignment -All -ExpandProperty Principal #$expand=* is BROKEN
$roles1 = Get-MgRoleManagementDirectoryRoleAssignment -All -ExpandProperty roleDefinition #fix to also fetch the roleDefinition
foreach ($role in $roles) { Add-Member -InputObject $role -MemberType NoteProperty -Name roleDefinition1 -Value ($roles1 | ? {$_.id -eq $role.id}).roleDefinition } #and another fix needed as PowerShell populates empty roleDefinition property...

#Use the Get-MgRoleManagementDirectoryRoleEligibilitySchedule cmdlet to collect a list of all PIM eligible role assignments.
if ($IncludePIMEligibleAssignments) {
    Write-Verbose "Collecting PIM eligible role assignments..."

    $roles += (Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All -ExpandProperty * | select id,principalId,directoryScopeId,roleDefinitionId,status,principal,@{n="roleDefinition1";e={$_.roleDefinition}})
}

if (!$roles) { Write-Verbose "No valid role assignments found, verify the required permissions have been granted?"}

Write-Verbose "A total of $($roles.count) role assignments were found, of which $(($roles | ? {$_.directoryScopeId -eq "/"}).Count) are tenant-wide and $(($roles | ? {$_.directoryScopeId -ne "/"}).Count) are AU-scoped. $(($roles | ? {!$_.status}).Count) roles are permanently assigned, you might want to address that!"
#endregion Roles

#region Output
#prepare the script output
Write-Verbose "Preparing the output..."
$report = @()
foreach ($role in $roles) {
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
        "AssignmentType" = (&{if ($role.status -eq "Provisioned") {"Eligible"} else {"Permanent"}})
        "IsBuiltIn" = $role.roleDefinition1.isBuiltIn
        "RoleTemplate" = $role.roleDefinition1.templateId
    }
    $report += @([pscustomobject]$reportLine)
}
#endregion Output

#format and export
$report | sort PrincipalDisplayName #| Export-CSV -nti -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_AzureADRoleInventory.csv"
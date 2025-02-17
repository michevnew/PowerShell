#Requires -Version 3.0
#Make sure to fill in all the required variables before running the script
#Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Directory.Read.All (required)
#    RoleManagement.Read.Directory (optional, needed to retrieve PIM eligible role assignments)
#    PrivilegedEligibilitySchedule.Read.AzureADGroup (optional, needed to retrieve Privileged Access Group assignments)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/5958/reporting-on-ent…ts-including-pim

#Add IsPrivileged to the output, requires /beta currently

[CmdletBinding()] #Make sure we can use -Verbose
Param([switch]$IncludePIMEligibleAssignments, #Indicate whether to include PIM elibigle role assignments in the output.
      [switch]$IncludePAGAssignments #Indicate whether to include Privileged Access Group assignments in the output.
)

#region Authentication
#Get an Access token. Make sure to fill in all the variable values here. Or replace with your own preferred method to obtain token.
$tenantId = "tenant.onmicrosoft.com"
$uri = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'
$clientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$client_secret = "verylongstring"

$Scopes = New-Object System.Collections.Generic.List[string]
$Scope = "https://graph.microsoft.com/.default"
$Scopes.Add($Scope)

$body = @{
    grant_type = "client_credentials"
    client_id = $clientId
    client_secret = $client_secret
    scope = $Scopes
}

try {
    Write-Verbose "Obtaining token..."
    $res = Invoke-WebRequest -Method Post -Uri $uri -Body $body -ErrorAction Stop -Verbose:$false
    $token = ($res.Content | ConvertFrom-Json).access_token

    $authHeader = @{
       'Authorization'="Bearer $token"
    }}
catch { Write-Output "Failed to obtain token, aborting..." ; return }
Write-Verbose "Warning: we are not checking whether the token contains all the required scopes, make sure application with sufficient permissions has been used!"
#endregion Authentication

#region Roles
Write-Verbose "Collecting role assignments..."
#Use the /roleManagement/directory/roleAssignments endpoint to collect a list of all role assignments. We cannot expand multiple properties, so we do two passes here.
$roles = @()
$uri = 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?$expand=principal' #$expand=* is BROKEN

try {
    do {
        $result = Invoke-WebRequest -Uri $uri -Verbose:$false -ErrorAction Stop -Headers $authHeader
        $uri = $($result | ConvertFrom-Json).'@odata.nextLink'
        #If we are getting multiple pages, best add some delay to avoid throttling
        Start-Sleep -Milliseconds 200
        $roles += ($result | ConvertFrom-Json).Value
    } while ($uri)
}
catch { Write-Error "Unable to obtain role assignments, make sure the required permissions have been granted..."; return }

#fix to also fetch the roleDefinition
$uri = 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?$expand=roleDefinition' #$expand=* is BROKEN

$roles1 = @()
try {
    do {
        $result = Invoke-WebRequest -Uri $uri -Verbose:$false -ErrorAction Stop -Headers $authHeader
        $uri = $($result | ConvertFrom-Json).'@odata.nextLink'
        #If we are getting multiple pages, best add some delay to avoid throttling
        Start-Sleep -Milliseconds 200
        $roles1 += ($result | ConvertFrom-Json).Value
    } while ($uri)
}
catch { Write-Error "Unable to obtain role assignments, make sure the required permissions have been granted..."; return }

#and another fix needed as PowerShell populates empty roleDefinition property...
foreach ($role in $roles) { Add-Member -InputObject $role -MemberType NoteProperty -Name roleDefinition1 -Value ($roles1 | ? {$_.id -eq $role.id}).roleDefinition }

#process PIM eligible role assignments, do not end the script if we fail to collect them
if ($IncludePIMEligibleAssignments) {
    Write-Verbose "Collecting PIM eligible role assignments..."
    $uri = 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleEligibilitySchedules?$expand=roleDefinition,principal'

    try {
        do {
            $result = Invoke-WebRequest -Uri $uri -Verbose:$false -ErrorAction Stop -Headers $authHeader
            $uri = $($result | ConvertFrom-Json).'@odata.nextLink'
            #If we are getting multiple pages, best add some delay to avoid throttling
            Start-Sleep -Milliseconds 200
            $roles += (($result | ConvertFrom-Json).Value | select id,principalId,directoryScopeId,roleDefinitionId,status,principal,@{n="roleDefinition1";e={$_.roleDefinition}})
        } while ($uri)
    }
    catch { Write-Warning "Unable to obtain PIM eligible role assignments, make sure the required permissions have been granted..." }

    $roleactivations = @()
    #Collect all PIM activated role assignments.
    $uri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignmentScheduleInstances?`$filter=assignmentType eq 'Activated'"
    try {
        do {
            $result = Invoke-WebRequest -Uri $uri -Verbose:$false -ErrorAction Stop -Headers $authHeader
            $uri = $($result | ConvertFrom-Json).'@odata.nextLink'
            #If we are getting multiple pages, best add some delay to avoid throttling
            Start-Sleep -Milliseconds 200
            $roleactivations += ($result | ConvertFrom-Json).Value
        } while ($uri)
    }
    catch { Write-Warning "Unable to obtain PIM eligible role assignments, make sure the required permissions have been granted..." }

    #If an eligible role is assigned, it will appear as Permanent in the output of /roleManagement/directory/roleAssignments, so we need some clean up
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
    $Proles = $roles | ? {$_.Principal.'@odata.type' -eq '#microsoft.graph.group'} #not necessarily PIM-managed group (can be "old" PAG, in both cases the role-assignable flag should be set)
    if (!$Proles) { Write-Verbose "No role assignments with Group principal found, skipping PAG collection" }

    foreach ($role in $Proles) {
        Write-Verbose "Collecting Privileged Access Group members for $($role.PrincipalId) ..."
        #Get the list of permanent/active members, easily done via the /transitiveMembers endpoint (with the added benefit of expanding nested groups)
        $dMembers = @{};$dMembersId = @()
        $uri = "https://graph.microsoft.com/v1.0/groups/$($role.PrincipalId)/transitiveMembers?`$select=id,displayName,userPrincipalName"

        try {
            do {
                $result = Invoke-WebRequest -Uri $uri -Verbose:$false -ErrorAction Stop -Headers $authHeader
                $uri = $($result | ConvertFrom-Json).'@odata.nextLink'
                #If we are getting multiple pages, best add some delay to avoid throttling
                Start-Sleep -Milliseconds 200
                ($result | ConvertFrom-Json).Value | % { $dMembers[$_.id] = $_.userPrincipalName }
                $dMembersId += (($result | ConvertFrom-Json).Value | % { if ($_.userPrincipalName) {$_.UserPrincipalName} else {"$($_.displayName) ($($_.Id))"} })
            } while ($uri)
        }
        catch { Write-Verbose "No members found for PAG $($role.PrincipalId), skipping..." }
        $role | Add-Member -MemberType NoteProperty -Name "Active group members" -Value $dMembers
        $role | Add-Member -MemberType NoteProperty -Name "Active group members IDs" -Value ($dMembersId -join ";")

        #Get the list of eligible members, done via the /beta/identityGovernance/privilegedAccess/group/eligibilitySchedules endpoint. #NOT expanding groups here
        #If a member is both eligible and active, it will appear in both lists!
        $eMembers = @{};$eMembersId = @()
        $uri = "https://graph.microsoft.com/beta/identityGovernance/privilegedAccess/group/eligibilitySchedules?`$filter=groupId eq `'$($role.principalId)`'&`$expand=principal"

        try {
            do {
                $result = Invoke-WebRequest -Uri $uri -Verbose:$false -ErrorAction Stop -Headers $authHeader
                $uri = $($result | ConvertFrom-Json).'@odata.nextLink'
                #If we are getting multiple pages, best add some delay to avoid throttling
                Start-Sleep -Milliseconds 200
                ($result | ConvertFrom-Json).Value | % { $eMembers[$_.principal.id] = $_.principal.userPrincipalName }
                $eMembersId += (($result | ConvertFrom-Json).Value | % { if ($_.principal.userPrincipalName) {$_.principal.userPrincipalName} else {"$($_.principal.displayName) ($($_.principal.Id))"} })
            } while ($uri)
        }
        catch { Write-Warning "Unable to retrieve eligible members of the $($role.PrincipalId) group, make sure the application has been granted PrivilegedEligibilitySchedule.Read.AzureADGroup permissions!" }
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
        if ($role.principal.'@odata.type' -eq '#microsoft.graph.group') {
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
        "Principal" = switch ($role.principal.'@odata.type') {
            '#microsoft.graph.user' {$role.principal.userPrincipalName}
            '#microsoft.graph.servicePrincipal' {$role.principal.appId}
            '#microsoft.graph.group' {$role.principalid}
        }
        "PrincipalDisplayName" = $role.principal.displayName
        "PrincipalType" = $role.principal.'@odata.type'.Split(".")[-1]
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
        #"IsPrivileged" = $role.roleDefinition1.IsPrivileged
        #"AssignmentMode" = $role.roleDefinition1.assignmentMode
        #"RichDescription" = $role.roleDefinition1.richDescription
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
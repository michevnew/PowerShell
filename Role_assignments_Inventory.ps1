#Requires -Version 3.0
[CmdletBinding()] #Make sure we can use -Verbose
Param([switch]$IncludeRoleGroups,[switch]$IncludeUnassignedRoleGroups,[switch]$IncludeDelegatingAssingments)

#Simple function to check for existing Exchange Remote PowerShell session or establish a new one
function Check-Connectivity {
    #Make sure we are connected to Exchange Remote PowerShell
    Write-Verbose "Checking connectivity to Exchange Remote PowerShell..."
    if (!$session -or ($session.State -ne "Opened")) {
        try { $script:session = Get-PSSession -InstanceId (Get-AcceptedDomain | select -First 1).RunspaceId.Guid -ErrorAction Stop  }
        catch {
            try {
                #Failing to detect an active session, try connecting to ExO via Basic auth...
                $script:session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential (Get-Credential) -Authentication Basic -AllowRedirection -ErrorAction Stop 
                Import-PSSession $session -ErrorAction Stop | Out-Null 
            }  
            catch { Write-Error "No active Exchange Remote PowerShell session detected, please connect first. To connect to ExO: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx" -ErrorAction Stop }
        }
    }

    #As the function is called every once in a while, use it to trigger some artifical delay in order to prevent throttling
    Start-Sleep -Milliseconds 500
    return $true
}

#Enumerate all Role Groups that have no management role assignments
function getEmptyRoleGroups {
    $ERGoutput = @()
    $ERG = @(Get-RoleGroup | ? {!($_.RoleAssignments)})
    if ($ERG.Count) {
        $ERG | % {
            $ERGobj = New-Object psobject
            $ERGobj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $_.Id
            $ERGobj | Add-Member -MemberType NoteProperty -Name AssignmentType -Value "N/A"
            $ERGobj | Add-Member -MemberType NoteProperty -Name AssigneeName -Value $_.Id #($_.Members -join ",")
            $ERGobj | Add-Member -MemberType NoteProperty -Name Assignee -Value $_.ExchangeObjectId.Guid #($_.Members -join ",")
            $ERGobj | Add-Member -MemberType NoteProperty -Name AssigneeType -Value "RoleGroup"
            $ERGobj | Add-Member -MemberType NoteProperty -Name AssignedRoles -Value "N/A"
            $ERGoutput += $ERGobj 
        }
        return $ERGoutput
    }
    else { return }
}

#Find the user matching a given DisplayName. If multiple entries are returned, use the -RoleAssignee parameter to determine the correct one. If unique entry is found, return UPN, otherwise return DisplayName
function getUPN ($user,$role) {
    $UPN = @(Get-User $user -ErrorAction SilentlyContinue | ? {(Get-ManagementRoleAssignment -Role $role -RoleAssignee $_.SamAccountName -ErrorAction SilentlyContinue)})
    if ($UPN.Count -ne 1) { return $user }
    if ($UPN) { return $UPN.UserPrincipalName }
    else { return $user }
}

#Find the group matching a given DisplayName. If multiple entries are returned, use the -RoleAssignee parameter to determine the correct one. If unique entry is found, return the email address if present, or GUID. Otherwise return DisplayName
function getGroup ($group,$role) {
    $grp = @(Get-Group $group -ErrorAction SilentlyContinue | ? {(Get-ManagementRoleAssignment -Role $role -RoleAssignee $_.SamAccountName -ErrorAction SilentlyContinue)})
    if ($grp.Count -ne 1) { return $group }
    if ($grp) {
        if ($grp.WindowsEmailAddress.ToString()) { return $grp.WindowsEmailAddress.ToString() }
        else { return $grp.Guid.Guid.ToString() }
    }
    else { return $group }
}

#List all Role assignments in the organization, including Delegating ones if the corresponding parameter is invoked
function Get-RoleAssignmentsReport {

    Param
    (
    #Specify whether to include Role Group entries in the output
    [Switch]$IncludeRoleGroups,
    #Specify whether to include delegating Role assingments in the output
    [Switch]$IncludeDelegatingAssingments)

    $output = @()
    #If $IncludeDelegatingAssingments is specified, return the delegating assignments
    #-RoleAssigneeType filtering works for single values only, filter client-side
    if (!$IncludeDelegatingAssingments) { $RoleAssignments = @(Get-ManagementRoleAssignment -GetEffectiveUsers -Delegating:$false | ? {$_.RoleAssigneeType -notmatch "RoleAssignmentPolicy|PartnerLinkedRoleGroup"}) }
    else { $RoleAssignments = @(Get-ManagementRoleAssignment -Delegating:$true | ? {$_.RoleAssigneeName -ne "Organization Management"}) }

    #override for on-premises due to the -Delegating parameter issue
    if ($session.ComputerName -ne "outlook.office365.com") { $RoleAssignments = @(Get-ManagementRoleAssignment -GetEffectiveUsers | ? {$_.RoleAssigneeType -notmatch "RoleAssignmentPolicy|PartnerLinkedRoleGroup"}) }

    $PercentComplete = 0; $count = 1;
    foreach ($ra in $RoleAssignments) {
        #Progress message
        $ActivityMessage = "Processing role assignment $($ra.Name). Please wait..."
        $StatusMessage = ("Processing {0} of {1}: {2}" -f $count, @($RoleAssignments).count, $ra.Guid.Guid.ToString())
        $PercentComplete = ($count / @($RoleAssignments).count * 100)
        Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
        $count++

        #Since we are using the -GetEffectiveUsers parameter, the number of entries will be huge. And since for each entry we do a Get-User or Get-Group, add some anti-throttling controls via Check-Connectivity
        if ($count /50 -is [int]) {
            Start-Sleep -Seconds 1
        }
    
        #Process each Role assignment entry
        if ($ra.EffectiveUserName -eq "All Group Members" -and $ra.AssignmentMethod -eq "Direct") {
            #Only list the "parent" entry when it's not a Role Group or when -IncludeRoleGroups is $true
            if (($ra.RoleAssigneeType -ne "RoleGroup") -or $IncludeRoleGroups) { 
                $raobj = New-Object psobject
                $raobj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $ra.RoleAssigneeName
                $raobj | Add-Member -MemberType NoteProperty -Name AssignmentType -Value $ra.AssignmentMethod
                $raobj | Add-Member -MemberType NoteProperty -Name AssigneeName -Value $ra.RoleAssigneeName
                $raobj | Add-Member -MemberType NoteProperty -Name Assignee -Value (getGroup $ra.RoleAssignee $ra.Role)
                $raobj | Add-Member -MemberType NoteProperty -Name AssigneeType -Value $ra.RoleAssigneeType
                $raobj | Add-Member -MemberType NoteProperty -Name AssignedRoles -Value (&{If ($IncludeDelegatingAssingments) { "Delegating - " + $ra.Role } else {$ra.Role}})
                $output += $raobj            
            }
        }
        else {
            #User role assignments
            $raobj = New-Object psobject
            $raobj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $ra.RoleAssigneeName
            $raobj | Add-Member -MemberType NoteProperty -Name AssignmentType -Value $ra.AssignmentMethod
            $raobj | Add-Member -MemberType NoteProperty -Name AssigneeName -Value $ra.EffectiveUserName
            $raobj | Add-Member -MemberType NoteProperty -Name Assignee -Value (getUPN $ra.EffectiveUserName $ra.Role)
            $raobj | Add-Member -MemberType NoteProperty -Name AssigneeType -Value "User"
            $raobj | Add-Member -MemberType NoteProperty -Name AssignedRoles -Value (&{If ($IncludeDelegatingAssingments) { "Delegating - " + $ra.Role } else {$ra.Role}})
            $output += $raobj
        }
    }
    #return the output
    return $output
}

###########################
# MAIN SCRIPT STARTS HERE #
###########################

#Initialize the parameters
if (!$IncludeRoleGroups -and $IncludeUnassignedRoleGroups) {
    $IncludeUnassignedRoleGroups = $false
    Write-Verbose "The parameter -IncludeUnassignedRoleGroups can only be used when -IncludeRoleGroups is specified as well, ignoring..." 
}

#Check if we are connected to Exchange PowerShell
if (!(Check-Connectivity)) { return }

#Get role assignments
Write-Verbose "Processing Role Assignments..."
$output = Get-RoleAssignmentsReport -IncludeRoleGroups:$IncludeRoleGroups

#Get delegating role assignments
if ($IncludeDelegatingAssingments -and ($session.ComputerName -eq "outlook.office365.com")) { 
    Write-Verbose "Processing Delegating Role Assignments..."
    $output += @(Get-RoleAssignmentsReport -IncludeDelegatingAssingments -IncludeRoleGroups:$IncludeRoleGroups) 
}

#For completeness, return entries for Role Groups with no assignments
if ($IncludeRoleGroups -and $IncludeUnassignedRoleGroups) {
    Write-Verbose "Processing 'empty' Role Groups..."
    $output += @(getEmptyRoleGroups)   
}

#Dump the raw output to a CSV file
$output #| Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_RoleAssignments.csv" -NoTypeInformation -Encoding UTF8 -UseCulture

#Transform the output and return it to the console. Group assignments by individual user/group
$global:varRoleAssignments = $output | group Assignee | select @{n="DisplayName";e={($_.Group.AssigneeName | sort -Unique)}},@{n="Identifier";e={$_.Name}},@{n="ObjectType";e={($_.Group.AssigneeType | sort -Unique) -join ","}},@{n="AssignmentType";e={($_.Group.AssignmentType | sort -Unique) -join ","}},@{n="Roles";e={($_.Group.AssignedRoles | sort -Unique) -join ","}} | sort DisplayName
$global:varRoleAssignments | ft
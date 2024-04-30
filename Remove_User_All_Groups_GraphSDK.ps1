#Requires -Version 3.0
#Requires -Modules @{ ModuleName="Microsoft.Graph.Groups"; ModuleVersion="2.17.0" }
#Requires -Modules @{ ModuleName="Microsoft.Graph.Users"; ModuleVersion="2.17.0" }
#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.0.0" }

#The following permission are required for the script to work:
#    Directory.Read.All (to ensure best results with /memberOf and /ownedObjects)
#    Group.ReadWrite.All (for removal of group members/owners)
#    RoleManagement.ReadWrite.Directory (for removal of Directory roles) #NOT covered by Directory.ReadWrite.All
#    DelegatedPermissionGrant.ReadWrite.All (for removal of OAuth2PermissionGrants)
#    AdministrativeUnit.ReadWrite.All (for removal of members of Administrative Units)
#    Application.ReadWrite.All (for removal of application/service principal owners) #This one is a must if processing Apps, even Directory.ReadWrite.All on its own does NOT work
#    Exchange administrator (for processing Exchange objects)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/6062/remove-user-from-all-microsoft-365-groups-and-roles-and-more-via-the-graph-api-non-interactive

[CmdletBinding(SupportsShouldProcess)] #Make sure we can use -WhatIf and -Verbose
Param([Parameter(Position=0,Mandatory)][ValidateNotNullOrWhiteSpace()][Alias("Id")][String[]]$Identity, #The UPN or GUID of the user(s) to process
[ValidateNotNullOrEmpty()][string[]]$Exceptions, #Comma-separated list of group, role, AU, SP, app GUIDs to exclude from processing. GUIDs only! Up to 1000 values supported
[switch]$ProcessOwnership, #Whether to include Ownership assignments in the processing (/ownedObjects). Added as separate switch because of the Application.ReadWrite.All requirement. NO Exchange processing!
[ValidateNotNullOrEmpty()][string]$SubstituteOwner, #The UPN or GUID of the user to use as a substitute owner for groups where the user we are removing is the only owner
[switch]$ProcessExchangeGroups, #Whether to include Exchange Online groups in the processing. NO ownership processing!
[switch]$ProcessOauthGrants, #Whether to include OAuth2PermissionGrants in the processing.
[switch]$IncludeDirectoryRoles, #Whether to include Directory roles in the processing. When combined with the -ProcessExchangeGroups switch, will also process Exchange roles
[switch]$IncludeAdministrativeUnits, #Whether to include Administrative Units in the processing
[switch]$Quiet #Whether to suppress output to the console
)

#==========================================================================
# Helper functions
#==========================================================================

#Function to resolve Exceptions values, remove incomplete entries, remove duplicates, etc
#Needs Directory.Read.All
function Process-Exceptions {
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string[]]$Exceptions
    )

    #Remove entries that do not match a GUID regex.
    $EGUIDs = $Exceptions | Sort-Object -Unique | ? {$_.ToLower() -match "^[a-f0-9]{8}-([a-f0-9]{4}-){3}[a-f0-9]{12}$"}
    if (!$EGUIDs) { return }

    try {
        $result = Get-MgDirectoryObjectById -Ids $EGUIDs -Types @("group","directoryRole","administrativeUnit","application","servicePrincipal") -Verbose:$false -ErrorAction Stop
    }
    catch {
        Process-Error -ErrorMessage $_ -User $user -Group "N/A"
        #No need to terminate, we can continue without exceptions.
        if (!$Quiet) { Write-Warning "Unable to resolve exceptions" }
    }

    $EGUIDs = @($result.Id | Sort-Object -Unique)
    Write-Verbose "The following list of exceptions will be used: ""$(@($EGUIDs) -join ", ")"""
    return $EGUIDs
}

#THIS ONE NEEDS TO BE REWORKED!!!
#Function to handle errors
function Process-Error {
    param(
    [Parameter(Mandatory)]$ErrorMessage,
    [Parameter(Mandatory)]$User,
    [Parameter(Mandatory)]$Group
    )

    #Insufficient permissions granted to the service principal, terminate the script
    if (!$ErrorMessage.ErrorDetails.Message) { #ExO throws a 401 with no ErrorMessage... no way to differentiate token expiry from generirc permission-related issues or other errors
        if ($ErrorMessage.Exception.Message -match "Response status code does not indicate success: 401") { Write-Error "ERROR: Insufficient permissions to connect to Exchange Online. Verify correct permissions are assigned to the service principal!" -ErrorAction Stop }
    }
    if ($ErrorMessage.ErrorDetails.Message -match "InsufficientPermissionsException|Insufficient privileges to complete the operation|Authorization_RequestDenied|Authorization failed due to missing permission scope") { Write-Error "ERROR: Insufficient permissions to perform the removal operation. Verify correct permissions are assigned to the service principal!" -ErrorAction Stop }
    elseif ($ErrorMessage.ErrorDetails.Message -match "The role assigned to application") { Write-Error "ERROR: Insufficient permissions to connect to Exchange Online. Verify the admin role(s) assigned to the service principal!" -ErrorAction Stop } #ExO
    #Token has expired, renew it and retry the operation
    #ExO throws a 401 with no ErrorMessage... no way to differentiate from generirc permission-related issues
    elseif ($ErrorMessage.ErrorDetails.Message -match "Lifetime validation failed, the token is expired|Access token has expired") {
        Write-Warning "Access token has expired, renewing it..."

        Connect-MgGraph -NoWelcome
        if ($ProcessExchangeGroups) { Connect-ExchangeOnline -ShowBanner:$false -ShowProgress:$false -SkipLoadingFormatData -CommandName "Get-Group","Get-DistributionGroup","Remove-RoleGroupMember","Remove-DistributionGroupMember","Get-ManagementRoleAssignment","Remove-ManagementRoleAssignment" -Verbose:$false -ErrorAction Stop }

        if (!(Get-MgContext)) { Write-Error "Failed to renew token, aborting..." -ErrorAction Stop }
        if ($ProcessExchangeGroups -and !(Get-ConnectionInformation)) { Write-Error "Failed to renew token, aborting..." -ErrorAction Stop }
    }
    #The rest are non-terminal errors
    elseif ($ErrorMessage.ErrorDetails.Message -match "Cannot Update a mail-enabled security groups and or distribution list.") {
        Write-Warning "Group ""$Group"" is authored in Exchange Online, its membership cannot be managed by the Graph API..."
        Write-Verbose "HINT: Use the -ProcessExchangeGroups switch when running the script in order to remove it..."
    } #just in case the filter fails to catch an ExO group
    elseif ($ErrorMessage.ErrorDetails.Message -match "ManagementObjectNotFoundException|ADNoSuchObjectException|Couldn't find object") { Write-Warning "The specified object was not found, this should not happen..." }
    elseif ($ErrorMessage.ErrorDetails.Message -match "DynamicGroupMembershipChangeDeniedException|Membership for this group is managed automatically") { Write-Warning "Group ""$Group"" uses dynamic membership, adjust the membership filter instead..." }
    #Thrown when trying to remove a member from DynamicMembership group... gotta love the consistency
    elseif ($ErrorMessage.ErrorDetails.Message -match "Insufficient privileges to complete the operation") { Write-Warning "You cannot remove members of the ""$Group"" Dynamic group, adjust the membership filter instead..." }
    #This should NOT be a problem, as we use the Graph API for removal, but just in case...
    elseif ($ErrorMessage.ErrorDetails.Message -match "GroupOwnersCannotBeRemovedException|Only Members who are not owners") { Write-Warning "User object ""$user"" is Owner of the ""$Group"" group and cannot be removed..." }
    #Handle the case where the user is the only owner of the group
    elseif ($ErrorMessage.ErrorDetails.Message -match "MinGroupOwnersCriteriaBreachedException|the person you're removing is currently the only owner|GroupMemberRemoveException|The user is the only owner of the group|The group must have at least one owner") {
        if ($SubstituteOwner) {
            if (!$Quiet) { Write-Warning "User ""$user"" is the only Owner of the ""$Group"" group!" }
            Write-Verbose "Attempting to replace the owner of the group with the substitute owner..."
            return "TrySubstituteOwner"
        }
        else {
            Write-Warning "User ""$user"" is the only Owner of the ""$Group"" group and cannot be removed..."
            Write-Verbose "HINT: You can use the -SubstituteOwner parameter to specify a substitute owner for the group and the script will try to remediate such scenarios!"
        }
    }
    elseif ($ErrorMessage.ErrorDetails.Message -match "MemberNotFoundException") { Write-Warning "User ""$user"" is not a member of the group ""$Group"", this should not happen..." }
    elseif ($ErrorMessage.ErrorDetails.Message -match "Invalid object identifier|does not exist or one of its queried reference-property|Unsupported referenced-object resource identifier") { Write-Warning "Either the user or Group does not exist, or the user is not a member of the group. This should not happen..." }
    #attempting to add owner that already exists
    elseif ($ErrorMessage.ErrorDetails.Message -match "One or more added object references already exist") { Write-Warning "User ""$user"" is already an Owner of the group ""$Group""." }
    else { $ErrorMessage | fl * -Force; return } #catch-all for any unhandled errors
}

#Function to handle output, saves us some repetitive code
function Process-Output {
    param (
    [Parameter(Mandatory)][PSCustomObject]$Output,
    [Parameter(Mandatory)][string]$Message
    )

    Write-Verbose $Message

    #Resolve GUIDs to UPNs for prettier output
    if (($Output["User"] -notmatch '@') -and $GUIDs.ContainsValue($Output["User"])) { $Output["User"] = $GUIDs.GetEnumerator().Where({$_.Value -eq $Output["User"]}).Name }

    $global:out += $Output;
    if (!$Quiet -and !$WhatIfPreference) { $Output | select User, Group, ObjectType, Result } #Write output to the console unless the -Quiet parameter is used
}

#Needs GroupMember.Read.All, Directory.Read.All for best experience
function Get-Membership {
    param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$User
    )

    $MemberOf = @()
    try {
        $MemberOf = Get-MgUserMemberOf -UserId $User -All -Property id,displayName,groupTypes,securityEnabled,mailEnabled,onPremisesSyncEnabled,isAssignableToRole -Verbose:$false -ErrorAction Stop
    }
    catch {
        Process-Error -ErrorMessage $_ -User $user -Group "N/A"
        #If we ended up here, we encountered something unaccounted for, we should terminate
        Write-Error "Failed to fetch group membership for user $User, aborting..." -ErrorAction Stop
        return
    }

    $MemberOf
}

#Needs Directory.Read.All
function Get-Ownership {
    param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$User
    )

    $OwnerOf = @()
    try {
        $result = Get-MgUserOwnedObject -UserId $User -All -Property id,displayName,groupTypes,securityEnabled,mailEnabled,onPremisesSyncEnabled,isAssignableToRole -Verbose:$false -ErrorAction Stop
        #return only supported object types
        $OwnerOf = $result | ? {($_.AdditionalProperties.'@odata.type' -eq "#microsoft.graph.group") -or ($_.AdditionalProperties.'@odata.type' -eq "#microsoft.graph.servicePrincipal") -or ($_.AdditionalProperties.'@odata.type' -eq "#microsoft.graph.application")}
    }
    catch {
        Process-Error -ErrorMessage $_ -User $user -Group "N/A"
        #If we ended up here, we encountered something unaccounted for, we should terminate
        Write-Error "Failed to fetch group ownership for user $User, aborting..." -ErrorAction Stop
        return
    }

    $OwnerOf
}

#Needs DelegatedPermissionGrant.ReadWrite.All
function Process-OAuthGrants {
    [CmdletBinding(SupportsShouldProcess)]
    param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$User
    )

    $OAuthGrants = @()
    Write-Verbose "Processing oauth2PermissionGrants for user $user..."
    $uri = "https://graph.microsoft.com/v1.0/users/$user/oauth2PermissionGrants?`$select=id,clientId,principalId,consentType,resourceId,scope"
    try {
        $OAuthGrants = Get-MgUserOauth2PermissionGrant -UserId $User -All -Property id,clientId,principalId,consentType,resourceId,scope -Verbose:$false -ErrorAction Stop

        if ($OAuthGrants) {
            foreach ($Grant in $OAuthGrants) {
                Write-Verbose "Removing oauth2PermissionGrant $($Grant.id) for user $user..."
                if ($PSCmdlet.ShouldProcess("Grant $($Grant.Id) for user ""$user""")) {
                    $result = Remove-MgOauth2PermissionGrant -OAuth2PermissionGrantId $Grant.Id -Verbose:$false -ErrorAction Stop #suppress the output
                    Process-Output -Output @{"User" = $user;"Group" = "[$($Grant.resourceId)]:$($Grant.scope)";"ObjectType" = "OAuth2PermissionGrant";"Result" = "Success"} -Message "Successfully removed oauth2PermissionGrant $($Grant.id) for user $user."
                }
                else { Process-Output -Output @{"User" = $user;"Group" = "[$($Grant.resourceId)]:$($Grant.scope)";"ObjectType" = "OAuth2PermissionGrant";"Result" = "Skipped due to Confirm process"} -Message "Skipped removal of oauth2PermissionGrant $($Grant.id) for user $user." }
            }
        }
        else { Write-Verbose "No oauth2PermissionGrants found for user $user, skipping..." }
    }
    catch {
        Process-Error -ErrorMessage $_ -User $user -Group "N/A"
        Process-Output -Output @{"User" = $user;"Group" = "[$($Grant.resourceId)]:$($Grant.scope)";"ObjectType" = "OAuth2PermissionGrant";"Result" = "Failed"} -Message "Failed to remove oauth2PermissionGrant $($Grant.id) for user $user."
    }
}

#Needs Group.ReadWrite.All (or RoleManagement.ReadWrite.Directory for PAGs)
#Should be relevant only to Group objects, Application and ServicePrincipal objects can be ownerless.
function Set-SubstituteOwner {
    param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][PSCustomObject]$Group,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$SubstituteOwner
    )

    #If needed, add support for ServicePrincipal and Application objects
    #No validation should be needed, so moving on to the actual operation
    $body = @{
        "@odata.id" = "https://graph.microsoft.com/v1.0/users/$SubstituteOwner"
    }
    try {
        $result = New-MgGroupOwnerByRef -GroupId $Group.Id -BodyParameter ($body | ConvertTo-Json) -Verbose:$false -ErrorAction Stop #suppress the output
        Process-Output -Output @{"User" = "$SubstituteOwner";"Group" = "[Owner] $($Group.displayName)";"ObjectType" = "Group";"Result" = "Success (Owner add)"} -Message "Successfully added Owner $SubstituteOwner to Group ""$($Group.displayName)""."
    }
    catch {
        Process-Error -ErrorMessage $_ -User $SubstituteOwner -Group $Group.displayName
        Process-Output -Output @{"User" = "$SubstituteOwner";"Group" = "[Owner] $($Group.displayName)";"ObjectType" = "Group";"Result" = "Failed (Owner add)"} -Message "Failed to add Owner $SubstituteOwner to Group ""$($Group.displayName)""."
    }
}

# Needs RoleManagement.ReadWrite.Directory
function Process-ScopedRoles {
    [CmdletBinding(SupportsShouldProcess)]
    param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$User
    )

    $ScopedRoles = @()
    Write-Verbose "Processing scoped Directory role assignments for user $user..."
    try {
        $result = Get-MgRoleManagementDirectoryRoleAssignment -Filter "principalid eq `'$user`'"
        $ScopedRoles = $result | ? {$_.directoryScopeId -ne "/"}

        if ($ScopedRoles) {
            foreach ($Role in $ScopedRoles) {
                Write-Verbose "Removing scoped role assignment $($Role.id) for user $user..."
                if ($PSCmdlet.ShouldProcess("Scoped role assignment $($Role.Id) for user ""$user""")) {
                    $result = Remove-MgRoleManagementDirectoryRoleAssignment -UnifiedRoleAssignmentId $Role.Id -Verbose:$false -ErrorAction Stop #suppress the output
                    Process-Output -Output @{"User" = $user;"Group" = "[$($Role.directoryScopeId)]:$($Role.roleDefinitionId)";"ObjectType" = "Scoped Directory role assignment";"Result" = "Success"} -Message "Successfully removed scoped Directory role assignment $($Role.id) for user $user."
                }
                else { Process-Output -Output @{"User" = $user;"Group" = "[$($Role.directoryScopeId)]:$($Role.roleDefinitionId)";"ObjectType" = "Scoped Directory role assignment";"Result" = "Skipped due to Confirm process"} -Message "Skipped removal of scoped Directory role assignment $($Role.id) for user $user." }
            }
        }
        else { Write-Verbose "No scoped Directory role assignments found for user $user, skipping..." }
    }
    catch {
        Process-Error -ErrorMessage $_ -User $user -Group "N/A"
        #If we ended up here, we encountered something unaccounted for, we should terminate
        Write-Error "Failed to fetch scoped Directory role membershhip for user $User, aborting..." -ErrorAction Stop
        return
    }
}

# Needs RoleManagement.ReadWrite.Directory AND RoleEligibilitySchedule.ReadWrite.Directory
function Process-EligibleRoles {
    [CmdletBinding(SupportsShouldProcess)]
    param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$User
    )

    $EligibleRoles = @()
    Write-Verbose "Processing eligible Directory role assignments for user $user..."
    try {
        $EligibleRoles = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -Filter "principalId eq `'$user`' and memberType eq 'Direct'"

        if ($EligibleRoles) {
            foreach ($Role in $EligibleRoles) {
                Write-Verbose "Removing eligible role assignment $($Role.id) for user $user..."
                $body = @{
                    Action = "AdminRemove"
                    PrincipalId = $user
                    RoleDefinitionId = $Role.roleDefinitionId
                    DirectoryScopeId = $Role.directoryScopeId
                    MemberType = $Role.memberType
                    Justification = "Removed by script"
                }
                if ($PSCmdlet.ShouldProcess("Eligible role assignment $($Role.Id) for user ""$user""")) {
                    $result = New-MgRoleManagementDirectoryRoleEligibilityScheduleRequest -BodyParameter ($body | ConvertTo-Json) -Verbose:$false -ErrorAction Stop #suppress the output
                    Process-Output -Output @{"User" = $user;"Group" = "[$($Role.directoryScopeId)]:$($Role.roleDefinitionId)";"ObjectType" = "Eligible Directory role assignment";"Result" = "Success"} -Message "Successfully removed Eligible Directory role assignment $($Role.id) for user $user."
                }
                else { Process-Output -Output @{"User" = $user;"Group" = "[$($Role.directoryScopeId)]:$($Role.roleDefinitionId)";"ObjectType" = "Eligible Directory role assignment";"Result" = "Skipped due to Confirm process"} -Message "Skipped removal of Eligible Directory role assignment $($Role.id) for user $user." }
            }
        }
        else { Write-Verbose "No Eligible Directory role assignments found for user $user, skipping..." }
    }
    catch {
        Process-Error -ErrorMessage $_ -User $user -Group "N/A"
        #If we ended up here, we encountered something unaccounted for, we should terminate
        Write-Error "Failed to fetch Eligible Directory role membershhip for user $User, aborting..." -ErrorAction Stop
        return
    }
}

# Needs RoleManagement.ReadWrite.Directory
function Remove-RoleMembership {
    [CmdletBinding(SupportsShouldProcess)]
    param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][PSCustomObject[]]$Roles,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$User,
    [string[]]$ExceptionsList
    )

    foreach ($Role in $Roles) {
        #Skip Exception Roles
        if ($Role.id -in $ExceptionsList) {
            Process-Output -Output @{"User" = $user;"Group" = "$($Role.AdditionalProperties.displayName)";"ObjectType" = "Directory role";"Result" = "Skipped due to exception"} -Message "Role ""$($Role.AdditionalProperties.displayName)"" is in the exception list, skipping..."
            continue
        }

        #Do the removal
        Write-Verbose "Removing user $User from role ""$($Role.AdditionalProperties.displayName)""..."
        if ($PSCmdlet.ShouldProcess("User $User from role ""$($Role.AdditionalProperties.displayName)""")) {
            try {
                $result = Remove-MgDirectoryRoleMemberDirectoryObjectByRef -DirectoryRoleId $Role.Id -DirectoryObjectId $User -Verbose:$false -ErrorAction Stop #suppress the output
                Process-Output -Output @{"User" = $user;"Group" = "$($Role.AdditionalProperties.displayName)";"ObjectType" = "Directory role";"Result" = "Success"} -Message "Successfully removed user $User from role ""$($Role.AdditionalProperties.displayName)""."
            }
            catch {
                Process-Error -ErrorMessage $_ -Group $Role.AdditionalProperties.displayName -User $User
                Process-Output -Output @{"User" = $user;"Group" = "$($Role.AdditionalProperties.displayName)";"ObjectType" = "Directory role";"Result" = "Failed"} -Message "Failed to remove user $User from role ""$($Role.AdditionalProperties.displayName)""."
                continue
            }
        }
        else { Process-Output -Output @{"User" = $user;"Group" = "$($Role.AdditionalProperties.displayName)";"ObjectType" = "Directory role";"Result" = "Skipped due to Confirm process"} -Message "Skipping removal of user $user from role ""$($Role.AdditionalProperties.displayName)""." }
    }
}

#Needs AdministrativeUnit.ReadWrite.All
function Remove-AUMembership {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][PSCustomObject[]]$AUs,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$User,
        [string[]]$ExceptionsList
    )

    foreach ($AU in $AUs) {
        #Skip Exception AUs
        if ($AU.id -in $ExceptionsList) {
            Process-Output -Output @{"User" = $user;"Group" = "$($AU.AdditionalProperties.displayName)";"ObjectType" = "Administrative unit";"Result" = "Skipped due to exception"} -Message "Administrative Unit ""$($AU.AdditionalProperties.displayName)"" is in the exception list, skipping..."
            continue
        }

        #Skip Dynamic AUs
        if ($AU.membershipType -eq "Dynamic") {
            Process-Output -Output @{"User" = $user;"Group" = "$($AU.AdditionalProperties.displayName)";"ObjectType" = "Administrative unit";"Result" = "Skipped due to dynamic membership"} -Message "Administrative Unit ""$($AU.AdditionalProperties.displayName)"" is using dynamic membership, skipping..."
            continue
        }

        #Do the removal
        Write-Verbose "Removing user $User from Administrative Unit ""$($AU.AdditionalProperties.displayName)""..."
        if ($PSCmdlet.ShouldProcess("User $User from Administrative Unit ""$($AU.AdditionalProperties.displayName)""")) {
            try {
                $result = Remove-MgDirectoryAdministrativeUnitMemberDirectoryObjectByRef -AdministrativeUnitId $AU.Id -DirectoryObjectId $User -Verbose:$false -ErrorAction Stop #suppress the output
                Process-Output -Output @{"User" = $user;"Group" = "$($AU.AdditionalProperties.displayName)";"ObjectType" = "Administrative unit";"Result" = "Success"} -Message "Successfully removed user $User from Administrative Unit ""$($AU.AdditionalProperties.displayName)""."
            }
            catch {
                Process-Error -ErrorMessage $_ -Group $AU.AdditionalProperties.displayName -User $User
                Process-Output -Output @{"User" = $user;"Group" = "$($AU.AdditionalProperties.displayName)";"ObjectType" = "Administrative unit";"Result" = "Failed"} -Message "Failed to remove user $User from Administrative Unit ""$($AU.AdditionalProperties.displayName)""."
                continue
            }
        }
        else { Process-Output -Output @{"User" = $user;"Group" = "$($AU.AdditionalProperties.displayName)";"ObjectType" = "Administrative unit";"Result" = "Skipped due to Confirm process"} -Message "Skipping removal of user $user from Administrative Unit ""$($AU.AdditionalProperties.displayName)""." }
    }
}

#Needs Group.ReadWrite.All (or RoleManagement.ReadWrite.Directory for PAGs)
#Application.ReadWrite.All for application/service principal ownership
function Remove-Ownership {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][PSCustomObject[]]$Groups,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$User,
        [string]$SubstituteOwner,
        [switch]$PreventRecursion,
        [string[]]$ExceptionsList
    )

    #Process each group
    foreach ($Group in $Groups) {
        #Determine the endpoint based on the object type
        $cmdlet = switch ($($Group.AdditionalProperties.'@odata.type')) {
            "#microsoft.graph.group" { "Remove-MgGroupOwnerDirectoryObjectByRef -GroupId $($Group.Id)" }
            "#microsoft.graph.servicePrincipal" { "Remove-MgServicePrincipalOwnerDirectoryObjectByRef -ServicePrincipalId $($Group.Id)" }
            "#microsoft.graph.application" { "Remove-MgApplicationOwnerDirectoryObjectByRef -ApplicationId $($Group.Id)" }
            Default { return } #we terminate, as we've encountered something unaccounted for
        }
        $cmdlet += " -DirectoryObjectId $User -Verbose:`$false -ErrorAction Stop"

        #Skip Exception Groups
        if ($Group.id -in $ExceptionsList) {
            Process-Output -Output @{"User" = $user;"Group" = "[Owner] $($Group.AdditionalProperties.displayName)";"ObjectType" = $Group.AdditionalProperties.'@odata.type'.Split(".")[-1];"Result" = "Skipped due to exception"} -Message "Object ""$($Group.AdditionalProperties.displayName)"" is in the exception list, skipping..."
            continue
        }

        #Do the removal
        Write-Verbose "Removing Owner $User from object ""$($Group.AdditionalProperties.displayName)""..."
        if ($PSCmdlet.ShouldProcess("Owner $User from object ""$($Group.AdditionalProperties.displayName)""")) {
            try {
                $result = Invoke-Expression $cmdlet #suppress the output
                Process-Output -Output @{"User" = "$user";"Group" = "[Owner] $($Group.AdditionalProperties.displayName)";"ObjectType" = $Group.AdditionalProperties.'@odata.type'.Split(".")[-1];"Result" = "Success (Onwer remove)"} -Message "Successfully removed Owner $User from Object ""$($Group.AdditionalProperties.displayName)""."
            }
            catch {
                #Handle the case where the user is the only owner of the group
                if ((Process-Error -ErrorMessage $_ -Group $Group.AdditionalProperties.displayName -User $User) -eq "TrySubstituteOwner") {
                    #Detect recursion
                    if ($PreventRecursion) {
                        if (!$Quiet) { Write-Warning "We already attempted to substitute the owner for the Object ""$($Group.AdditionalProperties.displayName)"" and failed, skipping..." }
                        continue
                    }
                    #Try to replace the owner with the SubstituteOwner
                    try {
                        if (!$SubstituteOwner) { continue } #making double sure we have a value
                        Set-SubstituteOwner -Group $Group -SubstituteOwner $SubstituteOwner

                        #force reprocesing the removal operation (once!)
                        Start-Sleep -Seconds 1 #wait for the change to propagate
                        Remove-Ownership -PreventRecursion -Groups @($Group) -User $User -ExceptionsList $ExceptionsList -Confirm:$false #Skip the confirmation prompt as we already asked once!
                    }
                    catch { Process-Output -Output @{"User" = "$user";"Group" = "[Owner] $($Group.AdditionalProperties.displayName)";"ObjectType" = $Group.AdditionalProperties.'@odata.type'.Split(".")[-1];"Result" = "Failed (Owner remove)"} -Message "Failed to remove Owner $User from object ""$($Group.AdditionalPropertiesdisplayName)""." }
                }
                else { Process-Output -Output @{"User" = "$user";"Group" = "[Owner] $($Group.AdditionalProperties.displayName)";"ObjectType" = $Group.AdditionalProperties.'@odata.type'.Split(".")[-1];"Result" = "Failed (Owner remove)"} -Message "Failed to remove Owner $User from object ""$($Group.AdditionalProperties.displayName)""." }
                continue
            }
        }
        else { Process-Output -Output @{"User" = "$user";"Group" = "[Owner] $($Group.AdditionalProperties.displayName)";"ObjectType" = $Group.AdditionalProperties.'@odata.type'.Split(".")[-1];"Result" = "Skipped due to Confirm process"} -Message "Skipping removal of Owner $user from object ""$($Group.AdditionalProperties.displayName)""." }
    }
}

#Needs Group.ReadWrite.All (or RoleManagement.ReadWrite.Directory for PAGs)
function Remove-GroupMembership {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][PSCustomObject[]]$Groups,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][String]$User,
        [string[]]$ExceptionsList
    )

    foreach ($Group in $Groups) {
        #Quick fix to add the RecipientType property
        if ($Group.AdditionalProperties.mailEnabled -eq $true -and $Group.AdditionalProperties.groupTypes -notcontains "Unified") { $Group | Add-Member -MemberType NoteProperty -Name "RecipientType" -Value "Exchange Group" }
        else { $Group | Add-Member -MemberType NoteProperty -Name "RecipientType" -Value "Group" }

        #Skip Exception Groups
        if ($Group.id -in $ExceptionsList) {
            Process-Output -Output @{"User" = $user;"Group" = $Group.AdditionalProperties.displayName;"ObjectType" = $Group.RecipientType;"Result" = "Skipped due to exception"} -Message "Group ""$($Group.AdditionalProperties.displayName)"" is in the exception list, skipping..."
            continue
        }

        #Skip On-Prem Synced Groups
        if ($Group.AdditionalProperties.onPremisesSyncEnabled -eq $true) {
            Process-Output -Output @{"User" = $user;"Group" = $Group.AdditionalProperties.displayName;"ObjectType" = $Group.RecipientType;"Result" = "Skipped due to on-premises sync"} -Message "Group ""$($Group.AdditionalProperties.displayName)"" is synced from on-premises, skipping..."
            continue
        }

        #Skip Dynamic Groups
        if ($Group.AdditionalProperties.groupTypes -contains "DynamicMembership") {
            Process-Output -Output @{"User" = $user;"Group" = $Group.AdditionalProperties.displayName;"ObjectType" = $Group.RecipientType;"Result" = "Skipped due to dynamic membership"} -Message "Group ""$($Group.AdditionalProperties.displayName)"" is using dynamic membership, skipping..."
            continue
        }

        #Do the removal
        Write-Verbose "Removing user $User from group ""$($Group.AdditionalProperties.displayName)""..."
        if ($PSCmdlet.ShouldProcess("User $User from group ""$($Group.AdditionalProperties.displayName)""")) {
            #If Distribution Groups or Mail-enabled security group, use the Exchange methods
            if ($Group.AdditionalProperties.mailEnabled -eq $true -and $Group.AdditionalProperties.groupTypes -notcontains "DynamicMembership" -and $Group.AdditionalProperties.groupTypes -notcontains "Unified") {
                if (!$ProcessExchangeGroups) {
                    Process-Output -Output @{"User" = $user;"Group" = $Group.AdditionalProperties.displayName;"ObjectType" = $Group.RecipientType;"Result" = "Skipped due to Exchange group"} -Message "Group ""$($Group.AdditionalProperties.displayName)"" is authored in Exchange Online, please use the -ProcessExchangeGroups switch when running the script in order to remove it..."
                    continue
                }

                #Hack to get the group processed by Remove-ExchangeMembership
                $Group.RecipientType = "SecurityGroup"
                if (!$Group.displayName) { $Group | Add-Member -MemberType NoteProperty -Name "displayName" -Value $Group.AdditionalProperties.displayName }
                Remove-ExchangeMembership -Group $Group -User $User -ExceptionsList $ExceptionsList -Confirm:$false #already went through the confirmation process
            }
            #Otherwise, use the Graph API
            else {
                try {
                    Remove-MgGroupMemberDirectoryObjectByRef -GroupId $Group.Id -DirectoryObjectId $User -Verbose:$false -ErrorAction Stop
                    Process-Output -Output @{"User" = $user;"Group" = $Group.AdditionalProperties.displayName;"ObjectType" = $Group.RecipientType;"Result" = "Success"} -Message "Successfully removed user $User from group ""$($Group.AdditionalProperties.displayName)""."
                }
                catch {
                    Process-Error -ErrorMessage $_ -User $User -Group $Group.AdditionalProperties.displayName
                    Process-Output -Output @{"User" = $user;"Group" = $Group.AdditionalProperties.displayName;"ObjectType" = $Group.RecipientType;"Result" = "Failed"} -Message "Failed to remove user $User from group ""$($Group.AdditionalProperties.displayName)""."
                    continue
                }
            }
        }
        else { Process-Output -Output @{"User" = $user;"Group" = $Group.AdditionalProperties.displayName;"ObjectType" = $Group.RecipientType;"Result" = "Skipped due to Confirm process"} -Message "Skipping removal of user $User from group ""$($Group.AdditionalProperties.displayName)""." }
    }
}

#Needs Distribution Groups Management role (or Role Management if removing Role group membership)
function Remove-ExchangeMembership {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][PSCustomObject]$Group,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$User,
        [string[]]$ExceptionsList
    )

    #Resolve the entry first
    Write-Verbose "Resolving group entry ""$($Group.displayName)""..."
    $cmdlet = switch ($Group.RecipientType) {
        "RoleGroup" { "Get-Group" }
        "SecurityGroup" { "Get-DistributionGroup" }
        default { return } #we terminate, as we've encountered something unaccounted for
    }
    if (!$cmdlet) { Write-Verbose "Invalid group type ""$($Group.RecipientType)"" found, skipping..."; return }

    #If we passed the Group object from Graph, it should have id/ExternalDirectoryObjectId value, so we have our unique identifier
    if ($Group.id) { $cmdlet += " -Identity $($Group.id)" }
    #Else, filter by displayName
    else { $cmdlet += " -Filter ""Name -eq '$($Group.displayName)'""" }

    try {
        $result = Invoke-Expression $cmdlet -ErrorAction Stop -Verbose:$false
        if (!$result) { Write-Verbose "Group ""$($Group.displayName)"" not found, skipping..."; return }
        if ($result.count -gt 1) { Write-Verbose "Multiple groups matching the identifier $($Group.displayName) found, skipping..."; return }

        #Replace RecipientType with the real one
        $Group.RecipientType = $result.RecipientTypeDetails
        if (!$Group.ExchangeObjectId) { $Group | Add-Member -MemberType NoteProperty -Name "ExchangeObjectId" -Value $result.ExchangeObjectId.ToString() }
        if (!$Group.id) { $Group | Add-Member -MemberType NoteProperty -Name "Id" -Value $result.ExternalDirectoryObjectId }
        if (!$Group.onPremisesSyncEnabled -and !$Group.IsDirSynced) { $Group | Add-Member -MemberType NoteProperty -Name "IsDirSynced" -Value $result.IsDirSynced }
    }
    catch {
        Process-Error -ErrorMessage $_ -User $User -Group $Group.displayName
        Write-Verbose "Failed to resolve group ""$($Group.displayName)""..."
        return
    }

    #Mail-enabled security groups can be covered by exceptions, RoleGroups cannot (no ExternalDirectoryObjectId)
    if ($null -ne $Group.Id -and $Group.id -in $ExceptionsList) {
        Process-Output -Output @{"User" = $user;"Group" = $Group.displayName;"ObjectType" = $Group.RecipientType;"Result" = "Skipped due to exception"} -Message "Group ""$($Group.displayName)"" is in the exception list, skipping..."
        return
    }

    #Skip On-Prem Synced Groups
    if ($Group.onPremisesSyncEnabled -eq $true -or $Group.IsDirSynced -eq $true) {
        Process-Output -Output @{"User" = $user;"Group" = $Group.displayName;"ObjectType" = $Group.RecipientType;"Result" = "Skipped due to on-premises sync"} -Message "Group ""$($Group.displayName)"" is synced from on-premises, skipping..."
        return
    }

    #Skip already processed groups, this can happen if we have Exchagne group-based role assignments
    if ($script:processed[$($Group.ExchangeObjectId)]) {
        Write-Warning "We already tried to process ""$($Group.displayName)"" and $($script:processed[$($Group.ExchangeObjectId)]), skipping..."
        Process-Output -Output @{"User" = $user;"Group" = $Group.displayName;"ObjectType" = $Group.RecipientType;"Result" = "Skipped due to previous match"} -Message "Group ""$($Group.displayName)"" was already processed, skipping..."
        return
    }

    #Do the removal
    Write-Verbose "Removing user $User from group ""$($Group.displayName)""..."
    if ($PSCmdlet.ShouldProcess("User $User from group ""$($Group.displayName)""")) {
        $cmdlet = switch ($Group.RecipientType) {
            "RoleGroup" { "Remove-RoleGroupMember" }
            "SecurityGroup" { "Remove-DistributionGroupMember" }
            default { "Remove-DistributionGroupMember" }
        }
        $cmdlet += " -Identity $($Group.ExchangeObjectId) -Member $User -Confirm:`$false -BypassSecurityGroupManagerCheck"

        try {
            $result = Invoke-Expression $cmdlet -ErrorAction Stop -Verbose:$false #suppress the output
            Process-Output -Output @{"User" = $user;"Group" = $Group.displayName;"ObjectType" = $Group.RecipientType;"Result" = "Success"} -Message "Successfully removed user $User from group ""$($Group.displayName).""."
            $script:processed["$($Group.ExchangeObjectId)"] = "Succeeded" #Cannot use Id/ExternalDirectoryObjectId as RoleGroups do not have them populated
        }
        catch {
            Process-Error -ErrorMessage $_ -User $User -Group $Group.displayName
            Process-Output -Output @{"User" = $user;"Group" = $Group.displayName;"ObjectType" = $Group.RecipientType;"Result" = "Failed"} -Message "Failed to remove user $User from group ""$($Group.displayName)""."
            $script:processed["$($Group.ExchangeObjectId)"] = "Failed" #Cannot use Id/ExternalDirectoryObjectId as RoleGroups do not have them populated
        }
    }
    else { Process-Output -Output @{"User" = $user;"Group" = $Group.displayName;"ObjectType" = $Group.RecipientType;"Result" = "Skipped due to Confirm process"} -Message "Skipping removal of user $user from group ""$($Group.displayName)""." }
}

#Needs Role Management role in Exchange Online
function Remove-ExchangeRoleAssignments {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][PSCustomObject[]]$RoleAssignments) #No need to validate null, it just continues

    foreach ($RoleAssignment in $RoleAssignments) {
        Write-Verbose "Removing direct Management role assignment ""$RoleAssignment""..."
        if ($PSCmdlet.ShouldProcess("Management role assingnment ""$RoleAssignment""")) {
            try {
                $result = Remove-ManagementRoleAssignment -Identity $RoleAssignment -Confirm:$false -Verbose:$false -ErrorAction Stop #suppress the output
                Process-Output -Output @{"User" = $user.Name;"Group" = $RoleAssignment;"ObjectType" = "Exchange Role assignment";"Result" = "Success"} -Message "Successfully removed Management role assignment ""$RoleAssignment""."
            }
            catch {
                Process-Error -ErrorMessage $_ -User $user -Group $RoleAssignment
                Process-Output -Output @{"User" = $user.Name;"Group" = $RoleAssignment;"ObjectType" = "Exchange Role assignment";"Result" = "Failed"} -Message "Failed to remove Management role assignment ""$RoleAssignment""."
            }
        }
        else { Process-Output -Output @{"User" = $user.Name;"Group" = $RoleAssignment;"ObjectType" = "Exchange Role assignment";"Result" = "Skipped due to Confirm process"} -Message "Skipping removal of Management role assignment ""$RoleAssignment""." }
    }
}

#==========================================================================
# Main script
#==========================================================================

#Determine the required scopes, based on the parameters passed to the script
$RequiredScopes = switch ($PSBoundParameters.Keys) {
    "IncludeDirectoryRoles" { "Directory.ReadWrite.All", "RoleManagement.ReadWrite.Directory" } #Mandatory for Directory Role assignments
    "ProcessOwnership" { "Application.ReadWrite.All" } #Mandatory, NOT covered by Directory.ReadWrite.All
    #"ProcessOauthGrants" { "DelegatedPermissionGrant.ReadWrite.All" } #covered by Directory.ReadWrite.All
    #"IncludeAdministrativeUnits" { "AdministrativeUnit.ReadWrite.All" } #covered by Directory.ReadWrite.All
    Default { "Directory.Read.All", "Group.ReadWrite.All" }
}

#Connectivity bits
Connect-MgGraph -Scopes $RequiredScopes -NoWelcome -ErrorAction Stop
if ($ProcessExchangeGroups) { Connect-ExchangeOnline -ShowBanner:$false -ShowProgress:$false -SkipLoadingFormatData -CommandName "Get-Group","Get-DistributionGroup","Remove-RoleGroupMember","Remove-DistributionGroupMember","Get-ManagementRoleAssignment","Remove-ManagementRoleAssignment" -Verbose:$false -ErrorAction Stop }

if (!(Get-MgContext)) { Write-Error "Failed to connect to the Graph API, aborting..." -ErrorAction Stop }
if ($ProcessExchangeGroups) {
    if (!(Get-ConnectionInformation)) { Write-Error "Failed to connect to Exchange Online, aborting..." -ErrorAction Stop }
    if ((Get-MgContext).TenantId -ne (Get-ConnectionInformation).TenantId) { Write-Error "The Graph API and Exchange Online PowerShell connections are not using the same tenantID, aborting..." -ErrorAction Stop }
}

#Do we need these? Add them to #Requires instead?
if ($IncludeDirectoryRoles) {
    if (!(Get-Module Microsoft.Graph.Identity.Governance)) { Import-Module Microsoft.Graph.Identity.Governance -Verbose:$false -ErrorAction Stop }
    if ((Get-MgContext).Scopes -notcontains "Directory.ReadWrite.All" -and (Get-MgContext).Scopes -notcontains "RoleManagement.ReadWrite.Directory") {
        Write-Error "The current connection does not have the required permissions to process Directory Role assignments, aborting..." -ErrorAction Stop
    }
}
if ($ProcessOwnership) {
    if (!(Get-Module Microsoft.Graph.Applications)) { Import-Module Microsoft.Graph.Applications -Verbose:$false -ErrorAction Stop }
    if ((Get-MgContext).Scopes -notcontains "Application.ReadWrite.All") {
        Write-Error "The current connection does not have the required permissions to process ownership, aborting..." -ErrorAction Stop
    }
}
if ($IncludeAdministrativeUnits) {
    if (!(Get-Module Microsoft.Graph.Identity.DirectoryManagement)) { Import-Module Microsoft.Graph.Identity.DirectoryManagement -Verbose:$false -ErrorAction Stop }
    if (((Get-MgContext).Scopes -notcontains "AdministrativeUnit.ReadWrite.All") -and ((Get-MgContext).Scopes -notcontains "Directory.ReadWrite.All")) {
        Write-Error "The current connection does not have the required permissions to process Administrative Units, aborting..." -ErrorAction Stop
    }
}
if ($ProcessOauthGrants) {
    if (((Get-MgContext).Scopes -notcontains "DelegatedPermissionGrant.ReadWrite.All") -and ((Get-MgContext).Scopes -notcontains "Directory.ReadWrite.All")) {
        Write-Error "The current connection does not have the required permissions to process OAuth2PermissionGrants, aborting..." -ErrorAction Stop
    }
}
$global:out = @() #Change scope?

#As the script supports bulk processing, we need to resolve Identity value, remove incomplete entries, remove duplicates, etc
$GUIDs = @{};
foreach ($user in $Identity) {
    #Remove entries that do not match a GUID or UPN regex
    if ($user.ToLower() -notmatch "^[a-f0-9]{8}-([a-f0-9]{4}-){3}[a-f0-9]{12}$" -and $user.ToLower() -notmatch "^[a-z0-9_.+-]+@[a-z0-9-]+\.[a-z0-9-.]+$") {
        Write-Verbose "Invalid identifier $user, skipping..."
        continue
    }

    #Make sure a matching user object is found and return its GUID.
    try {
        $result = Get-MgUser -UserId $user -ErrorAction Stop -Verbose:$false
    }
    catch {
        Process-Error -ErrorMessage $_ -User $user -Group "N/A"
        Write-Verbose "User obejct with identifier $user not found, skipping..."
        continue
    }

    if (($result.count -gt 1) -or ($GUIDs[$user]) -or ($GUIDs.Values -eq $result.id)) { Write-Verbose "Multiple users matching the identifier $user found, skipping..."; continue }
    else { $GUIDs[$result.userPrincipalName] = $result.id }
}
if (!$GUIDs -or ($GUIDs.Count -eq 0)) { Write-Error "No matching users found for ""$Identity"", check the parameter values." -ErrorAction Stop; return }
Write-Verbose "The following list of users will be used: ""$($GUIDs.Keys -join ", ")"""

#Do the same for any exceptions
if ($Exceptions) { $EGUIDs = Process-Exceptions $Exceptions }
else { $EGUIDs = $null }

#Resolve the SubstituteOwner value if provided
if ($SubstituteOwner) {
    try {
        $result = Get-MgUser -UserId $SubstituteOwner -ErrorAction Stop -Verbose:$false
        $SubstituteOwner = $result.id
    }
    catch {
        Process-Error -ErrorMessage $_ -User $SubstituteOwner -Group "N/A"
        if (!$Quiet) { Write-Warning "Substitute owner object with identifier $SubstituteOwner not found, check the parameter value." }
        [AllowNull]$SubstituteOwner = $null #AllowEmptyString instead?
    }

    if ($GUIDs.ContainsValue($SubstituteOwner)) {
        if (!$Quiet) { Write-Warning "Substitute owner cannot be the same as the user being processed, skipping..." }
        [AllowNull]$SubstituteOwner = $null
    }
    else { Write-Verbose "Using Substitute Owner ""$SubstituteOwner""." }
}

#Process each user
foreach ($user in $GUIDs.GetEnumerator()) {
    Write-Verbose "Processing user ""$($user.Name)""..."
    Start-Sleep -Milliseconds 500 #Add some delay to avoid throttling...

    #Fetch the (direct) memberships for the user
    $memberOf = Get-Membership $user.value
    if (!$memberOf) { Write-Verbose "No membership returned for user $($user.Name), skipping..." }

    #Fetch objects owned by the user
    if ($ProcessOwnership) {
        $ownerOf = Get-Ownership $user.value
        if (!$ownerOf) { Write-Verbose "No ownership returned for user $($user.Name), skipping..." }
    }

    #Process oauth2PermissionGrants for the user
    if ($ProcessOauthGrants) {
        Process-OauthGrants -User $user.value
    }

    #Remove Directory role assignments
    if ($IncludeDirectoryRoles) {
        Write-Verbose "Processing Directory Role assignments for user $($User.Name)..."
        $memberOfRoles = $MemberOf | ? {$_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.directoryRole'}

        #Can exclude by roleID (NOT roleTemplateID) if needed
        if ($memberOfRoles) { Remove-RoleMembership -Roles $memberOfRoles -User $User.Value -ExceptionsList $EGUIDs }
        else { Write-Verbose "No Directory Role assignments found for user $($User.Name), skipping..." }

        #As /memberOf does not cover scoped Directory role assignments, we do it via /roleManagement/directory/roleAssignments
        #No exceptions as getByIds does not support role templates
        Process-ScopedRoles -User $user.value

        #Same for elibigle roles. No exceptions as getByIds does not support role templates
        Process-EligibleRoles -User $user.value

        #If processing Exchange groups, also remove Exchange Role assignments
        if ($ProcessExchangeGroups) {
            $script:processed = @{} #Track Exchange groups we've already processed. Do the same for AAD roles that map to Exchange ones?
            Write-Verbose "Processing Exchange Role assignments for user $($User.Name)..."
            #Get the list of Exchange Role assignments
            try {
                $memberOfExchangeRoles = Get-ManagementRoleAssignment -RoleAssignee $user.Name -Verbose:$false -ErrorAction Stop | ? {$_.RoleAssigneeType -ne "RoleAssignmentPolicy"}
            }
            catch {
                Process-Error -ErrorMessage $_ -User $User.Name -Group "N/A"
                Process-Output -Output @{"User" = $user.Name;"Group" = "N/A";"ObjectType" = "Role assignment";"Result" = "Failed"} -Message "Failed to remove Management role assignments for user $($User.Name), skipping..."
            }

            #Remove direct role assignments
            if ($DirectRoleAssignments = $memberOfExchangeRoles | ? {$_.RoleAssigneeType -eq "User" -and $_.AssignmentMethod -eq "Direct"}) {
                Remove-ExchangeRoleAssignments -RoleAssignments $DirectRoleAssignments.Name
            }
            else { Write-Verbose "No direct Exchange Role assignments found for user $($User.Name), skipping..." }

            #Remove group-based role assignments. Cannot use ExchangeObjectId here, only RoleAssigneeName :(
            #Can actually have multiple matching the same role and group... so add Name below?
            if ($GroupRoleAssignments = $memberOfExchangeRoles | ? {$_.RoleAssigneeType -eq "RoleGroup" -or $_.RoleAssigneeType -eq "SecurityGroup"} | select RoleAssigneeName,RoleAssigneeType -Unique) {
                foreach ($GRA in $GroupRoleAssignments) {
                    #Role can be assigned to MESG, in which case an exception can apply
                    Remove-ExchangeMembership -Group ([PSCustomObject]@{"displayName" = $GRA.RoleAssigneeName;"RecipientType" = $GRA.RoleAssigneeType}) -User $User.Name -ExceptionsList $EGUIDs
                }
            }
            else { Write-Verbose "No group-based Exchange Role assignments found for user $($User.Name), skipping..." }
        }
    }

    #Remove Administrative unit membership
    if ($IncludeAdministrativeUnits) {
        Write-Verbose "Processing Administrative Unit membership for user $($User.Name)..."
        $memberOfAUs = $MemberOf | ? {$_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.administrativeUnit'}

        if ($memberOfAUs) { Remove-AUMembership -AUs $memberOfAUs -User $User.Value -ExceptionsList $EGUIDs }
        else { Write-Verbose "No Administrative unit membership found for user $($User.Name), skipping..." }
    }

    #Remove ownership
    Write-Verbose "Processing ownership for user $($User.Name)..."
    if ($ownerOf) {
        if ($SubstituteOwner) { Remove-Ownership -Groups $ownerOf -User $User.Value -SubstituteOwner $SubstituteOwner -ExceptionsList $EGUIDs }
        else { Remove-Ownership -Groups $ownerOf -User $User.Value -ExceptionsList $EGUIDs }
    }
    else { Write-Verbose "No ownership found for user $($User.Name), skipping..." }

    #Remove Group membership
    Write-Verbose "Processing Group membership for user $($User.Name)..."
    $memberOfGroups = $MemberOf | ? {$_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group'}

    if ($memberOfGroups) { Remove-GroupMembership -Groups $memberOfGroups -User $User.Value -ExceptionsList $EGUIDs }
    else { Write-Verbose "No group membership found for user $($User.Name), skipping..." }
}
Write-Verbose "Processing complete, exiting..."

if ($out) {
    if (!$Quiet) { #Write output to the console unless the -Quiet parameter is used
        $out | select User, @{n="Object";e={$_.Group}},ObjectType, Result | Out-Default
    }
    #Export the results to a CSV file
    $out | select User, @{n="Object";e={$_.Group}},ObjectType, Result | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_UserRemovalInfo.csv" -NoTypeInformation -Encoding UTF8 -UseCulture -Confirm:$false -WhatIf:$false
    Write-Verbose "Results exported to ""$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_UserRemovalInfo.csv""."
}
else { Write-Verbose "Output is empty, skipping the export to CSV file..." }
Write-Verbose "Finish..."
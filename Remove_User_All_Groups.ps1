#Requires -Version 3.0
#Add switch to handle situations where the user is the only owner of a Group?
[CmdletBinding(SupportsShouldProcess)] #Make sure we can use -WhatIf and -Verbose
Param([ValidateNotNullOrEmpty()][Alias("UserToRemove")][String[]]$Identity,[switch]$IncludeAADSecurityGroups,[switch]$IncludeOffice365Groups)

function Check-Connectivity {
    [cmdletbinding()]
    [OutputType([bool])]
    param([switch]$IncludeAADSecurityGroups)

    #Make sure we are connected to Exchange Remote PowerShell
    Write-Verbose "Checking connectivity to Exchange Remote PowerShell..."
    if (!$session -or ($session.State -ne "Opened")) {
        try { $script:session = Get-PSSession -InstanceId (Get-AcceptedDomain | select -First 1).RunspaceId.Guid -ErrorAction Stop  }
        catch {
            try {
                $script:session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential (Get-Credential) -Authentication Basic -AllowRedirection -ErrorAction Stop
                Import-PSSession $session -ErrorAction Stop | Out-Null
            }
            catch { Write-Error "No active Exchange Remote PowerShell session detected, please connect first. To connect to ExO: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx" -ErrorAction Stop }
        }
    }

    #IF using the SG parameter
    if ($IncludeAADSecurityGroups) {
        Write-Verbose "Checking connectivity to Azure AD..."
        if (!(Get-Module AzureAD -ListAvailable -Verbose:$false | ? {($_.Version.Major -eq 2 -and $_.Version.Build -eq 0 -and $_.Version.Revision -gt 55) -or ($_.Version.Major -eq 2 -and $_.Version.Build -eq 1)})) { Write-Host -BackgroundColor Red "This script requires a recent version of the AzureAD PowerShell module. Download it here: https://www.powershellgallery.com/packages/AzureAD/"; return}
        try { Get-AzureADCurrentSessionInfo -ErrorAction Stop -WhatIf:$false -Verbose:$false | Out-Null }
        catch { try { Connect-AzureAD -WhatIf:$false -Verbose:$false -ErrorAction Stop | Out-Null } catch { return $false } }
    }

    return $true
}

function Remove-UserFromAllGroups {
<#
.Synopsis
   Removes user from all groups in Office 365
.DESCRIPTION
   The Remove-UserFromAllGroups function remove a given user, or a list of users, as members from any groups in the organization. Group types include Distribution Groups, Mail-Enabled Security Groups, Office 365 Groups. The command accepts pipeline input.
.PARAMETER Identity
    Identity the -Identity parameter to designate the list of users. Any valid Exchange user identifier can be specified. Multiple users can be specified in a comma-separated list or array, see examples below.
.PARAMETER IncludeAADSecurityGroups
    Specify whether to include Azure AD security groups. If this parameter is used, the script requires connectivity to Azure AD PowerShell.
.PARAMETER IncludeOffice365Groups
    Specify whether to include Office 365 (modern) groups.
.PARAMETER WhatIf
    The -WhatIf switch simulates the actions of the command. You can use this switch to view the changes that would occur without actually applying those changes.
.PARAMETER Verbose
    The -Verbose switch provides additional details on the cmdlet progress, it can be useful when troubleshooting issues.
.EXAMPLE
   Remove-UserFromAllGroups huku

   Removes the selected user from all distribution groups.
.EXAMPLE
   Remove-UserFromAllGroups HuKu -WhatIf -Verbose -IncludeAADSecurityGroups

   Removes the selected user from all distribution and security groups. Additional verbose output will be shown as the cmdlet execution progresses. No actual changes will be performed due to the -WhatIf switch being used.
.EXAMPLE
   Get-User | Remove-UserFromAllGroups

   The command accepts pipeline input. To manually pass multiple users, use the following format:

   C:\> "vasil","huku" | Remove-UserFromAllGroups
.INPUTS
   User object
.OUTPUTS
   None
#>

    [CmdletBinding(SupportsShouldProcess=$true)]

    Param
    (
    <#The Identity parameter specifies the identity of the user object.

This parameter accepts the following values:
* Alias: JPhillips
* Canonical DN: Atlanta.Corp.Contoso.Com/Users/JPhillips
* Display Name: Jeff Phillips
* Distinguished Name (DN): CN=JPhillips,CN=Users,DC=Atlanta,DC=Corp,DC=contoso,DC=com
* Domain\Account: Atlanta\JPhillips
* GUID: fb456636-fe7d-4d58-9d15-5af57d0354c2
* Immutable ID: fb456636-fe7d-4d58-9d15-5af57d0354c2@contoso.com
* Legacy Exchange DN: /o=Contoso/ou=AdministrativeGroup/cn=Recipients/cn=JPhillips
* SMTP Address: Jeff.Phillips@contoso.com
* User Principal Name: JPhillips@contoso.com
        #>
        [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true, ValueFromRemainingArguments=$false)]
        [ValidateNotNullOrEmpty()][Alias("UserToRemove")][String[]]$Identity,
        [switch]$IncludeAADSecurityGroups,
        [switch]$IncludeOffice365Groups)

    Begin {
        #Check if we are connected to Exchange Online PowerShell and if needed, to Azure AD...
        if (Check-Connectivity -IncludeAADSecurityGroups:$IncludeAADSecurityGroups) { Write-Verbose "Parsing the Identity parameter..." }
        else { Write-Host "ERROR: Connectivity test failed, exiting the script..." -ForegroundColor Red; continue }
    }

    Process {
        #Needed to handle pipeline input
        $GUIDs = @{}

        foreach ($us in $Identity) {
            Start-Sleep -Milliseconds 80 #Add some delay to avoid throttling...
            #Make sure a matching security principal object is found and return its UPN
            $GUID = Invoke-Command -Session $session -ScriptBlock { Get-User $using:us | Select-Object DistinguishedName,ExternalDirectoryObjectId } -ErrorAction SilentlyContinue
            if (!$GUID) { Write-Verbose "Security principal with identifier $us not found, skipping..."; continue }
            elseif (($GUID.count -gt 1) -or ($GUIDs[$us]) -or ($GUIDs.ContainsValue($GUID))) { Write-Verbose "Multiple users matching the identifier $us found, skipping..."; continue }
            else { $GUIDs[$us] = $GUID | Select-Object DistinguishedName,ExternalDirectoryObjectId }
        }
        if (!$GUIDs -or ($GUIDs.Count -eq 0)) { Write-Host "ERROR: No matching users found for ""$Identity"", check the parameter values." -ForegroundColor Red; return } #When in Process block, use return instead of continue
        Write-Verbose "The following list of users will be used: ""$($GUIDs.Values.DistinguishedName -join ", ")"""

        #Needed to handle array values for the Identity parameter
        foreach ($user in $GUIDs.GetEnumerator()) {
            Write-Verbose "Processing user ""$($user.Name)""..."
            Start-Sleep -Milliseconds 80 #Add some delay to avoid throttling...

            #Handle Exchange groups
            Write-Verbose "Obtaining group list for user ""$($user.Name)""..."
            if ($IncludeOffice365Groups) { $GroupTypes = @("GroupMailbox","MailUniversalDistributionGroup","MailUniversalSecurityGroup") }
            else { $GroupTypes = @("MailUniversalDistributionGroup","MailUniversalSecurityGroup") }

            $Groups = Invoke-Command -Session $session -ScriptBlock { Get-Recipient -Filter "Members -eq '$($using:user.Value.DistinguishedName)'" -RecipientTypeDetails $Using:GroupTypes | Select-Object DisplayName,ExternalDirectoryObjectId,RecipientTypeDetails } -ErrorAction SilentlyContinue -HideComputerName
            if (!$Groups) { Write-Verbose "No matching groups found for ""$($user.Name)"", skipping..." }
            else { Write-Verbose "User ""$($user.Name)"" is a member of $(($Groups | measure).count) group(s)." }

            #cycle over each Group
            foreach ($Group in $Groups) {
                Write-Verbose "Removing user ""$($user.Name)"" from group ""$($Group.DisplayName)"""
                if ($Group.RecipientTypeDetails.Value -eq "GroupMailbox") {
                    try { Invoke-Command -Session $session -ScriptBlock { Remove-UnifiedGroupLinks -Identity $using:Group.ExternalDirectoryObjectId -Links $using:user.Value.DistinguishedName -LinkType Member -Confirm:$false -WhatIf:$using:WhatIfPreference } -ErrorAction Stop -HideComputerName }
                    catch [System.Management.Automation.RemoteException] {
                        #Some exceptions return the same category.reason RecipientTaskException. Using "exception" string match instead
                        if ($_.CategoryInfo.Reason -eq "ManagementObjectNotFoundException") { Write-Host "ERROR: The specified object not found, this should not happen..." -ForegroundColor Red }
                        #Seems they've updated the cmdlets to have unique error codes now, so account for that
                        elseif ($_.CategoryInfo.Reason -eq "RecipientTaskException" -and $_.Exception -match "Couldn't find object") { Write-Host "ERROR: User object ""$($user.Name)"" not found, this should not happen..." -ForegroundColor Red }
                        elseif ($_.CategoryInfo.Reason -eq "GroupOwnersCannotBeRemovedException" -or ($_.CategoryInfo.Reason -eq "RecipientTaskException" -and $_.Exception -match "Only Members who are not owners")) { Write-Host "ERROR: User object ""$($user.Name)"" is Owner of the ""$($Group.DisplayName)"" group and cannot be removed..." -ForegroundColor Red }
                        elseif ($_.CategoryInfo.Reason -eq "MinGroupOwnersCriteriaBreachedException" -or ($_.CategoryInfo.Reason -eq "RecipientTaskException" -and $_.Exception -match "the person you're removing is currently the only owner")) { Write-Host "ERROR: User object ""$($user.Name)"" is the only Owner of the ""$($Group.DisplayName)"" group and cannot be removed..." -ForegroundColor Red }
                        #no error is thrown if trying to remove a user that is not a member
                        else {$_ | fl * -Force; continue} #catch-all for any unhandled errors
                    }
                    catch {$_ | fl * -Force; continue} #catch-all for any unhandled errors
                }
                else {
                    try { Invoke-Command -Session $session -ScriptBlock { Remove-DistributionGroupMember -Identity $using:Group.ExternalDirectoryObjectId -Member $using:user.Value.DistinguishedName -BypassSecurityGroupManagerCheck -Confirm:$false -WhatIf:$using:WhatIfPreference -ErrorAction Stop } }
                    catch [System.Management.Automation.RemoteException] {
                        if ($_.CategoryInfo.Reason -eq "ManagementObjectNotFoundException") { Write-Host "ERROR: The specified object not found, this should not happen..." -ForegroundColor Red }
                        elseif ($_.CategoryInfo.Reason -eq "MemberNotFoundException") { Write-Host "ERROR: User ""$($user.Name)"" is not a member of the ""$($Group.DisplayName)"" group..." -ForegroundColor Red }
                        else {$_ | fl * -Force; continue} #catch-all for any unhandled errors
                    }
                    catch {$_ | fl * -Force; continue} #catch-all for any unhandled errors
                }
            }

            #Handle Azure AD security groups
            if ($IncludeAADSecurityGroups) {
                Write-Verbose "Obtaining security group list for user ""$($user.Name)""..."
                $GroupsAD = Get-AzureADUserMembership -ObjectId $user.Value.ExternalDirectoryObjectId -All $true | ? {$_.ObjectType -eq "Group" -and $_.SecurityEnabled -eq $true -and $_.MailEnabled -eq $false}

                if (!$GroupsAD) { Write-Verbose "No matching security groups found for ""$($user.Name)"", skipping..." }
                else { Write-Verbose "User ""$($user.Name)"" is a member of $(($GroupsAD | measure).count) security group(s)." }

                #cycle over each Group
                foreach ($groupAD in $GroupsAD) {
                    Write-Verbose "Removing user ""$($user.Name)"" from group ""$($GroupAD.DisplayName)"""
                    if (!$WhatIfPreference) {
                        try { Remove-AzureADGroupMember -ObjectId $GroupAD.ObjectId -MemberId $user.Value.ExternalDirectoryObjectId -ErrorAction Stop }
                        catch [Microsoft.Open.AzureAD16.Client.ApiException] {
                            if ($_.Exception.Message -match ".*Insufficient privileges to complete the operation") { Write-Host "ERROR: You cannot remove members of the ""$($groupAD.DisplayName)"" Dynamic group, adjust the membership filter instead..." -ForegroundColor Red }
                            elseif ($_.Exception.Message -match ".*Invalid object identifier") { Write-Host "ERROR: Group ""$($groupAD.DisplayName)"" not found, this should not happen..." -ForegroundColor Red }
                            elseif ($_.Exception.Message -match ".*Unsupported referenced-object resource identifier") { Write-Host "ERROR: User ""$($user.Name)"" not found, this should not happen..." -ForegroundColor Red }
                            elseif ($_.Exception.Message -match ".*does not exist or one of its queried reference-property") { Write-Host "ERROR: User ""$($user.Name)"" is not a member of the ""$($groupAD.DisplayName)"" group..." -ForegroundColor Red }
                            else {$_ | fl * -Force; continue} #catch-all for any unhandled errors
                        }
                        catch {$_ | fl * -Force; continue} #catch-all for any unhandled errors
                    }
                    else { Write-Host "WARNING: The Azure AD module cmdlets do not support the use of -WhatIf parameter, action was skipped..." }
            }}
        }}
}


#Invoke the Remove-MailboxFolderPermissionsRecursive function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
if ($PSBoundParameters.Count -ne 0) { Remove-UserFromAllGroups @PSBoundParameters }
else { Write-Host "INFO: The script was run without parameters, consider dot-sourcing it instead." -ForegroundColor Cyan ; return }
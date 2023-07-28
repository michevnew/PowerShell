#Requires -Version 3.0
[CmdletBinding(SupportsShouldProcess)] #Make sure we can use -WhatIf and -Verbose
Param([ValidateNotNullOrEmpty()][Alias("UserToRemove")][String[]]$Identity,[switch]$IncludeSharedMailboxes,[switch]$IncludeResourceMailboxes)

function Check-Connectivity {
    [cmdletbinding()]
    [OutputType([bool])]
    param()

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

    return $true
}

function Remove-UserMBPermissions {
<#
.Synopsis
   Removes user's permissions across all mailboxes of the selected type(s)
.DESCRIPTION
   The Remove-UserMBPermissions function remove mailbox permissions for a given user, or a list of users, from all mailboxes in the organization. Mailbox types include User mailboxes, Shared mailboxes, Resource mailboxes. The command accepts pipeline input.
.PARAMETER Identity
    Identity the -Identity parameter to designate the list of users. Any valid Exchange user identifier can be specified. Multiple users can be specified in a comma-separated list or array, see examples below.
.PARAMETER IncludeSharedMailboxes
    Specify whether to include Shared mailboxes.
.PARAMETER IncludeResourceMailboxes
    Specify whether to include room and equipment mailboxes
.PARAMETER WhatIf
    The -WhatIf switch simulates the actions of the command. You can use this switch to view the changes that would occur without actually applying those changes.
.PARAMETER Verbose
    The -Verbose switch provides additional details on the cmdlet progress, it can be useful when troubleshooting issues.
.EXAMPLE
   Remove-UserMBPermissions huku

   Removes permissions for the specified user from all user mailboxes.
.EXAMPLE
   Remove-UserMBPermissions HuKu -WhatIf -Verbose -IncludeSharedMailboxes

   Removes permissions for the specified user from all user and shared mailboxes. Additional verbose output will be shown as the cmdlet execution progresses. No actual changes will be performed due to the -WhatIf switch being used.
.EXAMPLE
   Get-User | Remove-UserMBPermissions

   The command accepts pipeline input. To manually pass multiple users, use the following format:

   C:\> "vasil","huku" | Remove-UserMBPermissions
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
    [switch]$IncludeSharedMailboxes,
    [switch]$IncludeResourceMailboxes)

#region BEGIN
    #Make sure we are connected to Exchange Remote PowerShell
    if (Check-Connectivity) { Write-Verbose "Connected to Exchange Remote PowerShell, processing..." }
    else { Write-Host "ERROR: Connectivity test failed, exiting the script..." -ForegroundColor Red; continue }

    #Initialize the variable used to designate recipient types, based on the script parameters
    $included = @("UserMailbox")
    if($IncludeSharedMailboxes) { $included += "SharedMailbox"}
    if($IncludeResourceMailboxes) { $included += "RoomMailbox"; $included += "EquipmentMailbox"}

    #Prepare the list of users (security principals)
    Write-Verbose "Parsing the Identity parameter..."
    $GUIDs = @{}
    foreach ($us in $Identity) {
        Start-Sleep -Milliseconds 80 #Add some delay to avoid throttling...
        #Make sure a matching security principal object is found and return its UPN
        $GUID = Invoke-Command -Session $session -ScriptBlock { Get-User $using:us | Select-Object UserPrincipalName,Sid } -ErrorAction SilentlyContinue
        if (!$GUID) { Write-Verbose "Security principal with identifier $us not found, skipping..."; continue }
        elseif (($GUID.count -gt 1) -or ($GUIDs[$us]) -or ($GUIDs.ContainsValue($GUID))) { Write-Verbose "Multiple users matching the identifier $us found, skipping..."; continue }
        else { $GUIDs[$us] = $GUID | Select-Object UserPrincipalName,Sid }
    }
    if (!$GUIDs -or ($GUIDs.Count -eq 0)) { Write-Host "ERROR: No matching users found for ""$Identity"", check the parameter values." -ForegroundColor Red; return }
    Write-Verbose "The following list of users will be used: ""$($GUIDs.Values.UserPrincipalName -join ", ")"""

    #If only a handful of users, do it the stupid way. If more than say 5, call Get-MailboxPermissionInventory!
    if ($GUIDs.Count -ge 5) {
        Write-Verbose "More than 4 users to be processed, obtaining full mailbox permission inventory..."
        try {
            $mailboxes = .\Mailbox_Permissions_inventory.ps1 -IncludeUserMailboxes -IncludeSharedMailboxes:$IncludeSharedMailboxes -IncludeRoomMailboxes:$IncludeResourceMailboxes -Verbose:$VerbosePreference
            Write-Verbose "Obtained total of $($mailboxes.count) permission entries."
        }
        catch {
            Write-Error $_ -ErrorAction Continue
            Write-Verbose "Failed to obtain full mailbox permission inventory, using the stupid method instead..."
            $mailboxes = $null
    }}
#endregion

#region PROCESS
    $out = @()

    #Needed to handle array values for the Identity parameter
    foreach ($user in $GUIDs.GetEnumerator()) {
        Write-Verbose "Processing user ""$($user.Name)""..."
        Start-Sleep -Milliseconds 80 #Add some delay to avoid throttling...

        if (!$mailboxes -or $mailboxes.count -eq 0) {
            #Remove permissions the stupid way
            Write-Verbose "Removing mailbox permissions for user ""$($user.Name)""..."
            Invoke-Command -Session $session -ScriptBlock { Get-Mailbox -RecipientTypeDetails $Using:included -ResultSize Unlimited | Remove-MailboxPermission -User $using:user.Value.UserPrincipalName -AccessRights FullAccess -Confirm:$false -WhatIf:$using:WhatIfPreference } -ErrorAction Continue #-HideComputerName
        }

        else {
            #As we are using the full mailbox permission inventory, filter out only the entries relevant to the current user
            $umailboxes = $mailboxes | ? {$_.User -eq $user.Value.UserPrincipalName -or $_.User.SecurityIdentifier.Value -eq $user.Value.Sid} #add Sid to cater for on-premises installs
            if (!$umailboxes -or $umailboxes.count -eq 0) { Write-Verbose "No matching permissions found for $($user.Name), skipping..." ; continue }

            #cycle over each Mailbox
            foreach ($mailbox in $umailboxes) {
                Write-Verbose "Removing permissions for user ""$($user.Name)"" from mailbox ""$($mailbox.'Mailbox address'.Address)"""
                try {
                    Invoke-Command -Session $session -ScriptBlock { Remove-MailboxPermission -Identity $using:mailbox.'Mailbox address'.Address -User $using:user.Value.UserPrincipalName -AccessRights FullAccess -Confirm:$false -WhatIf:$using:WhatIfPreference -ErrorAction Stop }
                    $outtemp = New-Object psobject -Property ([ordered]@{"Mailbox" = $mailbox.'Mailbox address'.Address;"AccessLevel" = "Full Access";"User" = $user.Value.UserPrincipalName})
                    $out += $outtemp; if (!$WhatIfPreference) { $outtemp } #Write output to the console unless we are using -WhatIf
                }
                catch [System.Management.Automation.RemoteException] {
                    if ($_.CategoryInfo.Reason -eq "ManagementObjectNotFoundException") { Write-Host "ERROR: The specified object not found, this should not happen..." -ForegroundColor Red }
                    else {$_ | fl * -Force; continue} #catch-all for any unhandled errors
                }
                catch {$_ | fl * -Force; continue} #catch-all for any unhandled errors
        }}}
#endregion

    if ($out) {
        Write-Verbose "Exporting results to the CSV file..."
        $out | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxPermissionsRemoved.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
        if (!$WhatIfPreference) { return $out | Out-Default } #Write output to the console unless the -WhatIf parameter is used
        }
    else { Write-Verbose "Output is empty, skipping the export to CSV file..." }
    Write-Verbose "Finish..."
}

#Invoke the Remove-UserMBPermissions function and pass the command line parameters.
if (($PSBoundParameters | measure).count) { Remove-UserMBPermissions @PSBoundParameters }
else { Write-Host "INFO: The script was run without parameters, consider dot-sourcing it instead." -ForegroundColor Cyan }
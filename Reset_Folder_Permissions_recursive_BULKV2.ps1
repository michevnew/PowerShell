#Requires -Version 3.0
#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.0.0" }

[CmdletBinding(SupportsShouldProcess)] #Make sure we can use -WhatIf and -Verbose
Param([switch]$Quiet,[ValidateNotNullOrEmpty()][Alias("Identity")][String[]]$Mailbox,[switch]$ResetDefaultLevel,[ValidateNotNullOrEmpty()][String[]]$ParentFolderPath)

#For details on what the script does and how to run it, check: hhttps://www.michev.info/blog/post/5763/managing-mailbox-folder-permissions-in-bulk-in-microsoft-365

#Include these folder types by default
$includedfolders = @("Root", "Inbox", "Calendar", "Contacts", "DeletedItems", "Drafts", "JunkEmail", "Journal", "Notes", "Outbox", "SentItems", "Tasks", "CommunicatorHistory", "Clutter", "Archive", "Scheduled")
#$includedfolders = @("Root","Inbox","Calendar", "Contacts", "DeletedItems", "SentItems", "Tasks") #Trimmed down list of default folders

#Exclude additional Non-default folders created by Outlook or other mail programs. Folder NAMES, not types! So make sure to include translations too!
#Exclude SearchDiscoveryHoldsFolder and SearchDiscoveryHoldsUnindexedItemFolder as they're not marked as default folders
$excludedfolders = @("News Feed", "Quick Step Settings", "Social Activity Notifications", "Suggested Contacts", "SearchDiscoveryHoldsUnindexedItemFolder", "SearchDiscoveryHoldsFolder", "Calendar Logging", "EventCheckPoints")

function Check-Connectivity {
    [cmdletbinding()]
    [OutputType([bool])]
    param()

    #Make sure we are connected to Exchange Remote PowerShell
    Write-Verbose "Checking connectivity to Exchange Remote PowerShell..."

    #Check via Get-ConnectionInformation first
    if (Get-ConnectionInformation) { return $true }

    #Double-check and try to eastablish a session
    try { Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null }
    catch {
        try { Connect-ExchangeOnline -CommandName Get-EXOMailbox, Get-MailboxFolderPermission, Remove-MailboxFolderPermission, Set-MailboxFolderPermission, Get-ExOMailboxFolderStatistics -SkipLoadingFormatData -ShowBanner:$false -Verbose:$false } #custom for this script
        catch { Write-Error "No active Exchange Online session detected. To connect to ExO: https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps"; return $false }
    }

    return $true
}

function ReturnFolderList {
<#
.Synopsis
    Enumerates all user-accessible folders for the mailbox
.DESCRIPTION
    The ReturnFolderList cmdlet enumerates the folders for the given mailbox. To adjust the list of folders, add to the $includedfolders or $excludedfolders array, respectively.
.PARAMETER SMTPAddress
    Use the -SMTPAddress parameter to designate the mailbox where the desired folders reside
.PARAMETER ParentFolderPath
    Use the -ParentFolderPath to designate a starting point for listing folders. For instance, use "/Inbox/From Accounting/" to get all subfolders of the 'From Accounting' folder in your Inbox.
.EXAMPLE
    ReturnFolderList user@domain.com
    This command will return a list of all user-accessible folders for the user@domain.com mailbox.
.INPUTS
    SMTP address of the mailbox, with optional parent folder (full path).
.OUTPUTS
    Array with information about the mailbox folders.
#>

    param(
    [Parameter(Mandatory=$true, ValueFromPipeline=$true)]$SMTPAddress,
    [ValidateNotNullOrEmpty()][String[]]$ParentFolderPath)

    #Make sure we are connected to Exchange Remote PowerShell
    if (Check-Connectivity) { Write-Verbose "Connected to Exchange Remote PowerShell, processing..." }
    else { Write-Host "ERROR: Connectivity test failed, exiting the script..." -ForegroundColor Red; continue }

    $MBfolders = Get-ExOMailboxFolderStatistics -Identity $SMTPAddress -Verbose:$false | Select-Object Name,FolderType,FolderPath,Identity
    if($PSBoundParameters.ContainsKey('ParentFolderPath')) {
        $MBfolders = $MBfolders | ? {($_.FolderType -eq "User created" -or $_.FolderType -in $includedfolders) -and ($_.Name -notin $excludedfolders) -and ($_.FolderPath -match $ParentFolderPath+"*")}
    }
    else {
        $MBfolders = $MBfolders | ? {($_.FolderType -eq "User created" -or $_.FolderType -in $includedfolders) -and ($_.Name -notin $excludedfolders)}
    }

    if (!$MBfolders) { return }
    else { return ($MBfolders | select Name,FolderType,Identity) }
}

function GetFolderPermissions {
<#
.Synopsis
    Enumerates all permissions for the given mailbox folder
.DESCRIPTION
    The GetFolderPermissions cmdlet enumerates the permissions for each folder in the given mailbox. To adjust the list of folders, add to the $includedfolders or $excludedfolders array, respectively.
.EXAMPLE
    GetFolderPermissions user@domain.com:\Calendar

    This command will return a list of permissions for the Calendar fodler in the user@domain.com mailbox.
.INPUTS
    Identifier for the folder.
.OUTPUTS
    Array with information about the mailbox folder permissions.
#>

    param([Parameter(Mandatory=$true)]$foldername)

    if (Check-Connectivity) { Write-Verbose "Connected to Exchange Remote PowerShell, processing..." }
    else { Write-Host "ERROR: Connectivity test failed, exiting the script..." -ForegroundColor Red; continue }

    $FolderPerm = Get-MailboxFolderPermission $foldername -ErrorAction Stop | select Identity,@{n="User";e={$_.User.RecipientPrincipal.Guid.Guid}},@{n="UserType";e={$_.User.UserType.ToString()}},@{n="UserName";e={$_.User.DisplayName}},AccessRights,SharingPermissionFlags

    if (!$FolderPerm) { return }
    else { return $FolderPerm }
}

function Reset-MailboxFolderPermissionsRecursive {
<#
.Synopsis
    Restores the default permissions for all user-accessible folders for a given mailbox.
.DESCRIPTION
    The Reset-MailboxFolderPermissionsRecursive cmdlet removes permissions for all user-accessible folders for the given mailbox(es), specified via the -Mailbox parameter. The list of folders is generated via the ReturnFolderList function. Configure the $includedfolders and $excludedfolders variables to granularly control the folder list.
.PARAMETER Mailbox
    Use the -Mailbox parameter to designate the mailbox. Any valid Exchange mailbox identifier can be specified. Multiple mailboxes can be specified in a comma-separated list or array, see examples below.
.PARAMETER ParentFolderPath
	Use the -ParentFolderPath to designate a starting point for listing folders. For instance, use "/Inbox/From Accounting/" to get all subfolders of the 'From Accounting' folder in your Inbox.
.PARAMETER Quiet
    Use the -Quiet switch if you want to suppress output to the console.
.PARAMETER WhatIf
    The -WhatIf switch simulates the actions of the command. You can use this switch to view the changes that would occur without actually applying those changes.
.PARAMETER Verbose
    The -Verbose switch provides additional details on the cmdlet progress, it can be useful when troubleshooting issues.
.EXAMPLE
    Reset-MailboxFolderPermissionsRecursive -Mailbox user@domain.com

    This command removes permissions on all user-accessible folders in the user@domain.com mailbox.
.EXAMPLE
    Reset-MailboxFolderPermissionsRecursive -Mailbox shared@domain.com,room@domain.com

    This command removes permissions on all user-accessible folders in BOTH the room@domain.com and shared@domain.com mailboxes.
.EXAMPLE
    Reset-MailboxFolderPermissionsRecursive -Mailbox (Get-Mailbox -RecipientTypeDetails RoomMailbox) -Verbose

    This command removes permissions on all user-accessible folders in ALL Room mailboxes in the organization.
.INPUTS
    A mailbox identifier.
.OUTPUTS
    Array of Mailbox address, Folder name and User.
#>

    [cmdletbinding(SupportsShouldProcess)]

    Param(
    [Parameter(Mandatory=$true,ValueFromPipeline=$false)][ValidateNotNullOrEmpty()][Alias("Identity")][String[]]$Mailbox,
    [ValidateNotNullOrEmpty()][String[]]$ParentFolderPath,
    [switch]$ResetDefaultLevel,
    [switch]$Quiet)


#region BEGIN
    #Make sure we are connected to Exchange Remote PowerShell
    if (Check-Connectivity) { Write-Verbose "Connected to Exchange Remote PowerShell, processing..." }
    else { Write-Host "ERROR: Connectivity test failed, exiting the script..." -ForegroundColor Red; continue }

    #Prepare the list of mailboxes
    Write-Verbose "Parsing the Mailbox parameter..."
    $SMTPAddresses = @{}
    foreach ($mb in $Mailbox) {
        Start-Sleep -Milliseconds 80 #Add some delay to avoid throttling...
        #Make sure a matching mailbox is found and return its Primary SMTP Address
        $SMTPAddress = Get-ExOMailbox $mb -ErrorAction SilentlyContinue -Verbose:$false | Select-Object -ExpandProperty PrimarySmtpAddress
        if (!$SMTPAddress) { if (!$Quiet) { Write-Warning "Mailbox with identifier $mb not found, skipping..." }; continue }
        elseif (($SMTPAddress.count -gt 1) -or ($SMTPAddresses[$mb]) -or ($SMTPAddresses.ContainsValue($SMTPAddress))) { Write-Warning "Multiple mailboxes matching the identifier $mb found, skipping..."; continue }
        else { $SMTPAddresses[$mb] = $SMTPAddress }
    }
    if (!$SMTPAddresses -or ($SMTPAddresses.Count -eq 0)) { Throw "No matching mailboxes found, check the parameter values." }
    Write-Verbose "The following list of mailboxes will be used: ""$($SMTPAddresses.Values -join ", ")"""
    Write-Verbose "List of default folder TYPES that will be used: ""$($includedfolders -join ", ")"""
    Write-Verbose "List of folder NAMES that will be excluded: ""$($excludedfolders -join ", ")"""
    if ($PSBoundParameters.ContainsKey('ParentFolderPath')) { Write-Verbose "Using the -ParentFolderPath parameter to only act on folders under: ""$ParentFolderPath""" }
#endregion

#region PROCESS
    $out = @()
    foreach ($smtp in $SMTPAddresses.Values) {
        Write-Verbose "Processing mailbox ""$smtp""..."
        Start-Sleep -Milliseconds 800 #Add some delay to avoid throttling...

        Write-Verbose "Obtaining folder list for mailbox ""$smtp""..."
        if($PSBoundParameters.ContainsKey('ParentFolderPath')) { $folders = ReturnFolderList $smtp $ParentFolderPath }
		else { $folders = ReturnFolderList $smtp }
        Write-Verbose "A total of $($folders.count) folders found for $($smtp)."

        if (!$folders) { Write-Verbose "No matching folders found for $($smtp), skipping..." ; continue }

        #Cycle over each folder we are interested in
        foreach ($folder in $folders) {
            #"Fix" for folders with "/" characters, treat the Root folder separately
            if ($folder.FolderType -eq "Root") { $foldername = $smtp }
            else { $foldername = $folder.Identity.ToString().Replace([char]63743,"/").Replace($smtp,$smtp + ":") }

            Write-Verbose "Processing folder ""$foldername""..."
            $fPermissions = GetFolderPermissions $foldername
            if (!$ResetDefaultLevel) { $fPermissions = $fPermissions | ? {$_.UserName -notin @("Default","Anonymous","Owner@local","Member@local")}} #filter out default permissions
            if (!$fPermissions) { Write-Verbose "No permission entries found for $($foldername), skipping..." ; continue }

            #Remove the folder permissions for each delegate
            foreach ($u in $fPermissions) {
                if ($u.UserType -eq "Default") {#UserType enumeration https://docs.microsoft.com/en-us/previous-versions/office/developer/exchange-server-2010/ff319704(v%3Dexchg.140)
                    if ($ResetDefaultLevel) {
                        try {
                            Write-Verbose "Resetting permissions on ""$foldername"" for principal ""Default""."
                            if ($folder.FolderType -eq "Calendar") {
                                if (($u.AccessRights -join ",") -ne "AvailabilityOnly") { Set-MailboxFolderPermission -Identity $foldername -User Default -AccessRights AvailabilityOnly -WhatIf:$WhatIfPreference -Confirm:$false -ErrorAction Stop }
                                else { continue }
                                $outtemp = New-Object psobject -Property ([ordered]@{"Mailbox" = $smtp;"FolderName" = $folder.name;"User" = $u.UserName;"AccessRights" = "AvailabilityOnly"})
                                }
                            else {
                                if (($u.AccessRights -join ",") -ne "None") { Set-MailboxFolderPermission -Identity $foldername -User Default -AccessRights None -WhatIf:$WhatIfPreference -Confirm:$false -ErrorAction Stop }
                                else { continue }
                                $outtemp = New-Object psobject -Property ([ordered]@{"Mailbox" = $smtp;"FolderName" = $folder.name;"User" = $u.UserName;"AccessRights" = "None"})
                            }
                            $out += $outtemp; if (!$Quiet -and !$WhatIfPreference) { $outtemp } #Write output to the console unless the -Quiet parameter is used
                        }
                        catch [System.Exception] {
                            if ($_.Exception.Message -match "ManagementObjectNotFoundException") { Write-Host "ERROR: Folder ""$foldername"" not found, this should not happen..." -ForegroundColor Red }
                        }
                        catch {$_ | fl * -Force; continue} #catch-all for any unhandled errors
                    }
                    else { continue }
                }
                elseif ($u.UserType -eq "Anonymous") { continue } #Maybe set them all to none when $resetdefault is used?
                elseif ($u.UserType -eq "Unknown") { Write-Verbose "Skipping orphaned permissions entry: $($u.UserName)"; continue }
                else {
                    if ($u.UserType -eq "External") { $u.User = $u.UserName }
                    try {
                        if (!$u.User) { continue }
                        Write-Verbose "Removing permissions on ""$foldername"" for principal ""$($u.UserName)""."
                        Remove-MailboxFolderPermission -Identity $foldername -User $u.User -WhatIf:$WhatIfPreference -Confirm:$false -ErrorAction Stop
                        $outtemp = New-Object psobject -Property ([ordered]@{"Mailbox" = $smtp;"FolderName" = $folder.name;"User" = $u.UserName})
                        $out += $outtemp; if (!$Quiet -and !$WhatIfPreference) { $outtemp } #Write output to the console unless the -Quiet parameter is used
                    }
                    catch [System.Exception] {
                        if (!$Quiet) {
                            if ($_.Exception.Message -match "UserNotFoundInPermissionEntryException") { Write-Host "WARNING: No existing permissions entry found on ""$foldername"" for principal ""$($u.UserName)""" -ForegroundColor Yellow }
                            elseif ($_.Exception.Message -match "CannotChangePermissionsOnFolderException") { Write-Host "ERROR: Folder permissions for ""$foldername"" CANNOT be changed!" -ForegroundColor Red }
                            elseif ($_.Exception.Message -match "CannotRemoveSpecialUserException") { Write-Host "ERROR: Folder permissions for ""$($u.UserName)"" CANNOT be changed!" -ForegroundColor Red }
                            elseif ($_.Exception.Message -match "ManagementObjectNotFoundException") { Write-Host "ERROR: Folder ""$foldername"" not found, this should not happen..." -ForegroundColor Red }
                            elseif ($_.Exception.Message -match "InvalidInternalUserIdException|InvalidExternalUserIdException") { Write-Host "ERROR: ""$($u.UserName)"" is not a valid security principal for folder-level permissions..." -ForegroundColor Red }
                            else {$_ | fl * -Force; continue} #catch-all for any unhandled errors
                    }}
                    catch {$_ | fl * -Force; continue} #catch-all for any unhandled errors
                }
            }

    }}
#endregion
    if ($out) {
        Write-Verbose "Exporting results to the CSV file..."
        $out | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxFolderPermissionsRemoved.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
        if (!$Quiet -and !$WhatIfPreference) { return $out | Out-Default } #Write output to the console unless the -Quiet parameter is used
        }
    else { Write-Verbose "Output is empty, skipping the export to CSV file..." }
    Write-Verbose "Finish..."
}

#Invoke the Reset-MailboxFolderPermissionsRecursive function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
if ($PSBoundParameters.Count) { Reset-MailboxFolderPermissionsRecursive @PSBoundParameters -OutVariable global:varFolderPermissionsRemoved }
else { Write-Host "INFO: The script was run without parameters, consider dot-sourcing it instead." -ForegroundColor Cyan }
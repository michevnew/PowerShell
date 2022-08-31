#Requires -Version 3.0
[CmdletBinding(SupportsShouldProcess)] #Make sure we can use -WhatIf and -Verbose
Param(
[switch]$Quiet,[ValidateNotNullOrEmpty()][Alias("Identity")][String[]]$Mailbox,[ValidateNotNullOrEmpty()][Alias("Delegate")][String[]]$User,[Parameter(Mandatory=$false)][String]$ParentFolderPath,
[ValidateSet("None","Owner","PublishingEditor","Editor","PublishingAuthor","Author","NoneditingAuthor","Reviewer","Contributor","FreeBusyTimeOnly","FreeBusyTimeAndSubjectAndLocation",
"Custom","CreateItems","CreateSubfolders","DeleteAllItems","DeleteOwnedItems","EditAllItems","EditOwnedItems","FolderContact","FolderOwner","FolderVisible","ReadItems")]
[ValidateNotNullOrEmpty()][String[]]$AccessRights)

#Include these folder types by default
$includedfolders = @("Root","Inbox","Calendar", "Contacts", "DeletedItems", "Drafts", "JunkEmail", "Journal", "Notes", "Outbox", "SentItems", "Tasks", "CommunicatorHistory", "Clutter", "Archive")
#$includedfolders = @("Root","Inbox","Calendar", "Contacts", "DeletedItems", "SentItems", "Tasks") #Trimmed down list of default folders

#Exclude additional Non-default folders created by Outlook or other mail programs. Folder NAMES, not types! So make sure to include translations too!
#Exclude SearchDiscoveryHoldsFolder and SearchDiscoveryHoldsUnindexedItemFolder as they're not marked as default folders #Exclude "Calendar Logging" on older Exchange versions
$excludedfolders = @("News Feed","Quick Step Settings","Social Activity Notifications","Suggested Contacts", "SearchDiscoveryHoldsUnindexedItemFolder", "SearchDiscoveryHoldsFolder","Calendar Logging")

try { $script:session = Get-PSSession -InstanceId (Get-AcceptedDomain | select -First 1).RunspaceId.Guid -ErrorAction Stop }
catch { Write-Error "No active Exchange Remote PowerShell session detected, please connect first. To connect to ExO: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx" -ErrorAction Stop }

function ReturnFolderList {
<#
.Synopsis
    Enumerates all user-accessible folders for the mailbox
.DESCRIPTION
    The ReturnFolderList cmdlet enumerates the folders for the given mailbox. To adjust the list of folders, add to the $includedfolders or $excludedfolders array, respectively.
.PARAMETER SMTPAddress
	Use the -SMTPAddress parameter to designate the mailbox where the desired folders reside
.PARAMETER ParentFolderPath
	Use the -ParentFolderPath to designate a starting point for listing folders. For instance, use "/Inbox/From Accounting/" to get all subfolders of the 'From Accounting folder in your Inbox.
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
	[Parameter(Mandatory=$false)]$ParentFolderPath)

    if (!$session -or ($session.State -ne "Opened")) { Write-Error "No active Exchange Remote PowerShell session detected, please connect first. To connect to ExO: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx" -ErrorAction Stop }

    $MBfolders = Invoke-Command -Session $session -ScriptBlock { Get-MailboxFolderStatistics $using:SMTPAddress | Select-Object Name,FolderType,FolderPath,Identity } -HideComputerName -ErrorAction Stop
    if($PSBoundParameters.ContainsKey('ParentFolderPath')) {
		$MBfolders = $MBfolders | ? {($_.FolderType -eq "User created" -or $_.FolderType -in $includedfolders) -and ($_.Name -notin $excludedfolders) -and ($_.FolderPath -match $ParentFolderPath+"*")}
	}
	else {
		$MBfolders = $MBfolders | ? {($_.FolderType -eq "User created" -or $_.FolderType -in $includedfolders) -and ($_.Name -notin $excludedfolders)}
	}

    if (!$MBfolders) { return }
    else { return ($MBfolders | select Name,FolderType,Identity) }
}


function Set-MailboxFolderPermissionsRecursive {
<#
.Synopsis
    Sets permissions for all user-accessible folders for a given mailbox.
.DESCRIPTION
    The Set-MailboxFolderPermissionsRecursive cmdlet configures permissions for all user-accessible folders for the given mailbox(es), specified via the -Mailbox parameter. The list of folders is generated via the ReturnFolderList function. Configure the $includedfolders and $excludedfolders variables to granularly control the folder list.
.PARAMETER Mailbox
    Use the -Mailbox parameter to designate the mailbox. Any valid Exchange mailbox identifier can be specified. Multiple mailboxes can be specified in a comma-separated list or array, see examples below.
.PARAMETER User
    Use the -User parameter to designate the delegate. Any valid Exchange security principal identifier can be specified. Multiple delegates can be specified in a comma-separated list or array, see examples below.
.PARAMETER ParentFolderPath
	Use the -ParentFolderPath to designate a starting point for listing folders. For instance, use "/Inbox/From Accounting/" to get all subfolders of the 'From Accounting folder in your Inbox.
.PARAMETER AccessRights
    Use the -AccessRights parameter to specify the permission level to be granted. For list of accepted permissions see for example: https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/add-mailboxfolderpermission?view=exchange-ps
    Roles have precedence over individual permissions entries. If an existing permission entry for the same User is detected, it will be updated to match the newly provided value.
.PARAMETER Quiet
    Use the -Quiet switch if you want to suppress output to the console.
.PARAMETER WhatIf
    The -WhatIf switch simulates the actions of the command. You can use this switch to view the changes that would occur without actually applying those changes.
.PARAMETER Verbose
    The -Verbose switch provides additional details on the cmdlet progress, it can be useful when troubleshooting issues.
.EXAMPLE
    Set-MailboxFolderPermissionsRecursive -Mailbox user@domain.com -User delegate@domain.com -AccessRights Editor
    This command add Editor level permissions on all user-accessible folders in the user@domain.com mailbox for the delegate@domain.com delegate.
.EXAMPLE
    Set-MailboxFolderPermissionsRecursive -Mailbox shared@domain.com,room@domain.com -User delegate@domain.com -AccessRights Owner
    This command add Owner level permissions on all user-accessible folders in BOTH the room@domain.com and shared@domain.com mailboxes for the delegate@domain.com delegate.
.EXAMPLE
    Set-MailboxFolderPermissionsRecursive -Mailbox (Get-Mailbox -RecipientTypeDetails RoomMailbox) -User delegate -AccessRights Owner -Verbose
    This command add Owner level permissions on all user-accessible folders in ALL Room mailboxes in the organization for the delegate.
.INPUTS
    A mailbox identifier, permissions level and delegate identifier.
.OUTPUTS
    Array of Mailbox address, Folder name, User and permission level added.
#>

    [cmdletbinding(SupportsShouldProcess)]

    Param(
    [Parameter(Mandatory=$true,ValueFromPipeline=$false)][ValidateNotNullOrEmpty()][Alias("Identity")][String[]]$Mailbox,
    [Parameter(Mandatory=$true,ValueFromPipeline=$false)][ValidateNotNullOrEmpty()][Alias("Delegate")][String[]]$User,
	[Parameter(Mandatory=$false)]$ParentFolderPath,
    [Parameter(Mandatory=$true)][ValidateSet("None","Owner","PublishingEditor","Editor","PublishingAuthor","Author","NoneditingAuthor","Reviewer","Contributor","FreeBusyTimeOnly","FreeBusyTimeAndSubjectAndLocation",
    "Custom","CreateItems","CreateSubfolders","DeleteAllItems","DeleteOwnedItems","EditAllItems","EditOwnedItems","FolderContact","FolderOwner","FolderVisible","ReadItems")]
    [ValidateNotNullOrEmpty()][String[]]$AccessRights,
    [switch]$Quiet)


#region BEGIN
    #Make sure we are connected to Exchange Remote PowerShell
    Write-Verbose "Checking connectivity to Exchange Remote PowerShell..."
    if (!$session -or ($session.State -ne "Opened")) {
        try { $script:session = Get-PSSession -InstanceId (Get-AcceptedDomain | select -First 1).RunspaceId.Guid -ErrorAction Stop  }
        catch { Write-Error "No active Exchange Remote PowerShell session detected, please connect first. To connect to ExO: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx" -ErrorAction Stop }
    }

    #Prepare the list of mailboxes
    Write-Verbose "Parsing the Mailbox parameter..."
    $SMTPAddresses = @{}
    foreach ($mb in $Mailbox) {
        Start-Sleep -Milliseconds 80 #Add some delay to avoid throttling...
        #Make sure a matching mailbox is found and return its Primary SMTP Address
        $SMTPAddress = (Invoke-Command -Session $session -ScriptBlock { Get-Mailbox $using:mb | Select-Object -ExpandProperty PrimarySmtpAddress } -ErrorAction SilentlyContinue).Address
        if (!$SMTPAddress) { if (!$Quiet) { Write-Warning "Mailbox with identifier $mb not found, skipping..." }; continue }
        elseif (($SMTPAddress.count -gt 1) -or ($SMTPAddresses[$mb]) -or ($SMTPAddresses.ContainsValue($SMTPAddress))) { Write-Warning "Multiple mailboxes matching the identifier $mb found, skipping..."; continue }
        else { $SMTPAddresses[$mb] = $SMTPAddress }
    }
    if (!$SMTPAddresses -or ($SMTPAddresses.Count -eq 0)) { Throw "No matching mailboxes found, check the parameter values." }
    Write-Verbose "The following list of mailboxes will be used: ""$($SMTPAddresses.Values -join ", ")"""
    
    #Prepare the list of users (security principals)
    Write-Verbose "Parsing the User parameter..."
    $GUIDs = @{}
    foreach ($us in $User) {
        #Allow some of the default values #@("Default","Anonymous","Owner@local","Member@local") #You can only add/set Default and Anonymous!
        if ($us -match "^(Default|Anonymous|Owner@local|Member@local)$") { $GUIDs[$us] = $us;  continue }

        Start-Sleep -Milliseconds 80 #Add some delay to avoid throttling...
        #Make sure a matching security principal object is found and return its UPN
        $GUID = (Invoke-Command -Session $session -ScriptBlock { Get-SecurityPrincipal $using:us | Select-Object -ExpandProperty Guid } -ErrorAction SilentlyContinue).Guid
        if (!$GUID) { if (!$Quiet) { Write-Warning "Security principal with identifier $us not found, skipping..." }; continue }
        elseif (($GUID.count -gt 1) -or ($GUIDs[$us]) -or ($GUIDs.ContainsValue($GUID))) { Write-Warning "Multiple principals matching the identifier $us found, skipping..."; continue }
        else { $GUIDs[$us] = $GUID }
    }
    if (!$GUIDs -or ($GUIDs.Count -eq 0)) { Throw "No matching security principals found, check the parameter values." }
    Write-Verbose "The following list of security principals will be used: ""$($GUIDs.Values -join ", ")"""

    #Validate values for AccessRights
    Write-Verbose "Parsing the AccessRights parameter..."
    #If the string matches any of the built-in roles and no other roles/definitions are present, use said role
    #Ignore Custom, FreeBusyTimeAndSubjectAndLocation and FreeBusyTimeOnly, as we cannot handle those
    if ($AccessRights -in @("Author","Contributor","Editor","None","NonEditingAuthor","Owner","PublishingEditor","PublishingAuthor","Reviewer")) { $permissions = $AccessRights } #matches only a single entry from the list, so OK to use
    #Else, strip all roles (multiple roles cannot be assigned) and look for matching individual permissions instead, combine all valid entries
    else {$permissions = $AccessRights | ? {$_ -match "^(CreateItems|CreateSubfolders|DeleteAllItems|DeleteOwnedItems|EditAllItems|EditOwnedItems|FolderContact|FolderOwner|FolderVisible|ReadItems)$"}}
    if (!$permissions) { Throw "Invalid AccessRights string ""$AccessRights""" }
    Write-Verbose "The following permissions will be used: ""$($permissions -join ", ")"""
    Write-Verbose "List of default folder TYPES that will be used: ""$($includedfolders -join ", ")"""
    Write-Verbose "List of folder NAMES that will be excluded: ""$($excludedfolders -join ", ")"""
#endregion

#region PROCESS   
    $out = @()
    foreach ($smtp in $SMTPAddresses.Values) {#should be unique, if needed select/sort
        Write-Verbose "Processing mailbox ""$smtp""..."
        Start-Sleep -Milliseconds 800 #Add some delay to avoid throttling...
        Write-Verbose "Obtaining folder list for mailbox ""$smtp""..."
        if($PSBoundParameters.ContainsKey('ParentFolderPath')) {$folders = ReturnFolderList $smtp $ParentFolderPath}
		else { $folders = ReturnFolderList $smtp }
        Write-Verbose "A total of $($folders.count) folders found for $($smtp)."

        if (!$folders) { Write-Verbose "No matching folders found for $($smtp), skipping..." ; continue }
        
        #Cycle over each folder we are interested in
        foreach ($folder in $folders) {
            #"Fix" for folders with "/" characters, treat the Root folder separately
            if ($folder.FolderType -eq "Root") { $foldername = $smtp }
            else { $foldername = $folder.Identity.ToString().Replace([char]63743,"/").Replace($smtp,$smtp + ":") }

            #Add/Set the folder permissions for each delegate
            Write-Verbose "Processing folder ""$foldername""..."
            foreach ($u in $GUIDs.Clone().GetEnumerator()) {#Use .Clone() in order to be able to dynamically remove entries if needed...
                try {
                    Write-Verbose "Adding ""$permissions"" permissions on ""$foldername"" for principal ""$($u.Name)""."
                    Invoke-Command -Session $session -ScriptBlock { Add-MailboxFolderPermission -Identity $Using:foldername -User $Using:u.Value -AccessRights $using:permissions -WhatIf:$using:WhatIfPreference } -ErrorAction Stop -HideComputerName | select @{n="Mailbox";e={$smtp}},FolderName,User,@{n="AccessRights";e={$_.AccessRights -join ","}} -OutVariable outtemp | Out-Null
                    $out += $outtemp; if (!$Quiet -and !$WhatIfPreference) { "$outtemp" } #Write output to the console unless the -Quiet parameter is used
                }
                catch [System.Management.Automation.RemoteException] {
                    if ($_.CategoryInfo.Reason -eq "UserAlreadyExistsInPermissionEntryException") { 
                        if (!$Quiet) { Write-Host "WARNING: Existing permissions entry found on ""$foldername"" for principal ""$($u.Name)"", replacing with ""$permissions""." -ForegroundColor Yellow }
                        Invoke-Command -Session $session -ScriptBlock { Set-MailboxFolderPermission -Identity $Using:foldername -User $Using:u.Value -AccessRights $Using:permissions -WhatIf:$using:WhatIfPreference -WarningAction SilentlyContinue } -ErrorAction Stop -HideComputerName #Set- doesnt have any output, but we can suppress the Warnings
                        $outtemp = New-Object psobject -Property ([ordered]@{"Mailbox" = $smtp;"FolderName" = $folder.name;"User" = $u.Name;"AccessRights" = ($permissions -join ",")})
                        $out += $outtemp; if (!$Quiet -and !$WhatIfPreference) { $outtemp } #Make sure the Set- output is displayed unless the -Quiet parameter is used
                    }
                    elseif ($_.CategoryInfo.Reason -eq "CannotChangePermissionsOnFolderException") { Write-Host "ERROR: Folder permissions for ""$foldername"" CANNOT be changed!" -ForegroundColor Red }
                    elseif ($_.CategoryInfo.Reason -eq "ManagementObjectNotFoundException") { Write-Host "ERROR: Folder ""$foldername"" not found, this should not happen..." -ForegroundColor Red }
                    elseif ($_.CategoryInfo.Reason -eq "InvalidInternalUserIdException") { 
                        Write-Host "ERROR: ""$($u.Name)"" is not a valid security principal for folder-level permissions, removing from list..." -ForegroundColor Red
                        $GUIDs.Remove($u.Name)
                        if ($GUIDs.Count) { continue } else { Write-Verbose "No valid security principals for folder-level permissions remaining, exiting the script..." ; return $out | Out-Default }
                        }
                    elseif ($_.CategoryInfo.Reason -eq "ADServerSettingsChangedException") { throw "Connectivity to the Exchange server is broken, existing the script..." } #sometimes results in non-reconnectable session, test/exit as needed
                    else {$_ | fl * -Force; continue} #catch-all for any unhandled errors
                }
                catch {$_ | fl * -Force; continue} #catch-all for any unhandled errors
            }
            
    }}
#endregion
    if ($out) {
        Write-Verbose "Exporting results to the CSV file..."
        $out | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxFolderPermissionsAdded.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
        if (!$Quiet -and !$WhatIfPreference) { return $out | Out-Default } #Write output to the console unless the -Quiet parameter is used
        }
    else { Write-Verbose "Output is empty, skipping the export to CSV file..." }
    Write-Verbose "Finish..."
}

#Invoke the Set-MailboxFolderPermissionsRecursive function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
if ($PSBoundParameters.Count) { Set-MailboxFolderPermissionsRecursive @PSBoundParameters -OutVariable global:varFolderPermissionsAdded }
else { Write-Host "INFO: The script was run without parameters, consider dot-sourcing it instead." -ForegroundColor Cyan }

#Invoke-Command -Session $session -ScriptBlock { Remove-MailboxFolderPermission -Identity $Using:foldername -User $Using:u.Value -Confirm:$false }

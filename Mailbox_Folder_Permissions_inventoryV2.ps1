#Requires -Version 3.0
param([switch]$IncludeAll,[switch]$IncludeUserMailboxes,[switch]$IncludeSharedMailboxes,[switch]$IncludeRoomMailboxes,[switch]$IncludeDefaultPermissions,[string[]]$ExcludeUsers)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/4095/updated-version-of-the-mailbox-folder-permissions-inventory-script

function Get-MailboxFolderPermissionInventory {
<#
.Synopsis
    Lists permissions for all user-accessible folders for all mailboxes of the selected type(s).
.DESCRIPTION
    The Get-MailboxFolderPermissionInventory cmdlet enumerates the folders for all mailboxes of the selected type(s) and lists their permissions. To adjust the list of folders, add to the $includedfolders or $excludedfolders array, respectively.
    Running the cmdlet without parameters will return entries for all User mailboxes only. Specifying particular mailbox type(s) can be done with the corresponding switch parameter.
    The Default permission entry level is not returned unless you specify the -IncludeDefaultPermissions switch when running the cmdlet/script.
    To exclude certain Users from the permission inventory, use the -ExcludeUsers parameter and specify the UPN of the user(s) you want to exclude.
    To specify a variable in which to hold the cmdlet output, use the -OutVariable parameter.

.EXAMPLE
    Get-MailboxFolderPermissionInventory -IncludeUserMailboxes

    This command will return a list of permissions for the user-accessible folders for all User mailboxes.

.EXAMPLE
    Get-MailboxFolderPermissionInventory -IncludeAll -OutVariable global:var
    $var | Export-Csv -NoTypeInformation "accessrights.csv"

    To export the results to a CSV file, use the OutVariable parameter.
.INPUTS
    None.
.OUTPUTS
    Array with information about the mailbox, delegate and type of permissions.
#>

    [CmdletBinding()]

    Param
    (
    #Specify whether to include user mailboxes in the result
    [Switch]$IncludeUserMailboxes,
    #Specify whether to include Shared mailboxes in the result
    [Switch]$IncludeSharedMailboxes,
    #Specify whether to include Room, Equipment and Booking mailboxes in the result
    [Switch]$IncludeRoomMailboxes,
    #Specify whether to return all mailbox types
    [Switch]$IncludeAll,
    #Specify whether to return permissions for the Default entry
    [switch]$IncludeDefaultPermissions,
    #Specify a list of users (UPNs) for which NOT to return permissions (think service accounts, admin accounts, etc)
    [string[]]$ExcludeUsers)
    #Add switch for GroupMailboxes: Get-MailboxFolderPermission -GroupMailbox itsupport
    #Add switch for SupervisoryReviewPolicyMailbox, once they are actually discoverable via Get-Mailbox!

    #Include these folder types by default
    $includedfolders = @("Root", "Inbox", "Calendar", "Contacts", "DeletedItems", "Drafts", "JunkEmail", "Journal", "Notes", "Outbox", "SentItems", "Tasks", "CommunicatorHistory", "Archive", "RssSubscription")
    #$includedfolders = @("Root","Inbox","Calendar", "Contacts", "DeletedItems", "SentItems", "Tasks") #Trimmed down list of default folders

    #Non-default folders created by Outlook or other mail programs. Folder NAMES, not types!
    #Exclude SearchDiscoveryHoldsFolder and SearchDiscoveryHoldsUnindexedItemFolder as they're not marked as default folders
    $excludedfolders = @("News Feed", "Quick Step Settings", "Social Activity Notifications", "Suggested Contacts", "SearchDiscoveryHoldsUnindexedItemFolder", "SearchDiscoveryHoldsFolder", "Calendar Logging") #Exclude "Calendar Logging" on older Exchange versions

    #Initialize the variable used to designate mailbox types, based on the input parameters
    $included = @()
    if($IncludeSharedMailboxes) { $included += "SharedMailbox"}
    if($IncludeRoomMailboxes) { $included += "RoomMailbox"; $included += "EquipmentMailbox"; $included += "SchedulingMailbox"}
    #if no parameters specified, return only User mailboxes
    if($IncludeUserMailboxes -or !$included) { $included += "UserMailbox"}

    #Make sure we have a V2 version of the module
    try { Get-Command Get-EXOMailbox -ErrorAction Stop | Out-Null }
    catch { Write-Error "This script requires the Exchange Online V2 PowerShell module. Learn more about it here: https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module" -ErrorAction Stop }

    #Confirm connectivity to Exchange Online
    try { Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null }
    catch {
        try { Connect-ExchangeOnline -ErrorAction Stop }
        catch { Write-Error "No active Exchange Online session detected. To connect to ExO: https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps" -ErrorAction Stop }
    }

    #Get the list of mailboxes, depending on the parameters specified when invoking the script
    if ($IncludeAll) {
        $MBList = Get-ExOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox,SchedulingMailbox | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails
    }
    else {
        $MBList = Get-ExOMailbox -ResultSize Unlimited -RecipientTypeDetails $included | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails
    }

    #If no mailboxes are returned from the above cmdlet, stop the script and inform the user
    if (!$MBList) { Write-Error "No mailboxes of the specifyied types were found, specify different criteria." -ErrorAction Stop}

    #Once we have the mailbox list, cycle over each mailbox to gather folder permissions inventory
    $arrPermissions = @()
    $count = 1; $PercentComplete = 0;
    foreach ($MB in $MBList) {

        #Progress message
        $ActivityMessage = "Retrieving data for mailbox $($MB.Identity). Please wait..."
        $StatusMessage = ("Processing mailbox {0} of {1}: {2}" -f $count, @($MBList).count, $MB.PrimarySmtpAddress.ToString())
        $PercentComplete = ($count / @($MBList).count * 100)
        Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
        $count++

        #Get the folder statistics for each mailbox and use them to filter out folders we are not interested in
        $MBSMTP = $MB.PrimarySmtpAddress.ToString()
        $MBfolders = Get-ExOMailboxFolderStatistics $MBSMTP
        $MBfolders = $MBfolders | ? {($_.FolderType -eq "User created" -or $_.FolderType -in $includedfolders) -and ($_.Name -notin $excludedfolders)}
        #If no folders left after applying the filters, move to next mailbox
        if (!$MBfolders) { continue }

        #Cycle over each folder we are interested in.
        Start-Sleep -Milliseconds 800 #Add some delay to avoid throttling...
        foreach ($folder in $MBfolders) {
            #Use folderId to avoid issues with special characters. Replace with "human readable" identifier when we get to output
            $folderid = $MBSMTP + ":" + $folder.FolderId
            $foldername = $folder.Identity.ToString().Replace([char]63743,"/").Replace($MBSMTP,$MBSMTP + ":")

            #Get the folder permissions
            $MBrights = Get-ExOMailboxFolderPermission -Identity $folderid

            #Exclude default folders and users as per the parameters passed to the script
            if (!$IncludeDefaultPermissions) { $MBrights = $MBrights | ? {$_.User.DisplayName -notin @("Default","Anonymous","Owner@local","Member@local")}}
            if ($ExcludeUsers) { $MBrights = $MBrights | ? {$_.User.UserPrincipalName -notin $ExcludeUsers}}
            #No non-default permissions found, continue to next folder
            if (!$MBrights) { continue }

            #Prepare the output object
            foreach ($entry in $MBrights) {
                #Prepare the output object
                $objPermissions = New-Object PSObject
                $i++;Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Number" -Value $i
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox address" -Value $MBSMTP
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox type" -Value $MB.RecipientTypeDetails
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Folder identity" -Value $foldername

                $varUser = "";$varType = "";
                if ($entry.User.UserType.ToString() -eq "Internal") { $varUser = $entry.User.UserPrincipalName; $varType = "Internal" }
                elseif ($entry.User.UserType.ToString() -eq "Default") { $varUser = $entry.User.DisplayName; $varType = "Default" }
                elseif ($entry.User.UserType.ToString() -eq "External") { $varUser = $entry.User.UserPrincipalName.Replace("ExchangePublishedUser.",$null); $varType = "External" }
                elseif ($entry.User.UserType.ToString() -eq "Unknown") { $varUser = $entry.User.DisplayName; $varType = "Orphaned" }
                else { continue }

                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "User" -Value $varUser
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Permissions" -Value $($entry.AccessRights -join ";")
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Permission Type" -Value $varType

                $arrPermissions += $objPermissions
                #Uncomment if the script is failing due to connectivity issues, the line below will write the output to a CSV file for each individual permissions entry
                #$objPermissions | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd'))_MailboxFolderPermissions.csv" -Append -NoTypeInformation -Encoding UTF8 -UseCulture
            }}
        }
    #Output the result to the console host. Rearrange/sort as needed.
    if ($arrPermissions) { return ($arrPermissions | select * -ExcludeProperty Number) }
}

#Invoke the Get-MailboxFolderPermissionInventory function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
Get-MailboxFolderPermissionInventory @PSBoundParameters -OutVariable global:varPermissions
$varPermissions | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxFolderPermissions.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
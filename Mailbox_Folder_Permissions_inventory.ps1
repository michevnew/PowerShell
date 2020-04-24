#Requires -Version 3.0
param([switch]$IncludeAll,[switch]$IncludeUserMailboxes,[switch]$IncludeSharedMailboxes,[switch]$IncludeRoomMailboxes,[switch]$CondensedOutput,[switch]$IncludeDefaultPermissions,[string[]]$ExcludeUsers)

function Get-MailboxFolderPermissionInventory {
<#
.Synopsis
    Lists permissions for all user-accessible folders for all mailboxes of the selected type(s).
.DESCRIPTION
    The Get-MailboxFolderPermissionInventory cmdlet enumerates the folders for all mailboxes of the selected type(s) and lists their permissions. To adjust the list of folders, add to the $includedfolders or $excludedfolders array, respectively.
    Running the cmdlet without parameters will return entries for all User mailboxes only. Specifying particular mailbox type(s) can be done with the corresponding switch parameter.
    The Default permission entry level is not returned unless you specify the -IncludeDefaultPermissions switch when running the cmdlet/script.
    To exclude certain Users from the permission inventory, use the -ExcludedUsers parameter.
    To specify a variable in which to hold the cmdlet output, use the -OutVariable parameter.
    To use condensed output (one line per folder), use the -CondensedOutput switch.

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
    #Specify whether to include Room and Equipment mailboxes in the result
    [Switch]$IncludeRoomMailboxes,
    #Specify whether to return all mailbox types
    [Switch]$IncludeAll,
    #Specify whether to write the output in condensed format
    [Switch]$CondensedOutput,
    #Specify whether to return permissions for the Default entry
    [switch]$IncludeDefaultPermissions,
    #Specify a list of users (SMTP addresses) for which NOT to return permissions (think service accounts, admin accounts, etc)
    [string[]]$ExcludeUsers)
    #Add switch for GroupMailboxes: Get-MailboxFolderPermission -GroupMailbox itsupport
    #Add switch for SupervisoryReviewPolicyMailbox, once they are actually discoverable via Get-Mailbox!

    #Include these folder types by default
    $includedfolders = @("Root","Inbox","Calendar", "Contacts", "DeletedItems", "Drafts", "JunkEmail", "Journal", "Notes", "Outbox", "SentItems", "Tasks", "CommunicatorHistory", "Clutter", "Archive")
    #$includedfolders = @("Root","Inbox","Calendar", "Contacts", "DeletedItems", "SentItems", "Tasks") #Trimmed down list of default folders

    #Non-default folders created by Outlook or other mail programs. Folder NAMES, not types!
    #Exclude SearchDiscoveryHoldsFolder and SearchDiscoveryHoldsUnindexedItemFolder as they're not marked as default folders
    $excludedfolders = @("News Feed","Quick Step Settings","Social Activity Notifications","Suggested Contacts", "SearchDiscoveryHoldsUnindexedItemFolder", "SearchDiscoveryHoldsFolder","Calendar Logging") #Exclude "Calendar Logging" on older Exchange versions
    
    #Initialize the variable used to designate mailbox types, based on the input parameters
    $included = @()
    if($IncludeSharedMailboxes) { $included += "SharedMailbox"}
    if($IncludeRoomMailboxes) { $included += "RoomMailbox"; $included += "EquipmentMailbox"}
    #if no parameters specified, return only User mailboxes
    if($IncludeUserMailboxes -or !$included) { $included += "UserMailbox"}
        
    #Confirm connectivity to Exchange Online
    try { $session = Get-PSSession -InstanceId (Get-OrganizationConfig).RunspaceId.Guid -ErrorAction Stop  }
    catch { Write-Error "No active Exchange Online session detected, please connect to ExO first: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx" -ErrorAction Stop }

    #Get the list of mailboxes, depending on the parameters specified when invoking the script
    if ($IncludeAll) {
        $MBList = Invoke-Command -Session $session -ScriptBlock { Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails } -HideComputerName
    }
    else {
        $MBList = Invoke-Command -Session $session -ScriptBlock { Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails $Using:included | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails } -HideComputerName
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
        $MBfolders = Invoke-Command -Session $session -ScriptBlock { Get-MailboxFolderStatistics $using:MBSMTP | Select-Object Name,FolderType,Identity } -HideComputerName
        $MBfolders = $MBfolders | ? {($_.FolderType -eq "User created" -or $_.FolderType -in $includedfolders) -and ($_.Name -notin $excludedfolders)}
        #If no folders left after applying the filters, move to next mailbox
        if (!$MBfolders) { continue }

        #Cycle over each folder we are interested in.
        Start-Sleep -Milliseconds 800 #Add some delay to avoid throttling...
        foreach ($folder in $MBfolders) {
            #"Fix" for folders with "/" characters
            $foldername = $folder.Identity.ToString().Replace([char]63743,"/").Replace($MBSMTP,$MBSMTP + ":")

            #Get the folder permissions
            if ($folder.FolderType -eq "Root") { $MBrights = Invoke-Command -Session $session -ScriptBlock { Get-MailboxFolderPermission -Identity $using:MBSMTP } -HideComputerName }
            else { $MBrights = Invoke-Command -Session $session -ScriptBlock { Get-MailboxFolderPermission -Identity $using:foldername } -HideComputerName }
            
            #Exclude default folders and users as per the parameters passed to the script
            if (!$IncludeDefaultPermissions) { $MBrights = $MBrights | ? {$_.User.DisplayName -notin @("Default","Anonymous","Owner@local","Member@local")}}
            if ($ExcludeUsers) { $MBrights = $MBrights | ? {$_.User.ADRecipient.PrimarySmtpAddress.ToString() -notin $ExcludeUsers}}
            #No non-default permissions found, continue to next folder
            if (!$MBrights) { continue }

            if ($condensedoutput) {
                #Prepare the output object
                $objPermissions = New-Object PSObject
                $i++;Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Number" -Value $i
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox address" -Value $MBSMTP
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox type" -Value $MB.RecipientTypeDetails
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Folder identity" -Value $foldername
                if ($IncludeDefaultPermissions) { Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Default level" -Value $(($MBrights | ? {$_.User.DisplayName -eq "Default"}).AccessRights -join ";") }
                #Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Anonymous level" -Value $(($MBrights | ? {$_.User.DisplayName -eq "Anonymous"}).AccessRights -join ";")

                $internal = "";$external = "";$orphaned = ""
                foreach ($entry in $MBrights) {
                    if ($entry.User.UserType.ToString() -eq "Internal") {
                        $internal = ("$($entry.User.RecipientPrincipal.PrimarySmtpAddress.ToString()):$($entry.AccessRights)" + ";" + $internal)
                    }
                    elseif ($entry.User.UserType.ToString() -eq "External") {
                        $external = ("$($entry.User.RecipientPrincipal.PrimarySmtpAddress.Replace("ExchangePublishedUser.",$null)):$($entry.AccessRights)" + ";" + $external)
                    }
                    elseif ($entry.User.UserType.ToString() -eq "Unknown") {
                        $orphaned = ("$($entry.User.DisplayName):$($entry.AccessRights)" + ";" + $orphaned)
                    }
                    else { continue }
                }
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Internal levels" -Value $internal.Trim(";")
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "External levels" -Value $external.Trim(";")
                Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Orphaned levels" -Value $orphaned.Trim(";")

                $arrPermissions += $objPermissions
                #Uncomment if the script is failing due to connectivity issues, the line below will write the output to a CSV file for each individual permissions entry
                #$objPermissions | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd'))_MailboxFolderPermissions.csv" -Append -NoTypeInformation -Encoding UTF8 -UseCulture
            }
            else {
                #Write each permission entry on separate line
                foreach ($entry in $MBrights) {
                    #Prepare the output object
                    $objPermissions = New-Object PSObject
                    $i++;Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Number" -Value $i
                    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox address" -Value $MBSMTP
                    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox type" -Value $MB.RecipientTypeDetails
                    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Folder identity" -Value $foldername

                    $varUser = "";$varType = "";
                    if ($entry.User.UserType.ToString() -eq "Internal") { $varUser = $entry.User.RecipientPrincipal.PrimarySmtpAddress.ToString(); $varType = "Internal" }
                    elseif ($entry.User.UserType.ToString() -eq "Default") { $varUser = $entry.User.DisplayName; $varType = "Default" }
                    elseif ($entry.User.UserType.ToString() -eq "External") { $varUser = $entry.User.RecipientPrincipal.PrimarySmtpAddress.Replace("ExchangePublishedUser.",$null); $varType = "External" }
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
        }  
    #Output the result to the console host. Rearrange/sort as needed.
    $arrPermissions | select * -ExcludeProperty Number,PSComputerName,RunspaceId,PSShowComputerName
}

#Invoke the Get-MailboxFolderPermissionInventory function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
Get-MailboxFolderPermissionInventory @PSBoundParameters -OutVariable global:varPermissions | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxFolderPermissions.csv" -NoTypeInformation -Encoding UTF8 -UseCulture

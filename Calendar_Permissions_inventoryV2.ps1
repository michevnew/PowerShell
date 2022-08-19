param([switch]$IncludeAll,[switch]$IncludeUserMailboxes,[switch]$IncludeSharedMailboxes,[switch]$IncludeRoomMailboxes)

#For details on what the script does and how to run it, check: https://www.michev.info/Blog/Post/4007/

function Get-CalendarPermissionInventory {
<#
.Synopsis
    Lists all permissions on the default Calendar folder for all mailboxes of the selected type(s).
.DESCRIPTION
    The Get-CalendarPermissionInventory cmdlet finds the default Calendar folder for all mailboxes of the selected type(s) and lists its permissions.
    Running the cmdlet without parameters will return entries for User mailboxes only. Specifying particular mailbox type(s) can be done with the corresponding switch parameter.
    To specify a variable in which to hold the cmdlet output, use the OutVariable parameter.

.EXAMPLE
    Get-CalendarPermissionInventory -IncludeUserMailboxes

    This command will return a list of permissions for the default Calendar folder for all user mailboxes.

.EXAMPLE
    Get-CalendarPermissionInventory -IncludeAll -OutVariable global:var
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
    [Switch]$IncludeAll)

    
    #Initialize the variable used to designate mailbox types, based on the input parameters
    $included = @()
    if($IncludeSharedMailboxes) { $included += "SharedMailbox"}
    if($IncludeRoomMailboxes) { $included += "RoomMailbox"; $included += "EquipmentMailbox"; $included += "SchedulingMailbox"}
    #if no parameters specified, return only User mailboxes
    if($IncludeUserMailboxes -or !$included) { $included += "UserMailbox"}

    #Make sure we have a V2 version of the module
    try { Get-Command Get-EXOMailbox -ErrorAction Stop | Out-Null }
    catch { "This script requires the Exchange Online V2 PowerShell module. Learn more about it here: https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module" } 
    
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
    if (!$MBList) { Write-Error "No mailboxes of the specifyied types were found, specify different criteria." -ErrorAction Stop }

    #Once we have the mailbox list, cycle over each mailbox to gather Calendar permissions inventory
    $arrPermissions = @()
    $count = 1; $PercentComplete = 0;
    foreach ($MB in $MBList) {
        
        #Progress message
        $ActivityMessage = "Retrieving data for mailbox $($MB.Identity). Please wait..."
        $StatusMessage = ("Processing {0} of {1}: {2}" -f $count, @($MBList).count, $MB.PrimarySmtpAddress.ToString())
        $PercentComplete = ($count / @($MBList).count * 100)
        Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
        $count++

        #Get the default Calendar folder for each mailbox, depending on localization, etc.
        $calendarfolder = (Get-ExOMailboxFolderStatistics $MB.PrimarySmtpAddress.ToString() -FolderScope Calendar | ? {$_.FolderType -eq "Calendar"}).Identity.ToString().Replace("\",":\").Replace([char]63743,"/")
        #Get the Calendar folder permissions
        $MBrights = Get-ExOMailboxFolderPermission -Identity $calendarfolder
        #No non-default permissions found, continue to next mailbox
        if (!$MBrights) { continue }

        foreach ($entry in $MBrights) {
            #Prepare the output object
            $objPermissions = New-Object PSObject
            $i++;Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Number" -Value $i
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox address" -Value $MB.PrimarySmtpAddress.ToString()
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox type" -Value $MB.RecipientTypeDetails.ToString()
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Calendar folder" -Value $calendarfolder

            $varUser = "";$varType = "";
            if ($entry.User.UserType.ToString() -eq "Internal") { $varUser = $entry.User.UserPrincipalName.ToString(); $varType = "Internal" }
            elseif ($entry.User.UserType.ToString() -eq "Default") { $varUser = $entry.User.DisplayName; $varType = "Default" }
            elseif ($entry.User.UserType.ToString() -eq "External") { $varUser = $entry.User.UserPrincipalName.Replace("ExchangePublishedUser.",$null); $varType = "External" }
            elseif ($entry.User.UserType.ToString() -eq "Unknown") { $varUser = $entry.User.DisplayName; $varType = "Orphaned" }
            else { continue }
                    
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "User" -Value $varUser
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Permissions" -Value $($entry.AccessRights -join ";")
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Permission Type" -Value $varType
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Sharing permission flags" -Value $entry.SharingPermissionFlags

            $arrPermissions += $objPermissions
        }
    }
    #Output the result to the console host. Rearrange/sort as needed.
    if ($arrPermissions) { return ($arrPermissions | select * -ExcludeProperty Number) }
}

#Invoke the Get-CalendarPermissionInventory function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
Get-CalendarPermissionInventory @PSBoundParameters -OutVariable global:varPermissions 
$varPermissions | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_CalendarPermissions.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
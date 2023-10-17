param([switch]$IncludeAll,[switch]$IncludeUserMailboxes,[switch]$IncludeSharedMailboxes,[switch]$IncludeRoomMailboxes,[switch]$IncludeSoftDeleted)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/4021/updated-version-of-the-mailbox-permissions-inventory-script

function Get-MailboxPermissionInventory {
<#
.Synopsis
    Lists all non-default permissions for all mailboxes of the selected type(s).
.DESCRIPTION
    The Get-MailboxPermissionInventory cmdlet lists all mailboxes of the selected type(s) that have at least one object with non-default permissions added. Running the cmdlet without parameters will return entries for all User, Shared, Room, Equipment, Scheduling, and Team mailboxes.
    Specifying particular mailbox type(s) can be done with the corresponding parameter. To specify a variable in which to hold the cmdlet output, use the OutVariable parameter.
    To include soft-deleted mailboxes in the output, use the -IncludeSoftDeleted switch.

.EXAMPLE
    Get-MailboxPermissionInventory -IncludeUserMailboxes

    This command will return a list of user mailboxes that have at least one delegate, along with the delegate permissions.

.EXAMPLE
    Get-MailboxPermissionInventory -IncludeAll -OutVariable global:var
    $var | Export-Csv -NoTypeInformation "accessrights.csv"

    To export the results to a CSV file, use the OutVariable parameter.
.INPUTS
    None.
.OUTPUTS
    Array with information about the mailbox, delegate and type of permissions applied.
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
    #Specify whether to include soft-deleted mailboxes in the result
    [Switch]$IncludeSoftDeleted,
    #Specify whether to include every type of mailbox in the result
    [Switch]$IncludeAll)

    #Initialize the variable used to designate recipient types, based on the script parameters
    $included = @()
    if ($IncludeUserMailboxes) { $included += "UserMailbox"}
    if ($IncludeSharedMailboxes) { $included += "SharedMailbox"}
    if ($IncludeRoomMailboxes) { $included += "RoomMailbox"; $included += "EquipmentMailbox"; $included += "SchedulingMailbox"}

    #Make sure we have a V2 version of the module
    try { Get-Command Get-EXOMailbox -ErrorAction Stop | Out-Null }
    catch { Write-Error "This script requires the Exchange Online V2 PowerShell module. Learn more about it here: https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module" -ErrorAction Stop}

    #Confirm connectivity to Exchange Online
    try { Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null }
    catch {
        try { Connect-ExchangeOnline -ErrorAction Stop }
        catch { Write-Error "No active Exchange Online session detected. To connect to ExO: https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps" -ErrorAction Stop }
    }

    #Get the list of mailboxes, depending on the parameters specified when invoking the script
    if ($IncludeAll -or !$included) {
        $MBList = Get-ExOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox,SchedulingMailbox | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails
        if ($IncludeSoftDeleted) { $MBList += Get-ExOMailbox -SoftDeletedMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox,SchedulingMailbox | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails }
    }
    else {
        $MBList = Get-ExOMailbox -ResultSize Unlimited -RecipientTypeDetails $included | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails
        if ($IncludeSoftDeleted) { $MBList += Get-ExOMailbox -SoftDeletedMailbox -ResultSize Unlimited -RecipientTypeDetails $included | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails }
    }

    #If no mailboxes are returned from the above cmdlet, stop the script and inform the user
    if (!$MBList) { Write-Error "No mailboxes of the specified types were found, specify different criteria." -ErrorAction Stop}

    #Once we have the mailbox list, cycle over each mailbox to gather permissions inventory
    $arrPermissions = @()
    $count = 1; $PercentComplete = 0;
    foreach ($MB in $MBList) {
        #Progress message
        $ActivityMessage = "Retrieving data for mailbox $($MB.Identity). Please wait..."
        $StatusMessage = ("Processing {0} of {1}: {2}" -f $count, @($MBList).count, $MB.PrimarySmtpAddress.ToString())
        $PercentComplete = ($count / @($MBList).count * 100)
        Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
        $count++

        #Gather permissions for each mailbox. Uncomment the end part to only return Full Access permissions and ignore orphaned entries
        if ($MB.Identity -match "Soft Deleted Objects\\") { $MBrights = Get-ExOMailboxPermission -SoftDeletedMailbox -Identity $MB.Identity | ? {($_.User -ne "NT AUTHORITY\SELF") -and ($_.IsInherited -ne $true)}} #or better use GUID?
        else { $MBrights = Get-ExOMailboxPermission -Identity $MB.PrimarySmtpAddress.ToString() | ? {($_.User -ne "NT AUTHORITY\SELF") -and ($_.IsInherited -ne $true)}}
        #No non-default permissions found, continue to next mailbox
        if (!$MBrights) { continue }

        foreach ($entry in $MBrights) {
            #Prepare the output object
            $objPermissions = New-Object PSObject
            $i++;Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Number" -Value $i
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "User" -Value $entry.user
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "User Sid" -Value $entry.UserSid
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox address" -Value $MB.PrimarySmtpAddress.ToString()
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox type" -Value $MB.RecipientTypeDetails
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Is soft-deleted" -Value (& {If($MB.Identity -match "Soft Deleted Objects\\") {"True"} else {"False"}})
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Access Rights" -Value ($entry.AccessRights -join ",")

            $arrPermissions += $objPermissions
        }
    }

    #Output the result to the console host. Rearrange/sort as needed.

    if ($arrPermissions) { return ($arrPermissions | select * -ExcludeProperty Number) }
}

#Invoke the Get-MailboxPermissionInventory function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
Get-MailboxPermissionInventory @PSBoundParameters -OutVariable global:varPermissions
$varPermissions | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxPermissions.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
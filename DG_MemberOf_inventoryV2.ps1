param([switch]$IncludeAll,[switch]$IncludeUserMailboxes,[switch]$IncludeSharedMailboxes,[switch]$IncludeRoomMailboxes,[switch]$IncludeMailUsers,[switch]$IncludeMailContacts,[switch]$IncludeGuestUsers,[switch]$IncludeUsers)

function Check-Connectivity {
    #Make sure we are connected to Exchange Online PowerShell
    Write-Verbose "Checking connectivity to Exchange Online PowerShell..."

    #Check via Get-ConnectionInformation first
    if (Get-ConnectionInformation) { return $true }

    #Make sure we have a V2 version of the module
    try { Get-Command Get-EXOMailbox -ErrorAction Stop | Out-Null }
    catch { Write-Error "This script requires the Exchange Online V2/V3 PowerShell module. Learn more about it here: https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module"; return $false } 
    
    #Confirm connectivity to Exchange Online
    try { Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null }
    catch {
        try { Connect-ExchangeOnline -ErrorAction Stop }
        catch { Write-Error "No active Exchange Online session detected. To connect to ExO: https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps"; return $false }
    }

    return $true
}


function Get-DGMembershipInventory {
<#
.Synopsis
    Lists all groups for which a given user is a member of, for all of the selected user types.
.DESCRIPTION
    The Get-DGMembershipInventory cmdlet prepares a list of all users of the selected type(s) and for each user, returns all groups he is a member of. Running the cmdlet without parameters will return entries for all User mailboxes.
    Specifying particular user type(s) can be done with the corresponding parameter. To specify a variable in which to hold the cmdlet output, use the OutVariable parameter.

.EXAMPLE
    Get-DGMembershipInventory -IncludeUserMailboxes

    This command will return a list of groups for each user mailboxes in the company

.EXAMPLE
    Get-DGMembershipInventory -IncludeAll -OutVariable global:var
    $var | Export-Csv -NoTypeInformation "DGMemberOf.csv"

    To export the results to a CSV file, use the OutVariable parameter.
.INPUTS
    None.
.OUTPUTS
    Array with information about the user and a list of groups he's a member of
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
    #Specify whether to include Mail users in the result
    [Switch]$IncludeMailUsers,
    #Specify whether to include Mail contacts in the result
    [Switch]$IncludeMailContacts,
    #Specify whether to include Guest (Mail) users in the result
    [Switch]$IncludeGuestUsers,
    #Specify whether to include User recipients in the result (RecipientTypeDetails="User")
    [switch]$IncludeUsers,
    #Specify whether to include every type of recipient in the result
    [Switch]$IncludeAll)

    #Initialize the variable used to designate recipient types, based on the script parameters
    $included = @()
    if($IncludeUserMailboxes) { $included += "UserMailbox" }
    if($IncludeSharedMailboxes) { $included += "SharedMailbox" }
    if($IncludeRoomMailboxes) { $included += "RoomMailbox"; $included += "EquipmentMailbox"; $included += "SchedulingMailbox"}
    if($IncludeMailUsers) { $included += "MailUser" }
    if($IncludeMailContacts) { $included += "MailContact" }
    if($IncludeGuestUsers) { $included += "GuestMailUser" }
    #if($IncludeUsers) { $included += "User" } #not needed, will mess up the array, separate check below
    
    #Check if we are connected to Exchange PowerShell
    if (!(Check-Connectivity)) { return }

    #Get the list of users, depending on the parameters specified when invoking the script
    #Stick to Get-Recipient, as Get-EXORecipient has problems with special characters ("#" or ":")
    $props = @("PrimarySmtpAddress","DistinguishedName","ExternalEmailAddress","ExternalDirectoryObjectId","RecipientTypeDetails")

    #Cover RecipientTypeDetails User
    if ($IncludeUsers -or $IncludeAll) {
        $MBList += Get-User -ResultSize Unlimited -RecipientTypeDetails User | Select-Object -Property $props
    }
    
    #Cover the rest of the recipient types
    if ($IncludeAll) {
        $MBList += Get-Recipient -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox,SchedulingMailbox,MailUser,MailContact,GuestMailUser | Select-Object -Property $props
    }
    elseif (!$included -or ($included -eq "UserMailbox" -and $Included.Length -eq 1)) {
        $MBList += Get-Recipient -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Select-Object -Property $props
    }
    else {
        $MBList += Get-Recipient -ResultSize Unlimited -RecipientTypeDetails $included | Select-Object -Property $props
    }
    
    #If no users are returned from the above cmdlet, stop the script and inform the user
    if (!$MBList) { Write-Error "No users of the specifyied types were found, specify different criteria." -ErrorAction Stop }

    #prepare the output
    $arrMembers = @(); $count = 1; $PercentComplete = 0;

    #cycle over each object from the list
    foreach ($mailbox in $MBList) { 
        #Display a simple progress message
        $ActivityMessage = "Retrieving data for mailbox $($mailbox.PrimarySmtpAddress). Please wait..."
        $StatusMessage = ("Processing {0} of {1}: {2}" -f $count, @($MBList).count, $mailbox.DistinguishedName)
        $PercentComplete = ($count / @($MBList).count * 100)
        Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
        $count++

        #Add some delay to avoid throttling
        Start-Sleep -Milliseconds 10

        #Use server-side filtering to obtain the list of groups a given user is a member of
        $dn =  $mailbox.DistinguishedName.Replace("'","''") #handle ' in DN
        $list = Get-Recipient -Filter "Members -eq '$dn'" | Select-Object -Property PrimarySmtpAddress

        #Prepare the output
        $objMembers = New-Object PSObject
        if ($mailbox.PrimarySmtpAddress) { Add-Member -InputObject $objMembers -MemberType NoteProperty -Name "Identifier" -Value $mailbox.PrimarySmtpAddress.ToString() }
        elseif ($mailbox.ExternalEmailAddress) { Add-Member -InputObject $objMembers -MemberType NoteProperty -Name "Identifier" -Value $mailbox.ExternalEmailAddress.ToString() }
        else { Add-Member -InputObject $objMembers -MemberType NoteProperty -Name "Identifier" -Value $mailbox.ExternalDirectoryObjectId.ToString() }
        Add-Member -InputObject $objMembers -MemberType NoteProperty -Name "RecipientTypeDetails" -Value $($mailbox.RecipientTypeDetails)
        Add-Member -InputObject $objMembers -MemberType NoteProperty -Name "Groups" -Value $($list.PrimarySmtpAddress -join ",")
        $arrMembers += $objMembers
    }

    #return the output
    $arrMembers | select Identifier,RecipientTypeDetails,Groups
}

#Invoke the Get-DGMembershipInventory function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
Get-DGMembershipInventory @PSBoundParameters -OutVariable global:varDGMemberOf | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_DGMemberOfReport.csv" -NoTypeInformation
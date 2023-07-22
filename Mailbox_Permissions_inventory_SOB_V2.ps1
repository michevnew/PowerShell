param([switch]$IncludeAll,[switch]$IncludeUserMailboxes,[switch]$IncludeSharedMailboxes,[switch]$IncludeRoomMailboxes,[switch]$IncludeGroupMailboxes,[switch]$IncludeDGs)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/5576/script-send-on-behalf-of-permissions-microsoft-365

function Get-SOBPermissionInventory {
<#
.Synopsis
    Lists all Send on behalf of permissions for all recipients of the selected type(s).
.DESCRIPTION
    The Get-SOBPermissionInventory cmdlet lists all recipients of the selected type(s) that have at least one Send on Behalf of delegate added.
    Running the cmdlet without parameters will return entries for all User, Shared, Room, Equipment, Scheduling, Discovery, Team and Group mailboxes, as well as Distribution, Dynamic Distribution and Mail-enabled Security groups.
    Specifying particular mailbox type(s) can be done with the corresponding parameter. To specify a variable in which to hold the cmdlet output, use the OutVariable parameter.

.EXAMPLE
    Get-SOBPermissionInventory -IncludeUserMailboxes

    This cmdlet will return a list of user mailboxes that have at least one delegate, along with the delegate permissions.

.EXAMPLE
    Get-SOBPermissionInventory -IncludeAll -OutVariable global:var
    $var | Export-Csv -NoTypeInformation "accessrights.csv"

    To export the results to a CSV file, use the OutVariable parameter.
.INPUTS
    None
.OUTPUTS
    Array with information about the recipient, delegate and type of permissions applied.
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
    #Specify whether to include Group mailboxes in the result
    [Switch]$IncludeGroupMailboxes,#Get-Mailbox will not return Team or Group mailboxes, needs special treatment
    #Specify whether to include Distribution groups in the result (includes DGs, Mail-enabled SGs, Room Lists)
    [switch]$IncludeDGs,
    #Specify whether to include every possible type of recipients in the result
    [Switch]$IncludeAll)


    #Initialize the variable used to designate recipient types, based on the script parameters
    $included = @();
    if($IncludeUserMailboxes) { $included += "UserMailbox"}
    if($IncludeSharedMailboxes) { $included += "SharedMailbox"}
    if($IncludeRoomMailboxes) { $included += "RoomMailbox"; $included += "EquipmentMailbox"; $included += "SchedulingMailbox" }
    if($IncludeGroupMailboxes) { $included += "GroupMailbox"} #Used to prevent empty $included
    if($IncludeDGs) { $included += "GroupMailbox"} #Used to prevent empty $included

    #Confirm connectivity to Exchange Online.
    Write-Verbose "Connecting to Exchange Online..."
    try { Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null }
    catch {
        try { Connect-ExchangeOnline -CommandName Get-Mailbox,Get-UnifiedGroup,Get-DistributionGroup,Get-DynamicDistributionGroup -SkipLoadingFormatData } #needs to be non-REST cmdlet
        catch { Write-Error "No active Exchange Online session detected. To connect to ExO: https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps";return }
    }

    #Get the list of recipients with SOB permissions configured, depending on the parameters specified when invoking the script
    if ($IncludeAll -or !($included.count)) {
        $included = @("UserMailbox","SharedMailbox","RoomMailbox","EquipmentMailbox","SchedulingMailbox","DiscoveryMailbox","TeamMailbox")
        $IncludeGroupMailboxes = $true
        $IncludeDGs = $true
    }

    $MBList = @();
    $MBList += (Get-Mailbox -ResultSize Unlimited -Filter {GrantSendOnBehalfTo -ne $null} -RecipientTypeDetails $included | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails,GrantSendOnBehalfTo)
    if ($IncludeGroupMailboxes) { $MBList += Get-UnifiedGroup -ResultSize Unlimited -Filter {GrantSendOnBehalfTo -ne $null} | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails,GrantSendOnBehalfTo }
    if ($IncludeDGs) { 
        $MBList += Get-DistributionGroup -ResultSize Unlimited -Filter {GrantSendOnBehalfTo -ne $null} | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails,GrantSendOnBehalfTo
        $MBList += Get-DynamicDistributionGroup -ResultSize Unlimited -Filter {GrantSendOnBehalfTo -ne $null} | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails,GrantSendOnBehalfTo
    }

    #If no recipients returned from the above cmdlet, stop the script and inform the user.
    if (!$MBList) { Write-Error "No recipients with Send on behalf of permissions found, specify different criteria." -ErrorAction Stop}


    $arrPermissions = @();

    #The list of recipients was already gathered along with all the needed data, prepare it for output
    foreach ($MB in $MBList) {
        #Handle objects with multiple SOB permission entries
        foreach ($entry in $MB.GrantSendOnBehalfTo) {
            $objPermissions = New-Object PSObject
            $i++;Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Number" -Value $i
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "GrantSendOnBehalfTo" -Value $entry
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Recipient Address" -Value $MB.PrimarySmtpAddress
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Recipient type" -Value $MB.RecipientTypeDetails
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Access Rights" -Value "Send on Behalf of"
            $arrPermissions += $objPermissions 
        }
    }

    #Output the result to the console host
    $arrPermissions | select 'Recipient Address','Recipient Type',GrantSendOnBehalfTo,'Access Rights'
}

#Invoke the Get-SOBPermissionInventory function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
Get-SOBPermissionInventory @PSBoundParameters -OutVariable global:varPermissions | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_SOBPermissions.csv" -NoTypeInformation
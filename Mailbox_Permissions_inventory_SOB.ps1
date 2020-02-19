param([switch]$IncludeAll,[switch]$IncludeUserMailboxes,[switch]$IncludeSharedMailboxes,[switch]$IncludeRoomMailboxes,[switch]$IncludeDiscoveryMailboxes,[switch]$IncludeGroupMailboxes,[switch]$IncludeTeamMailboxes,[switch]$IncludeDGs)

function Get-SOBPermissionInventory {
<#
.Synopsis
    Lists all Send on behalf of permissions for all recipients of the selected type(s).
.DESCRIPTION
    The Get-SOBPermissionInventory cmdlet lists all recipients of the selected type(s) that have at least one Send on Behalf of delegate added. Running the cmdlet without parameters will return entries for all User, Shared, Room, Equipment, Discovery, Team and Group mailboxes, as well as Distribution and Mail-enabled Security groups.
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
    #Specify whether to include Room and Equipment mailboxes in the result
    [Switch]$IncludeRoomMailboxes,
    #Specify whether to include Discovery mailboxes in the result
    [Switch]$IncludeDiscoveryMailboxes,
    #Specify whether to include Group mailboxes in the result
    [Switch]$IncludeGroupMailboxes,#Get-Mailbox will not return Team or Group mailboxes, needs special treatment
    #Specify whether to include Team mailboxes in the result
    [Switch]$IncludeTeamMailboxes,#Get-Mailbox will not return Team or Group mailboxes, needs special treatment
    #Specify whether to include Distribution groups in the result (includes DGs, Mail-enabled SGs, Room Lists)
    [switch]$IncludeDGs,
    #Specify whether to include every possible type of recipients in the result
    [Switch]$IncludeAll)


    #Initialize the variable used to designate recipient types, based on the script parameters
    $included = @();
    if($IncludeUserMailboxes) { $included += "UserMailbox"}
    if($IncludeSharedMailboxes) { $included += "SharedMailbox"}
    if($IncludeRoomMailboxes) { $included += "RoomMailbox"; $included += "EquipmentMailbox"}
    if($IncludeDiscoveryMailboxes) { $included += "DiscoveryMailbox"}
    if($IncludeTeamMailboxes) { $included += "TeamMailbox"}
    if($IncludeGroupMailboxes) { $included += "GroupMailbox"} #Used to prevent empty $included
    if($IncludeDGs) { $included += "GroupMailbox"} #Used to prevent empty $included

    #Confirm connectivity to Exchange Online
    try { Get-PSSession -InstanceId (Get-AcceptedDomain | select -First 1).RunspaceId.Guid -ErrorAction Stop | Out-Null }
    catch { Write-Error "No active Exchange Online session detected, please connect to ExO first: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx" -ErrorAction Stop }

    #Get the list of recipients with SOB permissions configured, depending on the parameters specified when invoking the script
    $MBList = @();
    if ($IncludeAll -or !$included) {
        $MBList += (Get-Mailbox -ResultSize Unlimited -Filter {GrantSendOnBehalfTo -ne $null} -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox,DiscoveryMailbox,TeamMailbox | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails,GrantSendOnBehalfTo)
        $MBList += (Get-UnifiedGroup -ResultSize Unlimited -Filter {GrantSendOnBehalfTo -ne $null} | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails,GrantSendOnBehalfTo) #workaround to include Group Mailboxes
        $MBList += (Get-DistributionGroup -ResultSize Unlimited -Filter {GrantSendOnBehalfTo -ne $null} | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails,GrantSendOnBehalfTo) #workaround to include DGs
        }
        
    else {
        $MBList += Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails $included -Filter {GrantSendOnBehalfTo -ne $null} | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails,GrantSendOnBehalfTo
        if ($IncludeGroupMailboxes) { $MBList += Get-UnifiedGroup -ResultSize Unlimited -Filter {GrantSendOnBehalfTo -ne $null} | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails,GrantSendOnBehalfTo }
        if ($IncludeDGs) { $MBList += Get-DistributionGroup -ResultSize Unlimited -Filter {GrantSendOnBehalfTo -ne $null} | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails,GrantSendOnBehalfTo }
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
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "User" -Value $entry
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Recipient Address" -Value $MB.PrimarySmtpAddress
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Recipient type" -Value $MB.RecipientTypeDetails
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Access Rights" -Value "Send on Behalf of"
            $arrPermissions += $objPermissions 
        }
    }

    #Output the result to the console host
    $arrPermissions | select User,'Recipient Address','Recipient Type','Access Rights'
}

#Invoke the Get-SOBPermissionInventory function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
Get-SOBPermissionInventory @PSBoundParameters -OutVariable global:varPermissions # | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_SOBPermissions.csv" -NoTypeInformation
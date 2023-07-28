param([switch]$IncludeAll,[switch]$IncludeUserMailboxes,[switch]$IncludeSharedMailboxes,[switch]$IncludeRoomMailboxes,[switch]$IncludeDiscoveryMailboxes,[switch]$IncludeTeamMailboxes)

function Get-MailboxPermissionInventory {
<#
.Synopsis
    Lists all non-default permissions for all mailboxes of the selected type(s).
.DESCRIPTION
    The Get-MailboxPermissionInventory cmdlet lists all mailboxes of the selected type(s) that have at least one object with non-default permissions added. Running the cmdlet without parameters will return entries for all User, Shared, Room, Equipment, Discovery, and Team mailboxes.
    Specifying particular mailbox type(s) can be done with the corresponding parameter. To specify a variable in which to hold the cmdlet output, use the OutVariable parameter.

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
    #Specify whether to include Room and Equipment mailboxes in the result
    [Switch]$IncludeRoomMailboxes,
    #Specify whether to include Discovery mailboxes in the result
    [Switch]$IncludeDiscoveryMailboxes,
    #Specify whether to include Team mailboxes in the result
    [Switch]$IncludeTeamMailboxes,
    #Specify whether to include every type of mailbox in the result
    [Switch]$IncludeAll)

    #Initialize the variable used to designate recipient types, based on the script parameters
    $included = @()
    if($IncludeUserMailboxes) { $included += "UserMailbox"}
    if($IncludeSharedMailboxes) { $included += "SharedMailbox"}
    if($IncludeRoomMailboxes) { $included += "RoomMailbox"; $included += "EquipmentMailbox"}
    if($IncludeDiscoveryMailboxes) { $included += "DiscoveryMailbox"}
    if($IncludeTeamMailboxes) { $included += "TeamMailbox"}

    #Confirm connectivity to Exchange Online
    try { $session = Get-PSSession -InstanceId (Get-AcceptedDomain | select -First 1).RunspaceId.Guid -ErrorAction Stop  }
    catch { Write-Error "No active Exchange Online session detected, please connect to ExO first: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx" -ErrorAction Stop }

    #Get the list of mailboxes, depending on the parameters specified when invoking the script
    if ($IncludeAll -or !$included) {
        $MBList = Invoke-Command -Session $session -ScriptBlock { Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox,DiscoveryMailbox,TeamMailbox | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails }
    }

    else {
        $MBList = Invoke-Command -Session $session -ScriptBlock { Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails $Using:included | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails }
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
        #The User property now returns UPN, no need for additional handling
        $MBrights = Get-MailboxPermission -Identity $MB.PrimarySmtpAddress.ToString() | ? {($_.User -ne "NT AUTHORITY\SELF") -and ($_.IsInherited -ne $true)} # -and ($_.AccessRights -match “FullAccess”) -and -not ($_.User -like "S-1-5*")}
        #No non-default permissions found, continue to next mailbox
        if (!$MBrights) { continue }

        foreach ($entry in $MBrights) {
            #Prepare the output object
            $objPermissions = New-Object PSObject
            $i++;Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Number" -Value $i
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "User" -Value $entry.user
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox address" -Value $MB.PrimarySmtpAddress
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Mailbox type" -Value $MB.RecipientTypeDetails
            Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Access Rights" -Value ($entry.AccessRights -join ",")
            #Uncomment if the script is failing due to connectivity issues, the line below will write the output to a CSV file for each individual permissions entry
            #$objPermissions | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd'))_MailboxPermissions.csv" -Append -NoTypeInformation
            $arrPermissions += $objPermissions
        }
    }

    #Output the result to the console host. Rearrange/sort as needed.
    #Maybe handle empty object?
    $arrPermissions | select User,'Mailbox address','Mailbox type','Access Rights'
}

#Invoke the Get-MailboxPermissionInventory function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
Get-MailboxPermissionInventory @PSBoundParameters -OutVariable global:varPermissions #| Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxPermissions.csv" -NoTypeInformation
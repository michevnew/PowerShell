param([switch]$IncludeAll,[switch]$IncludeUserMailboxes,[switch]$IncludeSharedMailboxes,[switch]$IncludeRoomMailboxes,[switch]$IncludeDiscoveryMailboxes,[switch]$IncludeTeamMailboxes,[switch]$CheckInboxRules,[switch]$CheckCalendarDelegates,[switch]$CheckTransportRules)

function Get-MailboxForwardingInventory {
<#
.Synopsis
    Lists forwarding settings for all mailboxes of the selected type(s).
.DESCRIPTION
    The Get-MailboxForwardingInventory cmdlet lists all mailboxes of the selected type(s) that have at least one form of forwarding configured. Running the cmdlet without parameters will return entries for all User, Shared, Room, Equipment, Discovery, and Team mailboxes.
    Specifying particular mailbox type(s) can be done with the corresponding parameter. To specify a variable in which to hold the cmdlet output, use the OutVariable parameter.

.EXAMPLE
    Get-MailboxForwardingInventory -IncludeUserMailboxes

    This command will return a list of user mailboxes that have at least one forwarding address configured, along with the all the relevant information.

.EXAMPLE
    Get-MailboxForwardingInventory -IncludeAll -OutVariable var
    $var | Export-Csv -NoTypeInformation "accessrights.csv"

    To export the results to a CSV file, use the OutVariable parameter.
.INPUTS
    None.
.OUTPUTS
    Array with information about the mailbox, forwarding address and type of forwarding configured.
#>

    [CmdletBinding()]
    
    Param
    (
    #Specify whether to check Inbox rules for forwarding
    [Switch]$CheckInboxRules,
    #Specify whether to check for Calendar Delegates
    [Switch]$CheckCalendarDelegates,
    #Specify whether to check Transport rules for forwarding
    [Switch]$CheckTransportRules,
    #Specify whether to include user mailboxes in the result
    [Switch]$IncludeUserMailboxes,    
    #Specify whether to include Shared mailboxes in the result
    [Switch]$IncludeSharedMailboxes,
    #Specify whether to include Room and Equipment mailboxes in the result
    [Switch]$IncludeRoomMailboxes,
    #Specify whether to include Discovery mailboxes in the result
    [Switch]$IncludeDiscoveryMailboxes,
    #Specify whether to include Team mailboxes in the result
    [Switch]$IncludeTeamMailboxes,#
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
        #Filterable, but if we are going to include all the methods, we need to cycle all mailboxes anyway
        if (!$CheckInboxRules -and !$CheckCalendarDelegates) { $MBList = Get-Mailbox -Filter {ForwardingSmtpAddress -ne $null -or ForwardingAddress -ne $null} -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox,DiscoveryMailbox,TeamMailbox | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails,ForwardingAddress,ForwardingSMTPAddress,DeliverToMailboxAndForward }
        else { $MBList = Invoke-Command -Session $session -ScriptBlock { Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox,DiscoveryMailbox,TeamMailbox | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails,ForwardingAddress,ForwardingSMTPAddress,DeliverToMailboxAndForward } }
    }

    else {
        if (!$CheckInboxRules -and !$CheckCalendarDelegates) { $MBList = Get-Mailbox -Filter {ForwardingSmtpAddress -ne $null -or ForwardingAddress -ne $null} -ResultSize Unlimited -RecipientTypeDetails $included | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails,ForwardingAddress,ForwardingSMTPAddress,DeliverToMailboxAndForward }
        else { $MBList = Invoke-Command -Session $session -ScriptBlock { Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails $Using:included | Select-Object -Property Displayname,Identity,PrimarySMTPAddress,RecipientTypeDetails,ForwardingAddress,ForwardingSMTPAddress,DeliverToMailboxAndForward } }
    }
    
    #If no mailboxes are returned from the above cmdlet, stop the script and inform the user.
    if (!$MBList) { Write-Error "No matching mailboxes found, specify different criteria." -ErrorAction Continue }

    #Once we have the mailbox list, cycle over each mailbox to gather forwarding, rules, calendar processing...
    #Remember to use ToString() as Invoke-Command returns actual object types!
    $arrForwarding = @()
    $count = 1; $PercentComplete = 0;
    foreach ($MB in $MBList) {
        #Progress message, will only be visible if processing rules/calendar delegates
        $ActivityMessage = "Retrieving data for mailbox $($MB.Identity.ToString()). Please wait..."
        $StatusMessage = ("Processing {0} of {1}: {2}" -f $count, @($MBList).count, $MB.PrimarySmtpAddress.ToString())
        $PercentComplete = ($count / @($MBList).count * 100)
        Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete -Verbose
        $count++
                
        #Gather forwarding configuration for each mailbox. 
        if ($MB.ForwardingAddress -or $MB.ForwardingSmtpAddress) {
            #Prepare the output object
            $objForwarding = New-Object PSObject
            $i++;Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Number" -Value $i
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Forwarding via" -Value ((&{If($MB.ForwardingAddress) {"ForwardingAddress attribute;"}}) + (&{If($MB.ForwardingSmtpAddress) {"ForwardingSmtpAddress attribute"}})).TrimEnd(";")
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Forwarding to" -Value ((&{If($MB.ForwardingAddress) {$MB.ForwardingAddress.ToString() + ";"}}) + (&{If($MB.ForwardingSmtpAddress) {$MB.ForwardingSmtpAddress.ToString().Split(":")[1] }})).TrimEnd(";")
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Keep original message" -Value (&{If($MB.DeliverToMailboxAndForward) {"True"} Else {"False"}})
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Mailbox address" -Value $MB.PrimarySmtpAddress.ToString()
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Mailbox type" -Value $MB.RecipientTypeDetails
            $arrForwarding += $objForwarding 
        }

        #check for Inbox rules
        if ($CheckInboxRules) {
            $varRules = Get-InboxRule -Mailbox $MB.PrimarySmtpAddress.ToString() | ? {$_.ForwardTo -ne $null -or $_.ForwardAsAttachmentTo -ne $null -or $_.RedirectTo -ne $null} | Select Name,ForwardTo,ForwardAsAttachmentTo,RedirectTo
            foreach ($rule in $varRules) {
                $objForwarding = New-Object PSObject
                $i++;Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Number" -Value $i
                Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Forwarding via" -Value "Inbox rule: $($rule.Name)"
                Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Forwarding to" -Value ((&{If($rule.ForwardTo) {$rule.ForwardTo}}) + (&{If($rule.ForwardAsAttachmentTo) {$rule.ForwardAsAttachmentTo}}) + (&{If($rule.RedirectTo) {$rule.RedirectTo}}))
                Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Keep original message" -Value (&{If($rule.RedirectTo) {"False"} Else {"True"}})
                Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Mailbox address" -Value $MB.PrimarySmtpAddress.ToString()
                Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Mailbox type" -Value $MB.RecipientTypeDetails
                $arrForwarding += $objForwarding
        }}

        #check for Calendar delegates
        if ($CheckCalendarDelegates) { 
        #In PowerShell, ResourceDelegates is only configurable for Resource mailboxes now! However users can still configure it via Outlook's Delegate settings.
        $varCaldelegates = Get-CalendarProcessing -Identity $MB.PrimarySmtpAddress.ToString() | select ResourceDelegates,ForwardRequestsToDelegates
        foreach ($delegate in $varCaldelegates.ResourceDelegates) {
            $objForwarding = New-Object PSObject
            $i++;Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Number" -Value $i
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Forwarding via" -Value "Calendar delegation"
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Forwarding to" -Value $delegate.ToString()
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Keep original message" -Value "N/A" #(&{If($delegate.ForwardRequestsToDelegates) {"True"} Else {"False"}}) #need to check with EWS
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Mailbox address" -Value $MB.PrimarySmtpAddress.ToString()
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Mailbox type" -Value $MB.RecipientTypeDetails
            $arrForwarding += $objForwarding
        }}
    }

    #Check for Transport rules
    if ($CheckTransportRules) {
        $varTRules = Get-TransportRule | ? {$_.RedirectMessageTo -ne $null -or $_.BlindCopyTo -ne $null -or $_.AddToRecipients -ne $null -or $_.CopyTo -ne $null -or $_.AddManagerAsRecipientType -ne $null} # just get all... | Select Name,CopyTo,BlindCopyTo,RedirectMessageTo,AddManagerAsRecipientType,AddToRecipients,ManagerAddresses
        foreach ($Trule in $varTRules) {
            $objForwarding = New-Object PSObject
            $i++;Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Number" -Value $i
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Forwarding via" -Value "Transport rule: $($Trule.Name)"
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Forwarding to" -Value ((&{If($Trule.RedirectMessageTo) {$Trule.RedirectMessageTo}}) + (&{If($Trule.BlindCopyTo) {$Trule.BlindCopyTo}}) + (&{If($Trule.AddToRecipients) {$Trule.AddToRecipients}}) + (&{If($Trule.CopyTo) {$Trule.CopyTo}}) + (&{If($Trule.AddManagerAsRecipientType) {$Trule.ManagerAddresses}}))
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Keep original message" -Value (&{If($Trule.RedirectMessageTo) {"False"} Else {"True"}})
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Mailbox address" -Value "N/A (via $($Trule.Conditions.Split(".")[-1]))"
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Mailbox type" -Value "N/A"
            $arrForwarding += $objForwarding
    }}
            
    #Output the result to the console host
    $arrForwarding | select 'Mailbox address','Mailbox type','Keep original message','Forwarding via','Forwarding to'
}


#Invoke the Get-MailboxForwardingInventory function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
Get-MailboxForwardingInventory @PSBoundParameters -OutVariable global:varForwarding # | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxForwarding.csv" -NoTypeInformation
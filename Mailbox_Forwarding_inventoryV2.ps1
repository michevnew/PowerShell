#Requires -Version 3.0
#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.0.0" }
[CmdletBinding()]
param([switch]$IncludeAll,[switch]$IncludeUserMailboxes,[switch]$IncludeSharedMailboxes,[switch]$IncludeRoomMailboxes,[switch]$CheckInboxRules,[switch]$CheckCalendarDelegates,[switch]$CheckTransportRules,[switch]$CheckTenantControls)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/4466/report-on-microsoft-365-mailbox-forwarding-all-methods-via-powershell

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
    #Specify whether to check tenant-wide forwarding controls
    [Switch]$CheckTenantControls,
    #Specify whether to include user mailboxes in the result
    [Switch]$IncludeUserMailboxes,
    #Specify whether to include Shared mailboxes in the result
    [Switch]$IncludeSharedMailboxes,
    #Specify whether to include Room, Equipment and Booking mailboxes in the result
    [Switch]$IncludeRoomMailboxes,
    #Specify whether to include every type of mailbox in the result
    [Switch]$IncludeAll)

    #Initialize the variable used to designate recipient types, based on the script parameters
    $included = @()
    if($IncludeUserMailboxes) { $included += "UserMailbox" }
    if($IncludeSharedMailboxes) { $included += "SharedMailbox" }
    if($IncludeRoomMailboxes) { $included += "RoomMailbox"; $included += "EquipmentMailbox"; $included += "SchedulingMailbox" }

    #Confirm connectivity to Exchange Online.
    Write-Verbose "Connecting to Exchange Online..."
    try { Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null }
    catch {
        try { Connect-ExchangeOnline -CommandName Get-InboxRule,Get-TransportRule,Get-CalendarProcessing,Get-HostedOutboundSpamFilterPolicy,Get-HostedOutboundSpamFilterRule,Get-RemoteDomain,Get-AcceptedDomain -SkipLoadingFormatData } #needs to be non-REST cmdlet
        catch { Write-Error "No active Exchange Online session detected. To connect to ExO: https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps";return }
    }

    if ($CheckTenantControls) {
        Write-Verbose "Processing Tenant's forwarding controls..."

        #Check for Enabled flag? Though it returns False for the Default policy...
        $varOutboundPolicy = Get-HostedOutboundSpamFilterPolicy | ? { $_.AutoForwardingMode -eq "On"} | select Name,IsDefault,Enabled,AutoForwardingMode
        if ($varOutboundPolicy) { #at least one policy has external forwarding allowed
            Write-Host "ATTENTION: External forwarding is allowed in the following Outbound spam filter policies:" -ForegroundColor Red
            Write-Host ($varOutboundPolicy.Name -join ",") -ForegroundColor Red
        }
        else { Write-Host "External forwarding is blocked." -ForegroundColor DarkGreen }

        $varRemoteDomains = Get-RemoteDomain | ? {$_.AutoForwardEnabled} | select DomainName,AutoForwardEnabled
        if ($varRemoteDomains) { #at least one Remote domain has external forwarding enabled
            Write-Host "ATTENTION: External forwarding is allowed in Remote domain settings for the followin domains:" -ForegroundColor Red
            Write-Host ($varRemoteDomains.DomainName -join ",") -ForegroundColor Red
        }
        else { Write-Host "External forwarding is blocked for all remote domains." -ForegroundColor DarkGreen }

        #Use the set of accepted domains as a simple check for internal/external
        $varAcceptedDomains = Get-AcceptedDomain | select -ExpandProperty DomainName
    }

    #Get the list of mailboxes, depending on the parameters specified when invoking the script
    Write-Verbose "Obtaining the list of mailboxes..."
    if ($IncludeAll -or !$included) {
        $included = @("UserMailbox","SharedMailbox","RoomMailbox","EquipmentMailbox","SchedulingMailbox","DiscoveryMailbox","TeamMailbox")
    }

    #Filterable, but if we are going to include all the methods, we need to cycle all mailboxes anyway
    if (!$CheckInboxRules -and !$CheckCalendarDelegates) { $MBList = Get-EXOMailbox -Filter {ForwardingSmtpAddress -ne $null -or ForwardingAddress -ne $null} -ResultSize Unlimited -RecipientTypeDetails $included -Properties ForwardingAddress,ForwardingSMTPAddress,DeliverToMailboxAndForward }
    else { $MBList = Get-ExOMailbox -ResultSize Unlimited -RecipientTypeDetails $included -Properties ForwardingAddress,ForwardingSMTPAddress,DeliverToMailboxAndForward }

    #If no mailboxes are returned from the above cmdlet, inform the user. Still cover Transport rules, so don't exist the script
    if (!$MBList) { Write-Error "No matching mailboxes found, specify different criteria." -ErrorAction Continue } #continue, as we might still need to cover Transport rules?

    #Once we have the mailbox list, cycle over each mailbox to gather forwarding, rules, calendar processing...
    $arrForwarding = @()
    $count = 1; $PercentComplete = 0;
    foreach ($MB in $MBList) {
        Write-Verbose "Processing mailbox $($MB.Identity) ..."
        #Progress message, will only be visible if processing rules/calendar delegates
        $ActivityMessage = "Retrieving data for mailbox $($MB.Identity). Please wait..."
        $StatusMessage = ("Processing {0} of {1}: {2}" -f $count, @($MBList).count, $MB.PrimarySmtpAddress)
        $PercentComplete = ($count / @($MBList).count * 100)
        Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete -Verbose
        $count++

        #Gather forwarding configuration for each mailbox.
        if ($MB.ForwardingAddress -or $MB.ForwardingSmtpAddress) {
            #Prepare the output object
            $objForwarding = New-Object PSObject
            $i++;Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Number" -Value $i
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Forwarding via" -Value ((&{If($MB.ForwardingAddress) {"ForwardingAddress attribute;"}}) + (&{If($MB.ForwardingSmtpAddress) {"ForwardingSmtpAddress attribute"}})).TrimEnd(";")
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Forwarding to" -Value ((&{If($MB.ForwardingAddress) {$MB.ForwardingAddress + ";"}}) + (&{If($MB.ForwardingSmtpAddress) {$MB.ForwardingSmtpAddress.Split(":")[1] }})).TrimEnd(";")
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Keep original message" -Value (&{If($MB.DeliverToMailboxAndForward) {"True"} Else {"False"}})
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Mailbox address" -Value $MB.PrimarySmtpAddress
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Mailbox type" -Value $MB.RecipientTypeDetails
            $arrForwarding += $objForwarding
        }

        #check for Inbox rules
        if ($CheckInboxRules) {
            Write-Verbose "Processing Inbox rules for mailbox $($MB.Identity) ..."
            Start-Sleep -Milliseconds 100
            $varRules = Get-InboxRule -Mailbox $MB.PrimarySmtpAddress -IncludeHidden | ? {$_.ForwardTo -ne $null -or $_.ForwardAsAttachmentTo -ne $null -or $_.RedirectTo -ne $null} | Select Name,ForwardTo,ForwardAsAttachmentTo,RedirectTo
            foreach ($rule in $varRules) {
                $objForwarding = New-Object PSObject
                $i++;Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Number" -Value $i
                Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Forwarding via" -Value "Inbox rule: $($rule.Name)"
                #Resolve this for internal? Also, delimiter
                Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Forwarding to" -Value ((&{If($rule.ForwardTo) {$rule.ForwardTo}}) + (&{If($rule.ForwardAsAttachmentTo) {$rule.ForwardAsAttachmentTo}}) + (&{If($rule.RedirectTo) {$rule.RedirectTo}}))
                Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Keep original message" -Value (&{If($rule.RedirectTo) {"False"} Else {"True"}})
                Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Mailbox address" -Value $MB.PrimarySmtpAddress
                Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Mailbox type" -Value $MB.RecipientTypeDetails
                $arrForwarding += $objForwarding
        }}

        #check for Calendar delegates
        if ($CheckCalendarDelegates) {
        Write-Verbose "Processing Calendar delegates for mailbox $($MB.Identity) ..."
        #In PowerShell, ResourceDelegates is only configurable for Resource mailboxes now! However users can still configure it via Outlook's Delegate settings.
        Start-Sleep -Milliseconds 100
        $varCaldelegates = Get-CalendarProcessing -Identity $MB.PrimarySmtpAddress | select ResourceDelegates,ForwardRequestsToDelegates
        foreach ($delegate in $varCaldelegates.ResourceDelegates) {
            $objForwarding = New-Object PSObject
            $i++;Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Number" -Value $i
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Forwarding via" -Value "Calendar delegation"
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Forwarding to" -Value $delegate
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Keep original message" -Value "N/A" #(&{If($delegate.ForwardRequestsToDelegates) {"True"} Else {"False"}}) #need to check with EWS
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Mailbox address" -Value $MB.PrimarySmtpAddress
            Add-Member -InputObject $objForwarding -MemberType NoteProperty -Name "Mailbox type" -Value $MB.RecipientTypeDetails
            $arrForwarding += $objForwarding
        }}
    }

    #Check for Transport rules
    if ($CheckTransportRules) {
        Write-Verbose "Processing Transport rules"
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

    #Output the result to the console host and CSV
    if ($arrForwarding) {
        $arrForwarding | % {
            if ($_.'Forwarding to') { #parse the forwarding addresses, try to find a matching internal recipient and against the list of accepted domains
                if ($CheckTenantControls) { #if we got the Accepted Domains data, try to determine if any forwarding addresses are internal/external
                        foreach ($entry in $_.'Forwarding to') {
                            $rec = $null
                            Start-Sleep -Milliseconds 100
                            if ($entry -match "EX:/") { $entry = $entry.Split("[")[1].Replace("]","").Replace("EX:","") } #fix for Inbox rules
                            if ($rec = Get-EXORecipient $entry -ErrorAction SilentlyContinue) { $entry = $rec.PrimarySmtpAddress.Split("@")[1] } #check if matching recipient is found internally
                            if ($entry -notmatch ($varAcceptedDomains -join "|")) { Add-Member -InputObject $_ -MemberType NoteProperty -Name "IsExternal" -Value $true; continue }
                        }
                        if (!$_.IsExternal) { Add-Member -InputObject $_ -MemberType NoteProperty -Name "IsExternal" -Value $false }
                    }
                else { Add-Member -InputObject $_ -MemberType NoteProperty -Name "IsExternal" -Value "N/A" } #else we don't know if external/internal
                $_.'Forwarding to' = ($_.'Forwarding to' -join ",") #fix for multiple values
            }
            else { Add-Member -InputObject $_ -MemberType NoteProperty -Name "IsExternal" -Value "N/A" }
        }

        Write-Verbose "Processing finished, outputing results ..."
        $arrForwarding | select 'Mailbox address','Mailbox type','Keep original message','Forwarding via','Forwarding to',IsExternal
        $arrForwarding | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxForwarding.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
        Write-Verbose "Finished. You can use the `$varForwarding global variable to play with the output before exporting."
    }
    else { Write-Verbose "Output is empty, skipping the export to CSV file..." }
}

#Invoke the Get-MailboxForwardingInventory function and pass the command line parameters. Make sure the output is stored in a variable for reuse, even if not specified in the input!
Get-MailboxForwardingInventory @PSBoundParameters -OutVariable global:varForwarding # | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxForwarding.csv" -NoTypeInformation
#Requires -Version 3.0
#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.9.0" }
#Requires -Modules @{ ModuleName="Microsoft.Graph.Users"; ModuleVersion="2.37.0" }
#Requires -Modules @{ ModuleName="Microsoft.Graph.Calendar"; ModuleVersion="2.37.0" }

#For details on what the script does and how to run it, check: https://michev.info/blog/post/8254/primer-use-invoke-changemeetingorganizer-to-bulk-change-meeting-organizer-in-exchange-online

[CmdletBinding()] #Make sure we can use -Verbose
Param([Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$Identity, #Identity of the mailbox where the original meeting resides. Can use any valid Exchange identifier
      [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][mailaddress]$NewOrganizer #SMTP address of the new organizer. Must be a user mailbox in the same tenant as the original organizer
)

#Connect to Exchange Online and resolve the provided identifiers.

Connect-ExchangeOnline -UserPrincipalName user@domain.com -ShowBanner:$false -CommandName "Invoke-ChangeMeetingOrganizer" -ErrorAction Stop -Verbose:$false
#Connect-ExchangeOnline -CertificateThumbprint 12345678903B4AF7C86FEF2F2B453DF0EE9C61FF -AppId xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx -Organization tenant.onmicrosoft.com -ShowBanner:$false

$Mailbox = Get-ExOMailbox -Identity $Identity -ErrorAction Stop -Verbose:$false
if ((!$Mailbox) -or ($Mailbox.RecipientTypeDetails -ne "UserMailbox")) {
    Write-Error "The specified identity is not a a valid user mailbox."
    return
}

$NewOrganizerM = Get-ExOMailbox -Identity $NewOrganizer -ErrorAction Stop -Verbose:$false
if ((!$NewOrganizerM) -or ($NewOrganizerM.RecipientTypeDetails -ne "UserMailbox")) {
    Write-Error "The specified new organizer is not a a valid user mailbox."
    return
}

#Now that we have the proper identifiers, fetch all relevant events from the original organizer's mailbox.
#We will use the Get-MgUserEvent cmdlet to retrieve the events. This requires connecting to Microsoft Graph with the appropriate permissions.
#If running via delegated permissions, the user running the script must have access to the original organizer's mailbox in addition to the Calendars.Read permission.

Connect-MgGraph -Scopes "Calendars.Read","Calendars.Read.Shared" -ErrorAction Stop -Verbose:$false -NoWelcome
#Connect-MgGraph -Scopes "Calendars.Read" -ClientId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -TenantId tenant.onmicrosoft.com -CertificateThumbprint 12345678903B4AF7C86FEF2F2B453DF0EE9C61FF -NoWelcome -ErrorAction Stop -Verbose:$false

#Fetch all future events from the original organizer's calendar. We will filter for events that are organized by the user, not cancelled and that have at least one attendee(s).
$events = Get-MgUserEvent -UserId $Mailbox.UserPrincipalName -Filter "start/dateTime ge '$(Get-Date -Format 'yyyy-MM-ddTHH:mm:ssZ')' and isCancelled eq false and isOrganizer eq true" -All -ErrorAction Stop -Property "id,subject,createdDateTime,uid,iCalUId,attendees,organizer,isOnlineMeeting,onlineMeetingUrl,recurrence" -Verbose:$false
#Easier to filter attendees client-side
$events = $events | ? {$_.Attendees.EmailAddress.Address -ne $Mailbox.PrimarySmtpAddress} | ? { ($_.attendees.Count -gt 0) }

if (!$events) {
    Write-Output "No matching events found for the specified mailbox. The script will exit."
    return
}

#Change the organizer for each event
if (!(Get-Command Invoke-ChangeMeetingOrganizer -ErrorAction SilentlyContinue)) {
    Write-Error "The Invoke-ChangeMeetingOrganizer function is not available. Check permissions and rerun the script."
    return
}

foreach ($e in $events) {
    Write-Verbose "Changing organizer for event '$($e.subject)' (ID: $($e.id))"
    if ($e.isOnlineMeeting) {
        Write-Warning "Meeting '$($e.subject)' (ID: $($e.id) is an online meeting. The original online meeting URL will be preseved, and the new organizer might not have sufficient permissions to lead the event."
    }
    if ($e.recurrence.Pattern.Type) { #stupid SDK returns a recurrence object with a null pattern for non-recurring events, so we need to check for the type property
        Write-Warning "Meeting '$($e.subject)' (ID: $($e.id) is a recurring meeting. The organizer change will take effect starting from the next occurence. Consider using the -TransferSeriesStartDate parameter to specify a different date for the switchover."
    }

    Invoke-ChangeMeetingOrganizer -Identity $Mailbox.UserPrincipalName -EventId $e.id -NewOrganizer $NewOrganizerM.PrimarySmtpAddress -Verbose:$VerbosePreference
    Start-sleep -Milliseconds 500 #Add a small delay to avoid throttling
}

Write-verbose "Finished."
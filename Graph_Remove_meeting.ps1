#Requires -Version 3.0
# Make sure to fill in all the required variables before running the script
# Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Calendars.ReadWrite #Needed to fetch matching calendar events and remove them
#    Exchange.ManageAsApp #Needed for the Exchange Online REST API calls
#    The app should also be granted Exchange Online role. The default View-Only Recipients role should suffice, or Global Reader/Reports reader on Entra side.

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/6300/how-to-remove-meetings-from-all-microsoft-365-mailboxes-via-the-graph-api

[CmdletBinding(SupportsShouldProcess)] #Make sure we can use -WhatIf and -Verbose
Param(
    [switch]$Quiet, #Suppress all output
    [PSCustomObject]$MeetingObject, #Event object to process
    [ValidateNotNullOrEmpty()][String]$MeetingId, #The ID of the meeting to remove, must be passed along with at least one mailbox to process
    [ValidateNotNullOrEmpty()][String[]]$IncludeMailboxes, #Additional mailboxes to process, not necesarily listed as attendees. Also works with DGs :)
    [ValidateNotNullOrEmpty()][String[]]$ExcludeMailboxes, #Do not process the following mailboxes
    [switch]$ProcessAllMailboxes #Process all mailboxes
)

function Renew-Token {
    param(
    [ValidateNotNullOrEmpty()][string]$Service
    )

    #prepare the request
    $url = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'

    #Define the scope based on the service value provided
    if (!$Service -or $Service -eq "Graph") { $Scope = "https://graph.microsoft.com/.default" }
    elseif ($Service -eq "Exchange") { $Scope = "https://outlook.office365.com/.default" }
    else { Write-Error "Invalid service specified, aborting..." -ErrorAction Stop; return }

    $Scopes = New-Object System.Collections.Generic.List[string]
    $Scopes.Add($Scope)

    $body = @{
        grant_type = "client_credentials"
        client_id = $appID
        client_secret = $client_secret
        scope = $Scopes
    }

    try {
        $authenticationResult = Invoke-WebRequest -Method Post -Uri $url -Body $body -ErrorAction Stop -Verbose:$false
        $token = ($authenticationResult.Content | ConvertFrom-Json).access_token
    }
    catch { throw $_ }

    if (!$token) { Write-Error "Failed to aquire token!" -ErrorAction Stop; return }
    else {
        Write-Verbose "Successfully acquired Access Token for $service"

        #Use the access token to set the authentication header
        if (!$Service -or $Service -eq "Graph") { Set-Variable -Name authHeaderGraph -Scope Global -Value @{'Authorization'="Bearer $token";'Content-Type'='application/json'} -Confirm:$false -WhatIf:$false }
        elseif ($Service -eq "Exchange") {
            Set-Variable -Name authHeaderExchange -Scope Global -Value @{'Authorization'="Bearer $token";'Content-Type'='application/json'} -Confirm:$false -WhatIf:$false

            #Add additional headers for Exchange
            $authHeaderExchange["X-ResponseFormat"] = "json"
            $authHeaderExchange["Prefer"] = "odata.maxpagesize=1000"
            $authHeaderExchange["connection-id"] = $([guid]::NewGuid().Guid).ToString()
            $authHeaderExchange["X-AnchorMailbox"] = "UPN:SystemMailbox{bb558c35-97f1-4cb9-8ff7-d53741dc928c}@$($TenantID)"
        }
        else { Write-Error "Invalid service specified, aborting..." -ErrorAction Stop; return }
    }
}

function Check-ExORecipient {
    param(
        [Parameter(Mandatory=$true)][string]$Identity
    )

    #Use the REST endpoint
    $uri = "https://outlook.office365.com/adminapi/beta/$($TenantID)/Recipient(`'$Identity`')"
    try {
        $result = Invoke-WebRequest -Method Get -Uri $uri -Headers $authHeaderExchange -Verbose:$false -ErrorAction Stop #suppress the output
    }
    catch {
        Write-Verbose "Recipient not found: $Identity"
        return
    }

    return ($result.Content | ConvertFrom-Json)
}

function Get-AllMailboxes {

    #Use the REST endpoint
    $rMailboxes = @()
    $uri = "https://outlook.office365.com/adminapi/beta/$($TenantID)/Mailbox?RecipientTypeDetails=UserMailbox,SharedMailbox,RoomMailbox,EquipmentMailbox&`$top=1000"

    do {
        $result = Invoke-WebRequest -Method Get -Uri $uri -Headers $authHeaderExchange -Verbose:$false -ErrorAction Stop #suppress the output
        $result = ($result.Content | ConvertFrom-Json)
        $uri = $result.'@odata.nextLink'

        $rMailboxes += $result.Value #this will fail if we only get a single mailbox
    } while ($uri)

    if (!$rMailboxes -or ($rMailboxes.Count -eq 0)) { Write-Error "No mailboxes found, aborting..." -ErrorAction Stop; return }

    foreach ($r in $rMailboxes) {
        $rInfo = [ordered]@{
            Email = $r.PrimarySmtpAddress
            DisplayName = $r.DisplayName
            RecipientType = $r.RecipientTypeDetails
            ObjectId = $r.ExternalDirectoryObjectId
        }
        $Mailboxes[$r.PrimarySmtpAddress] = $rInfo
    }
    #return ($Mailboxes | Select-Object -Property PrimarySmtpAddress,ExternalDirectoryObjectId)
}

function Get-DGMember {
    param(
        [Parameter(Mandatory=$true)][string]$Identity
    )

    $body = @{
        CmdletInput = @{
            CmdletName="Get-DistributionGroupMember"
            Parameters=@{"Identity"=$Identity}
        }
    }

    $uri = "https://outlook.office365.com/adminapi/beta/$($TenantID)/InvokeCommand?`$select=PrimarySmtpAddress,DisplayName,RecipientTypeDetails,ExternalDirectoryObjectId"
    try {
        $result = Invoke-WebRequest -Method POST -Uri $uri -Headers $authHeaderExchange -Body ($body | ConvertTo-Json -Depth 5) -ContentType "application/json" -Verbose:$false -ErrorAction Stop #suppress the output
    }
    catch {
        Write-Verbose "Group not found: $Identity"
        return
    }

    return ($result.Content | ConvertFrom-Json).value
}

function Get-UGMember {
    param(
        [Parameter(Mandatory=$true)][string]$Identity
    )

    $body = @{
        CmdletInput = @{
            CmdletName="Get-UnifiedGroupLinks"
            Parameters=@{"Identity"=$Identity;"LinkType"="Member"}
        }
    }

    $uri = "https://outlook.office365.com/adminapi/beta/$($TenantID)/InvokeCommand?`$select=PrimarySmtpAddress"
    try {
        $result = Invoke-WebRequest -Method POST -Uri $uri -Headers $authHeaderExchange -Body ($body | ConvertTo-Json -Depth 5) -ContentType "application/json" -Verbose:$false -ErrorAction Stop #suppress the output
    }
    catch {
        Write-Verbose "Group not found: $Identity"
        return
    }

    return ($result.Content | ConvertFrom-Json).value.PrimarySmtpAddress
}

#Check each attendee entry and resolve it to unique recipient. Needed because Attendees property can contain alias instead of email address.
#Also needed to handle distribution groups and unified group - expand members and process each.
function Process-Attendees {
    param(
        [Parameter(Mandatory=$true)][ValidateNotNull()]$Attendees,
        [string[]]$ExcludeAttendees
    )

    foreach ($attendee in $Attendees) {
        #If we pass input from the helper functions, it's a proper object
        if ($attendee.PrimarySmtpAddress) { $email = $attendee.PrimarySmtpAddress }
        else { $email = $attendee }

        if ($ExcludeAttendees -and ($ExcludeAttendees -contains $email)) { continue } #Skip if excluded
        if ($Mailboxes[$email]) { continue } #Skip if already processed
        try {
            if ($attendee.RecipientTypeDetails) { $r = $attendee } #Skip if we're passing an object, we already have all the info we need
            else {
                $r = Check-ExORecipient -Identity $email -ErrorAction Stop -Verbose:$VerbosePreference
                if (($r.count -gt 1) -or ($Mailboxes.Values.Email -contains $r.PrimarySmtpAddress)) { continue } #Skip if multiple results or duplicate email
            }

            if (($r.RecipientTypeDetails -eq 'MailUniversalDistributionGroup') -or ($r.RecipientTypeDetails -eq 'MailUniversalSecurityGroup')) { #Expand distribution groups
                $rList = Get-DGMember -Identity $r.PrimarySmtpAddress -ErrorAction Stop
                Process-Attendees -Attendees $rList
            }
            elseif ($r.RecipientTypeDetails -eq 'GroupMailbox') { #Expand Unified groups
                $rList = Get-UGMember -Identity $r.PrimarySmtpAddress -ErrorAction Stop
                Process-Attendees -Attendees $rList
            }
            elseif ($r.RecipientTypeDetails -in @('UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox')) {
                $rInfo = [ordered]@{
                    Email = $r.PrimarySmtpAddress
                    DisplayName = $r.DisplayName
                    RecipientType = $r.RecipientTypeDetails
                    ObjectId = $r.ExternalDirectoryObjectId
                }
                $Mailboxes[$email] = $rInfo
            }
            else { continue } #not covering any other recipient type
        }
        catch {
            continue #skip on any error
        }
    }
}

function Find-Event {
    param(
        [Parameter(Mandatory=$true,ParameterSetName="ById")][Parameter(Mandatory=$true,ParameterSetName="NotById")][ValidateNotNullOrEmpty()][string]$Mailbox,
        [Parameter(Mandatory=$false,ParameterSetName="ById")][ValidateNotNullOrEmpty()][string]$MeetingId, #EWSId, can copy it from OWA
        [Parameter(Mandatory=$false,ParameterSetName="ById")][ValidateNotNullOrEmpty()][string]$MeetingUid, #UID
        [Parameter(Mandatory=$false,ParameterSetName="NotById")][string]$Subject,
        [Parameter(Mandatory=$false,ParameterSetName="NotById")][datetime]$StartDate,
        [Parameter(Mandatory=$false,ParameterSetName="NotById")][datetime]$EndDate
    )

    if ($MeetingId) {
        if (!$MeetingId.StartsWith("AAMkAG")) { Write-Error "Invalid ID value provided, aborting..." -ErrorAction Stop; return }
        #Works with URLEncoded values, too
        $uri = "https://graph.microsoft.com/beta/users/$Mailbox/events/$($MeetingId)?`$select=id,uid,createdDateTime,subject,isCancelled,start,end,isOrganizer,type,attendees,organizer"
        try {
            $res = Invoke-WebRequest -Uri $uri -Headers $authHeaderGraph -Method Get -ErrorAction Stop -Verbose:$false
            $events = $res.Content | ConvertFrom-Json
        }
        catch { Write-Error "Failed to fetch events, aborting..." -ErrorAction Stop; return }
    }
    else {#Else we use filter
        $filter = @()
        if ($MeetingUid) {
            if (!$MeetingUId.StartsWith("040000008200E00074C5B7101A82E008")) { Write-Error "Invalid UID value provided, aborting..." -ErrorAction Stop; return }
            $filter = "uid eq '$MeetingUid'"
        }
        else {
            if ($StartDate -xor $EndDate) { Write-Error "Both StartDate and EndDate must be provided" -ErrorAction Stop; return }
            if ($StartDate -or $EndDate) { $filter += "start/dateTime ge '$StartDate' and end/dateTime le '$EndDate'" }
            if ($Subject) {
                if ([uri]::unEscapeDataString($Subject) -ne $Subject) { $filter += "subject eq '$Subject'" }
                else { $filter += "subject eq '$([uri]::EscapeDataString($Subject))'" }
            }
        }
        $filter = $filter -join " and "

        try {
            #Use /BETA here as /V1.0 does not return the uid on $select... WTF Microsoft?!
            $uri = "https://graph.microsoft.com/beta/users/$Mailbox/events?`$filter=$filter&`$top=100&`$orderby=start/dateTime&`$select=id,uid,createdDateTime,subject,isCancelled,start,end,isOrganizer,type,attendees,organizer"
            $res = Invoke-WebRequest -Uri $uri -Headers $authHeaderGraph -Method Get -ErrorAction Stop -Verbose:$false
            $events = ($res.Content | ConvertFrom-Json).Value
        }
        catch { Write-Error "Failed to fetch events, aborting..." -ErrorAction Stop; return }
    }

    if ($events.count -gt 1) {
        Write-Warning "Multiple events found, please select the one to process:"
        $objEvent = $events | Out-GridView -Title "Select event to process" -OutputMode Single
        if (!$objEvent) { Write-Error "No event selected, aborting..." -ErrorAction Stop; return }
    }
    elseif ($events.count -eq 0) { Write-Warning "No events found, please specify different criteria"; return }
    else { $objEvent = $events[0] }

    #We need the UID to process matching events/instances, so abort if empty
    if (!$objEvent.uid) { Write-Error "Null UID returned, aborting..." -ErrorAction Stop } #This should not happen

    return $objEvent
}

#==========================================================================
#Main script starts here
#==========================================================================

#Variables to configure
$tenantID = "tenant.onmicrosoft.com" #Your tenant root domain. Please do not use a GUID instead, as we use the value for the X-AnchorMailbox header
$appID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #the GUID of your app. For best result, use app with Sites.ReadWrite.All scope granted.
$client_secret = "verylongsecurestring" #client secret for the app

Renew-Token -Service "Graph"
Renew-Token -Service "Exchange"

if (!$PSBoundParameters.Count) { return } #Useful when dot-sourcing the script

$Mailboxes = @{}
#Check the input parameters to get the list of mailboxes to process
if ($MeetingObject -and $MeetingObject.attendees) { $IncludeMailboxes += ($MeetingObject.attendees.emailAddress | select -ExpandProperty address) }

if (!$IncludeMailboxes -and !$ProcessAllMailboxes) {
    Write-Error "No mailboxes provided, aborting..." -ErrorAction Stop; return
}

if ($ProcessAllMailboxes) { Get-AllMailboxes }
else {
    $IncludeMailboxes += ($MeetingObject.attendees.emailAddress | select -ExpandProperty address)
    if ($ExcludeMailboxes) {
        Process-Attendees $IncludeMailboxes -ExcludeAttendees $ExcludeMailboxes
    }
    else { Process-Attendees $IncludeMailboxes }
}

if (!$Mailboxes -or $Mailboxes.Count -eq 0) { Write-Error "No mailboxes found, aborting..." -ErrorAction Stop; return }
Write-Verbose "Processing a total of $($Mailboxes.Count) mailboxes provided via input parameters"

#If MeetingObject was passed, we can skip the search
if ($MeetingObject) { $objEvent = $MeetingObject }
else { #Else we need to fetch one event instance first in order to get the attendee list
    if (!$MeetingId) { Write-Error "No meeting ID provided, aborting..." -ErrorAction Stop; return }
    $eventFound = $false

    foreach ($Mbox in $Mailboxes.GetEnumerator()) {#Loop over the set of mailboxes
        if ($eventFound) { continue } #Skip if already found a match

        #Use /BETA here as /V1.0 does not return the uid on $select... WTF Microsoft?!
        $objEvent = Find-Event -Mailbox $Mbox.Value.ObjectId -MeetingUid $MeetingId
        if ($objEvent) {
            Write-Verbose "Found matching event in mailbox: $($Mbox.Value.Email)"
            $eventFound = $true
            break
        }
    }
    if (!$objEvent) { Write-Error "No matching event found, aborting..." -ErrorAction Stop; return }

    #We now have an event, add any attendees not already in the mailboxes list. Also expand distribution group and unified group membership.
    if (!$objEvent.attendees) { Write-Verbose "No attendees found in event object, processing only the mailboxes provided with input" }
    else {
        Write-Verbose "Processing attendees found in the event object"
        Process-Attendees ($objEvent.attendees.emailAddress | select -ExpandProperty address) -ExcludeAttendees $ExcludeMailboxes
        Write-Verbose "Processing event removal for a total of $($Mailboxes.Count) mailboxes after expanding attendees list"
    }
}

$output = @()
#Loop over the set of mailboxes to remove the event
foreach ($Mbox in $Mailboxes.GetEnumerator()) {
    Write-Verbose "Processing mailbox: $($Mbox.Value.Email)"

    #Find event with matching UID in the mailbox #Maybe leverage Find-Event here?
    try {
        $uri = "https://graph.microsoft.com/beta/users/$($Mbox.Value.ObjectId)/events?`$filter=uid eq '$($objEvent.uid)'&`$top=1&`$select=id,uid"
        $res = Invoke-WebRequest -Uri $uri -Headers $authHeaderGraph -Method Get -ErrorAction Stop -Verbose:$false
        $cEvent = ($res.Content | ConvertFrom-Json).Value
    }
    catch { Write-Verbose $_.Exception.Message; $cEvent | Out-Default; continue } #Move to next mailbox if we fail to fetch events

    if (!$cEvent) {
        Write-Warning "No matching event found in mailbox: $($Mbox.Value.Email)"
        $output += @{"User" = $Mbox.Value.Email;"Result" = "NotFound"}
        continue
    }
    else {
        Write-Verbose "Found matching event in mailbox: $($Mbox.Value.Email), processing removal..."
        if ($PSCmdlet.ShouldProcess($($Mbox.Value.Email),"Remove event: '$($objEvent.subject)'")) {
            $uri = "https://graph.microsoft.com/v1.0/users/$($Mbox.Value.ObjectId)/events/$($cEvent.id)"
            try {
                #Maybe add check for organizer and skip if the current user is the organizer?
                Invoke-WebRequest -Method Delete -Uri $uri -Headers $authHeaderGraph -SkipHeaderValidation -Verbose:$false -ErrorAction Stop | Out-Null #suppress the output
                Write-Verbose "Successfully removed event from mailbox: $($Mbox.Value.Email)"
                $output += @{"User" = $Mbox.Value.Email;"Result" = "Success"}
            }
            catch {
                Write-Verbose "Failed to remove event from mailbox: $($Mbox.Value.Email)"
                Write-Verbose $_.Exception.Message
                $output += @{"User" = $Mbox.Value.Email;"Result" = "Failure"}
                continue
            }
        }
        else {
            Write-Verbose "Skipped removal of event from mailbox: $($Mbox.Value.Email)"
            $output += @{"User" = $Mbox.Value.Email;"Result" = "Skipped"}
        }
    }
}

if (!$Quiet -and !$WhatIfPreference) { $output | select User, Result } #Write output to the console unless the -Quiet parameter is used
$output | select User, Result | Export-Csv -Path "$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MeetingRemoval_$($objEvent.uid).csv" -NoTypeInformation -Encoding UTF8 -UseCulture -Confirm:$false -WhatIf:$false
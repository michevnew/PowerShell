#Requires -Version 3.0
#Make sure to fill in all the required variables before running the script
#Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Calendars.ReadWrite (for creating calendar events)

param(
    [string[]]$Mailboxes, #List of mailboxes to process
    $CsvFilePath #Path to the CSV file containing the events
)

#For details on what the script does and how to run it, check: https://michev.info/blog/post/8724/how-to-populate-outlook-holidays-via-the-graph-api

#Test the CSV file path, if it doesn't exist, abort the script
if (-not (Test-Path -Path $CsvFilePath)) {
    Write-Host "CSV file path '$CsvFilePath' does not exist. Aborting..."
    return
}

$holidays = Import-Csv -Path $CsvFilePath | ? { $null -ne $_.Subject -and $null -ne $_.StartDateTime -and $null -ne $_.EndDateTime }
if (-not $holidays -or $holidays.Count -eq 0) {
	throw "No holidays found in the CSV file. Aborting..."
}
else {
    Write-Verbose "Found $($holidays.Count) holidays in the CSV file. Proceeding..."
}

#Get an Access token. Make sure to fill in all the variable values here. Or replace with your own preferred method to obtain token.
$tenantId = "tenant.onmicrosoft.com"
$uri = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'
$clientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$client_secret = "verylongstring"

#Get an Access token for the Graph API
$Scopes = New-Object System.Collections.Generic.List[string]
$Scope = "https://graph.microsoft.com/.default"
$Scopes.Add($Scope)

$body = @{
    grant_type = "client_credentials"
    client_id = $clientId
    client_secret = $client_secret
    scope = $Scopes
}

try {
    $res = Invoke-WebRequest -Method Post -Uri $uri -Verbose:$false -Body $body -UseBasicParsing
    $token = ($res.Content | ConvertFrom-Json).access_token

    $authHeader = @{
       'Authorization'="Bearer $token"
    }}
catch { Write-Host "Failed to obtain token, aborting..." ; return }

#Process each mailbox
foreach ($mailbox in $Mailboxes) {
    Write-Verbose "Processing mailbox: $mailbox"
    $uri = "https://graph.microsoft.com/v1.0/users/$mailbox/events?`$top=1"

    try {
        $res = Invoke-RestMethod -Method Get -Uri $uri -Headers $authHeader -Verbose:$false -ErrorAction Stop
        Write-Verbose "Successfully connected to mailbox: $mailbox"
    }
    catch {
        #handle 401/403/404 here
        if ($_.Exception.Response.StatusCode -eq 403) {
            Write-Warning "Mailbox '$mailbox' is not accessible. Skipping..."
            continue
        }
        elseif ($_.Exception.Response.StatusCode -eq 401) {
            Write-Error "UnAuthorized error thrown, check the app registration and the permissions granted. Aborting..." -ErrorAction Stop
            return
        }
        elseif ($_.Exception.Response.StatusCode -eq 404) {
            Write-Warning "Mailbox '$mailbox' not found. Skipping..."
            continue
        }
        else {
            Write-Warning "Failed to connect to mailbox '$mailbox'. Error: $_"
            continue
        }
    }

    foreach ($holiday in $holidays) {
        #Basic validation of the holiday data
        if (!$holiday.Subject) { $holday.Subject = "Holiday" }
        if (!$holiday.StartDateTime) { Write-Warning "Holiday '$($holiday.Subject)' does not have a valid StartDateTime. Skipping..." ; continue }
        if (!$holiday.EndDateTime) { $holiday.EndDateTime = $holiday.StartDateTime }
        If ((Get-Date $holiday.StartDateTime) -ge (Get-Date $holiday.EndDateTime)) { $holiday.EndDateTime = (Get-Date $holiday.StartDateTime).AddDays(1) }

        #Create the event payload
        $eventPayload = @{
            subject = $holiday.Subject
            body = @{
                contentType = "HTML"
                content = $holiday.Subject
            }
            start = @{
                dateTime = (Get-Date $holiday.StartDateTime).ToString("yyyy-MM-ddT00:00:00")
                timeZone = "UTC"
            }
            end = @{
                dateTime = (Get-Date $holiday.EndDateTime).ToString("yyyy-MM-ddT00:00:00")
                timeZone = "UTC"
            }
            categories = @("Holiday")
            isAllDay = $true
            showAs = "free"
        }

        #Convert the payload to JSON
        $jsonPayload = $eventPayload | ConvertTo-Json -Depth 5

        #Make the API call to create the event
        try {
            Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/users/$mailbox/events" -Headers $authHeader -Body $jsonPayload -ContentType "application/json" -verbose:$false -ErrorAction Stop
            Write-Verbose "Event '$($holiday.Subject)' created successfully for mailbox '$mailbox'."
        }
        catch {
            Write-Verbose "Failed to create event '$($holiday.Subject)' for mailbox '$mailbox'. Error: $_"
        }
        Start-sleep -Milliseconds 50 #Add a small delay to avoid throttling
    }

    Start-sleep -Milliseconds 500 #Add a small delay to avoid throttling
}
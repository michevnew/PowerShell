#Requires -Version 3.0
#Make sure to fill in all the required variables before running the script
#Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    User.Read.All (for "resolving" input values)
#    MailboxSettings.Read (to confirm a mailbox exists)
#    Mail.ReadBasic.All (to enumerate messages across mailboxes)
#    Mail.Read (if you also want item size included)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/5913/reporting-on-mailbox-item-count-and-size-by-year-via-the-graph-api

[CmdletBinding()] #Make sure we can use -Verbose
Param([string[]]$Mailbox,[switch]$IncludeItemSize=$false,[switch]$CompactOutput=$false)

#==========================================================================
#Helper functions
#==========================================================================
function Process-Folder {

    Param(
        # The ID of the mailbox to query
        [Parameter(Mandatory=$true)]$Id,
        # The ID of the folder to query
        [Parameter(Mandatory=$true)]$folderId,
        # Whether to include item size
        [Parameter(Mandatory=$false)][switch]$IncludeItemSize)

    # First we find the oldest message in the folder
    try {
        $uri = "https://graph.microsoft.com/v1.0/users/$Id/mailFolders/$folderId/messages?`$top=1&`$orderby=createdDateTime asc&`$select=id,createdDateTime"
        $result = Invoke-WebRequest -Uri $uri -Method GET -Headers $authHeader -ErrorAction Stop -Verbose:$false
        $result = ($result.Content | ConvertFrom-Json).value

        #Apparently, we can get null reply here, even if the folder has messages (totalItemCount > 0). Gotta love Graph...
        if ($result) { $cutoffDate = Get-Date $result.createdDateTime }
        else { $cutoffDate = Get-Date }
    }
    catch [Microsoft.PowerShell.Commands.HttpResponseException] {
        if ($_.ErrorDetails.Message -match "Access to OData is disabled.") { Write-Host "ERROR: An application access policy is blocking access to mailbox $id" -ForegroundColor Red; return }
        elseif ($_.ErrorDetails.Message -match "doesn't belong to the targeted mailbox") { Write-Host "ERROR: The specified folder does not belong to mailbox $id, this should not happen..." -ForegroundColor Red; return }
        elseif ($_.ErrorDetails.Message -match "Access is denied.") { Write-Host "ERROR: Including item size requires Mailbox.Read permissions, make sure you grant them first!" -ForegroundColor Red; return }
        else { $_ ; return }
    }
    catch { $_ ; return }

    $output = [System.Collections.SortedList]::new() #proper sorted list that mimics hashtable behavior

    do {
        #Now we can query the Graph API for the messages in the folder, using the cutoff date
        $startDate = (Get-date -Year ($cutoffDate).Year -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0).ToString("yyyy-MM-ddTHH:mm:ssZ")
        $endDate = (Get-date -Year ($cutoffDate).AddYears(1).Year -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0).ToString("yyyy-MM-ddTHH:mm:ssZ")
        $itemSize = 0
        if ($IncludeItemSize) {
            $uri = "https://graph.microsoft.com/v1.0/users/$Id/mailFolders/$folderId/messages?`$top=999&`$filter=createdDateTime+ge+$startDate+and+createdDateTime+lt+$endDate&`$count=true&`$select=createdDateTime&`$expand=singleValueExtendedProperties%28%24filter%3DId eq %27LONG 0x0E08%27%29"
            do {
                $result = Invoke-WebRequest -Uri $uri -Method GET -Headers $authHeader -Verbose:$false
                $uri = ($result.Content | ConvertFrom-Json).'@odata.nextLink'
                $itemSize += (($result.Content | ConvertFrom-Json).value.singleValueExtendedProperties.value | measure -Sum).Sum
            } while ($uri)
        }
        else {
            $uri = "https://graph.microsoft.com/v1.0/users/$Id/mailFolders/$folderId/messages?`$top=1&`$filter=createdDateTime+ge+$startDate+and+createdDateTime+lt+$endDate&`$count=true&`$select=createdDateTime"
            $result = Invoke-WebRequest -Uri $uri -Method GET -Headers $authHeader -Verbose:$false
            $itemSize = 0
        }

        $output[$cutoffDate.Year] = $($result.Content | ConvertFrom-Json).'@odata.count'.ToString() + ":" + $itemSize
        $cutoffDate = $cutoffDate.AddYears(1)
    } while ($cutoffDate.Year -le (Get-Date).Year)

    return $output
}

#==========================================================================
#Main script starts here
#==========================================================================

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
    $res = Invoke-WebRequest -Method Post -Uri $uri -Verbose:$false -Body $body
    $token = ($res.Content | ConvertFrom-Json).access_token

    $authHeader = @{
       'Authorization'="Bearer $token"
    }}
catch { Write-Host "Failed to obtain token, aborting..." ; return }

#Prepare the list of mailboxes
Write-Verbose "Parsing the Mailbox parameter..."
$Mailboxes = @{}
foreach ($mb in $Mailbox) {
    #Make sure a matching user object is found and check for mailbox settings to confirm a mailbox exists
    try {
        $result = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/users/$mb" -Method GET -Headers $authHeader -ErrorAction Stop -Verbose:$false
        $SMTPAddress = ($result.Content | ConvertFrom-Json).Mail
        $result = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/users/$mb/mailboxSettings/userPurpose" -Method GET -Headers $authHeader -ErrorAction Stop -Verbose:$false

        if (!$result -or (($result.Content | ConvertFrom-Json).value -notmatch "user|shared|room")) { Write-Warning "Failed to get mailbox settings for $mb, make sure the user actually has a mailbox..." ; continue }
        elseif (($SMTPAddress.count -gt 1) -or ($Mailboxes[$mb]) -or ($Mailboxes.ContainsValue($SMTPAddress))) { Write-Warning "Multiple mailboxes matching the identifier $mb found, skipping..."; continue }
        else { $Mailboxes[$mb] = $SMTPAddress }
    }
    catch { Write-Warning "Failed to confirm mailbox exists for $mb, skipping..." ; continue }
}

if (!$Mailboxes -or ($Mailboxes.Count -eq 0)) { Throw "No matching mailboxes found, check the parameter values." }
Write-Verbose "The following list of mailboxes will be used: ""$($Mailboxes.Values -join ", ")"""

$output = [System.Collections.Generic.List[Object]]::new() #output variable
#Loop over all mailboxes and get the list of folders
foreach ($entry in $Mailboxes.Values) {
    Start-Sleep -Milliseconds 500 #add some delay to avoid throttling

    Write-Verbose "Processing mailbox ""$($entry)""..."
    # Query the Graph API to get all folders within the mailbox
    try {
        #We use /beta here, as it includes subfolders in the result - saves us from having to use recursion
        $folders = @()
        $uri = "https://graph.microsoft.com/beta/users/$entry/mailFolders?&includeHiddenFolders=true&`$top=999&`$filter=totalItemCount gt 0"

        do {
            $result = Invoke-WebRequest -Method GET -Uri $uri -Headers $authHeader -ErrorAction Stop -Verbose:$false
            $uri = ($result.Content | ConvertFrom-Json).'@odata.nextLink'

            $folders += ($result.Content | ConvertFrom-Json).Value
        } while ($uri)

    }
    catch { Write-Warning "Failed to get folders for $entry, skipping..." ; continue }

    # Process any matching (sub)folders
    foreach ($folder in $folders) {
        Start-Sleep -Milliseconds 50 #add some delay to avoid throttling
        Write-Verbose "Processing folder ""$($folder.displayName)"" within mailbox ""$($entry)""..."

        #Get the folder statistics
        $out = @()
        $out = Process-Folder $entry $folder.Id -IncludeItemSize:$IncludeItemSize

        #Prepare the output object
        if ($CompactOutput) {
            $i++;$objStats = [PSCustomObject][ordered]@{
                "Number" = $i
                "Mailbox" = $entry
                "Folder" = $folder.displayName
                "Folder item count" = $folder.totalItemCount
                "Folder size" = $folder.sizeInBytes
                "Folder stats" = $out.GetEnumerator().ForEach({ "$($_.Name)=$($_.Value)" }) -join ";"
                "Folder ID" = $folder.Id
            }
            $output.Add($objStats)
        }
        else {
            foreach ($key in $out.keys) {
                $i++;$objStats = [PSCustomObject][ordered]@{
                    "Number" = $i
                    "Mailbox" = $entry
                    "Folder" = $folder.displayName
                    "Folder item count" = $folder.totalItemCount
                    "Folder size" = $folder.sizeInBytes
                    "Year" = $key
                    "Item count" = $out[$key].Split(":")[0]
                    "Item size" = $out[$key].Split(":")[1]
                    "Folder ID" = $folder.Id
                }
                $output.Add($objStats)
            }
        }

        if ($VerbosePreference) {
            Write-Verbose "Folder stats for folder ""$($folder.displayName)"" within mailbox ""$($entry)"":"
            #$out | Out-Default
            Write-Verbose ($out.GetEnumerator().ForEach({$_ | select @{n="Year";e={$_.Name}},@{n="ItemCount";e={$_.Value.Split(":")[0]}},@{n="ItemSize";e={$_.Value.Split(":")[1]}}}) | Out-String )
        }
    }
}

#Export the result to CSV file
$output | select * -ExcludeProperty Number | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxMessageStatsPerYear.csv" -NoTypeInformation -Encoding UTF8 -UseCulture -Verbose:$false
Write-Verbose "Output exported to $($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxMessageStatsPerYear.csv"
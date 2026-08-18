#Requires -Version 3.0
#Make sure to fill in all the required variables before running the script (lines 370-372)
#Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    User.Read.All (needed for getting the mailbox identifier and sufficient for the translateExchangeIds call)
#    MailboxFolder.Read.All (needed for enumerating mailbox folders)

[CmdletBinding()]
Param([Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][Alias("Identity")][String]$Mailbox,[switch]$IncludeNonIPM = $false)

#For details on what the script does and how to run it, check: https://michev.info/blog/post/8282/convert-mailbox-folder-identifiers-non-interactively

#region Helper functions

#Obtain an access token or renew it if needed
function Renew-Token {

    #prepare the request
    $url = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'

    #Define the scope
    $Scope = "https://graph.microsoft.com/.default"
    $Scopes = New-Object System.Collections.Generic.List[string]
    $Scopes.Add($Scope)

    $body = @{
        grant_type = "client_credentials"
        client_id = $appID
        client_secret = $client_secret
        scope = $Scopes
    }

    try {
        $authenticationResult = Invoke-RestMethod -Method Post -Uri $url -Body $body -ErrorAction Stop -Verbose:$false
        $token = $authenticationResult.access_token
    }
    catch { throw $_ }

    if (!$token) { Write-Error "Failed to aquire token!" -ErrorAction Stop; return }
    else {
        Write-Verbose "Successfully acquired Access Token"
        #Use the access token to set the authentication header
        Set-Variable -Name authHeaderGraph -Scope Global -Value @{'Authorization'="Bearer $token";'Content-Type'='application/json'} -Confirm:$false
    }
}

#Function to handle errors
function Process-Error {
    param([Parameter(Mandatory)]$ErrorMessage)

    #Insufficient permissions granted to the service principal, terminate the script
    if (!$ErrorMessage.ErrorDetails.Message) { #ExO throws a 401 with no ErrorMessage... no way to differentiate token expiry from generirc permission-related issues or other errors
        if ($ErrorMessage.Exception.Message -match "Response status code does not indicate success: 401") { Write-Error "ERROR: Insufficient permissions to connect to Exchange Online. Verify correct permissions are assigned to the service principal!" -ErrorAction Stop }
    }
    if ($ErrorMessage.ErrorDetails.Message -match "InsufficientPermissionsException|Insufficient privileges to complete the operation|Authorization_RequestDenied|Authorization failed due to missing permission scope") { Write-Error "ERROR: Insufficient permissions to perform the removal operation. Verify correct permissions are assigned to the service principal!" -ErrorAction Stop }
    elseif ($ErrorMessage.ErrorDetails.Message -match "Access to OData is disabled") { Write-Error "ERROR: Access is blocked by an Application Access Policy, the script will now exit..." -ErrorAction Stop } #ExO
    elseif ($ErrorMessage.ErrorDetails.Message -match "The role assigned to application") { Write-Error "ERROR: Insufficient permissions to connect to Exchange Online. Verify the admin role(s) assigned to the service principal!" -ErrorAction Stop } #ExO
    #Token has expired, renew it and retry the operation
    #ExO throws a 401 with no ErrorMessage... no way to differentiate from generirc permission-related issues
    elseif ($ErrorMessage.ErrorDetails.Message -match "Lifetime validation failed, the token is expired|Access token has expired") {
        Write-Warning "Access token has expired, renewing it..."
        $global:authHeaderGraph = $null; $global:authHeaderExchange = $null

        Renew-Token

        if (!$authHeaderGraph) { Write-Error "Failed to renew token, aborting..." -ErrorAction Stop }
        if (!$authHeaderExchange) { Write-Error "Failed to renew token, aborting..." -ErrorAction Stop }
    }
    #The rest are non-terminal errors
    elseif ($ErrorMessage.ErrorDetails.Message -match "ManagementObjectNotFoundException|ADNoSuchObjectException|Couldn't find object|Unable to retrieve mailbox folder statistics for mailbox") { Write-Error "The specified object was not found, check the input values..." -ErrorAction Stop }
    elseif ($ErrorMessage.ErrorDetails.Message -match "Invalid object identifier|The requested user .* is invalid|does not exist or one of its queried reference-property|Unsupported referenced-object resource identifier") { Write-Error "The specified object was not found, check the input values..." -ErrorAction Stop }
    else { $ErrorMessage | fl * -Force; return } #catch-all for any unhandled errors
}

#Do we even need this? Both UPN and ExternalDirectoryObjectId can be used with /admin/Exchange/Mailboxes/
function GetMailbox {
<#
.Synopsis
    Uses Graph's List Exchange settings method to get the mailbox identifier
.DESCRIPTION
    The GetMailbox function uses the Graph API's List Exchange settings method to retrieve the mailbox identifier for the specified identity value.
.PARAMETER Identity
	Use the Identity parameter to designate the identifier of the user object. This can be a UPN or GUID.
.EXAMPLE
    GetMailbox user@domain.com
    This function will return the identifier for the user@domain.com mailbox.
.INPUTS
    GUID or UPN of the user object.
.OUTPUTS
    The unique identifier for the user's primary mailbox.
#>

param([Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$Identity) #Best use UPN

    $uri = "https://graph.microsoft.com/v1.0/users/$Identity/settings/exchange"
    try {
        $result = Invoke-RestMethod -Method GET -Uri $uri -Headers $authHeaderGraph -Verbose:$false -ErrorAction Stop
    }
    catch {
        Process-Error -ErrorMessage $_
    }

    if (!$result) { Write-Error "Failed to retrieve mailbox settings for the specified user. The script will exit." -ErrorAction Stop; return }

    return $result.primaryMailboxId
}

function ReturnFolderList {
<#
.Synopsis
    Enumerates all user-accessible folders for the mailbox.
.DESCRIPTION
    The ReturnFolderList cmdlet enumerates the folders for the given mailbox.
.PARAMETER MailboxID
	Use the MailboxID parameter to designate the mailbox where the desired folders reside.
.PARAMETER IncludeNonIPM
    Use the IncludeNonIPM switch to include folders from the non-IPM tree.
.EXAMPLE
    ReturnFolderList user@domain.com
    This command will return a list of all user-accessible folders for the user@domain.com mailbox.
.INPUTS
    Identifier for the mailbox, UPN or GUID.
.OUTPUTS
    Array with information about the mailbox folders.
#>

    param(
	[Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$MailboxID, #Best use UPN
    [switch]$IncludeNonIPM) #whether to include folders from the non-IPM tree

    if ($IncludeNonIPM) {  $uri = "https://graph.microsoft.com/v1.0/admin/exchange/mailboxes/$MailboxID/folders/Root/childFolders?includeHiddenFolders=true&`$top=9999" }
    else {  $uri = "https://graph.microsoft.com/v1.0/admin/exchange/mailboxes/$MailboxID/folders?`$top=9999&includeHiddenFolders=true" }

    try {
        $result = Invoke-RestMethod -Method Get -Uri $uri -Headers $authHeaderGraph -Verbose:$false -ErrorAction Stop #suppress the output
    }
    catch {
        Process-Error -ErrorMessage $_
    }
    $MBfolders = $result.value

    if (!$MBfolders -or $MBfolders.count -eq 0) { Write-Error "No folders found for the specified mailbox. The script will exit." -ErrorAction Stop; return }

    #Crawl through the folders and retrieve all child folders
    #This needs to run recursively, as the Graph API only returns the first level of child folders
    foreach ($folder in $MBfolders) {

        #Enumerate any child folders, recursively
        if ($folder.childFolderCount -gt 0) {
            Start-Sleep -Milliseconds 100 #just in case
            $script:hasChildFolders = $true
            $MBfolders += (GetChildFolders -MailboxID $MailboxID -ParentFolderId $folder.id -ParentFolderName $folder.displayName)
        }
    }

    return ($MBfolders | select @{n="Name";e={$_.displayName}},@{n="FolderType";e={$_.type}},@{n="Identity";e={$Mailbox + "\" + $_.Path + $_.displayName}},@{n="RestId";e={$_.Id}} | Sort Identity)
}

function GetChildFolders {
<#
.Synopsis
    Retrieves child folders for a given parent folder.
.DESCRIPTION
    The GetChildFolders cmdlet retrieves child folders for the specified parent folder in a mailbox.
.PARAMETER ParentFolderId
    Use the ParentFolderId parameter to designate the folderId of the parent folder.
.PARAMETER ParentFolderName
    Use the ParentFolderName parameter to designate the name of the parent folder, which is used for constructing the Path property.
.PARAMETER MailboxID
    Use the MailboxID parameter to designate the mailbox where the desired folders reside.
.EXAMPLE
    GetChildFolders -Mailbox user@domain.com -ParentFolderId LgAAAAChKSJAhlnUTIHt -ParentFolderName "Inbox"
    This command will return a list of child folders for the specified parent folder.
.INPUTS
    Folder Id value obtained from the Graph API.
.OUTPUTS
    Array with information about the child folders.
#>
    param(
        [Parameter(Mandatory=$true)]$ParentFolderId, #the Id of the parent folder
        [Parameter(Mandatory=$true)]$ParentFolderName, #the Name of the parent folder (used for constructing the Path property)
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$MailboxID #the mailbox identifier
    )

    $uri = "https://graph.microsoft.com/v1.0/admin/exchange/mailboxes/$MailboxID/folders/$ParentFolderId/childFolders?includeHiddenFolders=true&`$top=9999"

    try {
        $result = Invoke-RestMethod -Method Get -Uri $uri -Headers $authHeaderGraph -Verbose:$false -ErrorAction Stop
    }
    catch {
        Process-Error -ErrorMessage $_
    }

    if (!$result.value) { Write-Warning "Query returned no child folders for $ParentFolderId, whereas childFolderCount value indicates such exist..."; continue }
    $childResult += ($result.value | Select-Object -Property *,@{n="Path";e={$ParentFolderName + "\"}})

    #Trigger the recursive call to get child folders of the current folder
    foreach ($folder in $result.value) {
        if ($folder.childFolderCount -gt 0) {
            $script:hasChildFolders = $true

            while ($script:hasChildFolders) {
                $script:hasChildFolders = $false
                GetChildFolders -MailboxID $MailboxID -ParentFolderId $folder.id -ParentFolderName ($ParentFolderName + "\" + $folder.displayName)
            }
        }
        else { $script:hasChildFolders = $false }
    }
    $script:hasChildFolders = $false

    return $childResult
}

function EntryIdToFolderId {
<#
.Synopsis
    Transforms entryId value to folderId format.
.DESCRIPTION
    The EntryIdToFolderId cmdlet transforms the entryId value to the folderId format used by Get-ExOMailboxFolderStatistics.
.PARAMETER EntryId
	Use the EntryId parameter to designate the original entryId value.
.PARAMETER FolderType
    Use the FolderType parameter to designate the folder type, which is used to determine the suffix for the folderId.
.EXAMPLE
    EntryIdToFolderId AAAAAE5be7M8YWJHgbjwLM4A2vABAABUr5vvaXZKnE_vqg7vlfUAAKKTW3cAAA2
    This command will convert the given entryId value to the folderId format.
.INPUTS
    EntryId value.
.OUTPUTS
    The converted folderId value in base64 format.
#>

    param([Parameter(Mandatory=$true)]$entryId, [string]$FolderType)

    # split URL-safe payload and padding count suffix (last char is the digit 0-2)
    if ($entryId -notmatch '^(?<id>[A-Za-z0-9\-_]+?)(?<pad>[0-2])$') { throw "Invalid EntryId format." }
    $entryIdPayload = $matches.id
    $equalCharCount = [int]$matches.pad

    # restore regular base64 string
    $entryIdBase64 = $entryIdPayload.Replace('_','/').Replace('-','+') + ('=' * $equalCharCount)

    # convert from base64 to bytes
    $entryIdBytes = [Convert]::FromBase64String($entryIdBase64)

    # convert byte array to hex string
    $entryIdHexString = [System.BitConverter]::ToString($entryIdBytes).Replace('-','')

    # determine the suffix based on the folder type
    $suffix = switch -Wildcard ($FolderType) {
        'IPF.Note*' { "01" }
        'IPF.Appointment*' { "02" }
        'IPF.Contact*' { "03" }
        'IPF.Task*' { "04" }
        'IPF.StickyNote*' { "05" }
        'IPF.Journal*' { "06" }
        #'IPF.Note' { "07" } #SearchDiscoveryHoldsFolder, SearchDiscoveryHoldsUnindexedItemFolder, AllCategorizedItems, AllContacts, AllItems, AllTodoTasks...
        default { "01" }
    }

    # restore folderId framing bytes (0x2E prefix, 0x03 suffix)
    $folderIdHexString = "2E$entryIdHexString" + $suffix

    # convert to byte array - two chars represents one byte
    $folderIdBytes = [byte[]]::new($folderIdHexString.Length / 2)

    For($i=0; $i -lt $folderIdHexString.Length; $i+=2){
        $folderIdTwoChars = $folderIdHexString.Substring($i, 2)
        $folderIdBytes[$i/2] = [convert]::ToByte($folderIdTwoChars, 16)
    }

    # convert bytes to base64 string
    $folderId = [Convert]::ToBase64String($folderIdBytes)

    return $folderId
}

function EntryIdToEDiscoveryId {
<#
.Synopsis
    Transforms entryId value to eDiscoveryID format.
.DESCRIPTION
    The EntryIdToFolderId cmdlet transforms the entryId value to the format used by eDiscovery and Content Search.
.PARAMETER EntryId
	Use the EntryId parameter to designate the original entryId value.
.EXAMPLE
    EntryIdToEDiscoveryId AAAAAE5be7M8YWJHgbjwLM4A2vABAABUr5vvaXZKnE_vqg7vlfUAAKKTW3cAAA2
    This command will convert the given entryId value to the folderId format.
.INPUTS
    EntryId value.
.OUTPUTS
    The converted eDiscoveryID value.
#>

    param([Parameter(Mandatory=$true)]$entryId)

    # split URL-safe payload and padding count suffix (last char is the digit 0-2)
    if ($entryId -notmatch '^(?<id>[A-Za-z0-9\-_]+?)(?<pad>[0-2])$') { throw "Invalid EntryId format." }
    $entryIdPayload = $matches.id
    $equalCharCount = [int]$matches.pad

    # restore regular base64 string
    $entryIdBase64 = $entryIdPayload.Replace('_','/').Replace('-','+') + ('=' * $equalCharCount)

    # convert from base64 to bytes
    $entryIdBytes = [Convert]::FromBase64String($entryIdBase64)

    # convert byte array to hex string
    $entryIdHexString = [System.BitConverter]::ToString($entryIdBytes).Replace('-','')
    return $entryIdHexString.Substring(44)
}

function RestIdtoEntryId {
<#
.Synopsis
    Transforms RestId value to entryId format.
.DESCRIPTION
    The RestIdtoEntryId cmdlet transforms the RestId value obtained from the Graph API to the entryId format used by Exchange.
.PARAMETER Mailbox
    Use the Mailbox parameter to designate the mailbox where the folders reside. Mandatory. The calling user must have User.ReadBasic.All permissions for the call to succeed.
.PARAMETER Ids
    Use the Ids parameter to provide the RestId values to be converted. Multiple values can be provided as an array, but the total number of IDs in a single call must not exceed 1000.
.EXAMPLE
    RestIdtoEntryId -Mailbox user@domain.com -Ids $entryId1, $entryId2, $entryId3
    This command will convert the given RestId value(s) to the entryId format.
.INPUTS
    RestId value obtained from the Graph API and the mailbox identifier.
.OUTPUTS
    The converted entryId value.
#>
    param([Parameter(Mandatory=$true)][string[]]$Ids, #max 1000
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$Mailbox) #don't cast as SMTP, as the Graph expects GUID/UPN

    #Hash table to store the translated IDs
    $EntryIds = @{}

    do {
        $batch = $Ids | Select-Object -First 1000
        $Ids = $Ids | Select-Object -Skip 1000

        #Prepare the request body for the batch of IDs
        $params = @{
            inputIds = @($batch)
            sourceIdType = "restId"
            targetIdType = "entryId"
        }

        # Execute the request
        try {
           $res = Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$Mailbox/translateExchangeIds" -Headers $authHeaderGraph -Body ($params | ConvertTo-Json -Depth 5) -ContentType "application/json" -Verbose:$false -ErrorAction Stop
        }
        catch {
            Process-Error -ErrorMessage $_
        }

        foreach ($item in $res.value) {
            if ($item.TargetId -and !$item.ErrorDetails.Code) { $EntryIds[$item.SourceId] = $item.TargetId }
            else { Write-Warning "Failed to translate ID: $($item.SourceId). Error details: $($item.ErrorDetails.Code)"; $EntryIds[$item.SourceId] = "N/A"; continue }
        }
    } while ($Ids.Count -gt 0)

    return $EntryIds
}
#endregion

#==========================================================================
# Main script
#==========================================================================

#Variables to configure
$tenantID = "tenant.onmicrosoft.com" #your tenantID or tenant root domain
$appID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #the GUID of your app
$client_secret = "verylongsecurestring" #client secret for the app

Renew-Token

#Get the mailbox identifier for the specified Identity value
$MailboxID = GetMailbox -Identity $Mailbox #Do we need this? Both UPN and ExternalDirectoryObjectId can be used with /admin/Exchange/Mailboxes/

#Get the folder list, output contains FolderId, eDiscoveryId and EntryId
$temp = ReturnFolderList $MailboxID -IncludeNonIPM:$IncludeNonIPM

#Convert EntryId to RestId using the Graph API, store in a hash table for easy retrieval
$EntryIds = RestIdtoEntryId -Ids $temp.RestId -Mailbox $MailboxID

#Append the output with the EntryId values from the hash table
$output = $temp | select Name,FolderType,Identity,@{n="EntryId";e={$EntryIds[$_.RestId]}},RestId

$output = $output | select Name,FolderType,Identity,@{n="FolderId";e={EntryIdToFolderId -entryId $_.EntryId -FolderType $_.FolderType}},@{n="eDiscoveryId";e={EntryIdToEDiscoveryId -entryId $_.EntryId}},EntryId,RestId

#Export the output to CSV
$output | Select * | Export-Csv -Path "$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxFolderIDs.csv" -NoTypeInformation -Encoding UTF8
Write-host "CSV export completed: $($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxFolderIDs.csv" -ForegroundColor Green

#Generate a HTML export
$htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Mailbox Folder IDs for $Mailbox</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
        table { border-collapse: collapse; width: 100%; background: white; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        th { background: #0078d4; color: white; padding: 12px; text-align: left; cursor: pointer; user-select: none; }
        td { padding: 5px 10px; border-bottom: 1px solid #ddd; }
        tr:nth-child(even) { background: #f0f4fa; }
        tr:hover { background: #d0e7fa; }
        .button { display: inline-block; padding: 5px 10px; margin: 0 2px; background-color: #0078d4; color: white; text-decoration: none; border-radius: 3px; font-size: 12px; }
        .button:hover { background-color: #005a9e; }
        .sort-indicator { margin-left: 5px; }
    </style>
</head>
<body>
    <h1>Mailbox Folder IDs for $Mailbox</h1>
    <table id="folderTable" style="white-space:nowrap; font-size 12px;">
        <thead>
            <tr>
                <th onclick="sortTable(0)">Name <span class="sort-indicator"></span></th>
                <th onclick="sortTable(1)">FolderType <span class="sort-indicator"></span></th>
                <th onclick="sortTable(2)">Identity <span class="sort-indicator"></span></th>
                <th onclick="sortTable(3)">FolderId <span class="sort-indicator"></span></th>
                <th onclick="sortTable(3)">eDiscoveryId <span class="sort-indicator"></span></th>
                <th onclick="sortTable(5)">EntryId <span class="sort-indicator"></span></th>
                <th onclick="sortTable(7)">RestId <span class="sort-indicator"></span></th>
                <th>RestId Link</th>
            </tr>
        </thead>
        <tbody>
"@

foreach ($folder in $output) {
    $htmlContent += @"
            <tr>
                <td>$($folder.Name)</td>
                <td>$($folder.FolderType)</td>
                <td>$($folder.Identity)</td>
                <td>$($folder.FolderId)</td>
                <td>$($folder.eDiscoveryId)</td>
                <td>$($folder.EntryId)</td>
                <td>$($folder.RestId)</td>
                <td><a href="https://developer.microsoft.com/graph/graph-explorer?request=admin%2FExchange%2FMailboxes%2F$mailboxID%2FFolders%2F$($folder.RestId)&method=GET&version=v1.0&GraphUrl=https://graph.microsoft.com" class="button" target="_blank">Open in Graph explorer</a></td>
            </tr>
"@
}

$htmlContent += @"
        </tbody>
    </table>
    <script>
        function sortTable(columnIndex) {
            const table = document.getElementById('folderTable');
            const tbody = table.querySelector('tbody');
            const rows = Array.from(tbody.querySelectorAll('tr'));
            let isAscending = true;

            const header = table.querySelectorAll('th')[columnIndex];
            if (header.classList.contains('sort-asc')) {
                isAscending = false;
                header.classList.remove('sort-asc');
                header.classList.add('sort-desc');
            } else {
                header.classList.remove('sort-desc');
                header.classList.add('sort-asc');
            }

            table.querySelectorAll('th').forEach((h, i) => {
                if (i !== columnIndex) {
                    h.classList.remove('sort-asc', 'sort-desc');
                }
            });

            rows.sort((a, b) => {
                const cellA = a.cells[columnIndex].textContent.trim();
                const cellB = b.cells[columnIndex].textContent.trim();

                const compareResult = cellA.localeCompare(cellB, undefined, { numeric: true });
                return isAscending ? compareResult : -compareResult;
            });

            rows.forEach(row => tbody.appendChild(row));
        }
    </script>
</body>
</html>
"@

$outputPath = "$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxFolderIDs.html"
$htmlContent | Out-File -FilePath $outputPath -Encoding UTF8
Write-Host "HTML report generated: $outputPath" -ForegroundColor Green
#Requires -Version 3.0
#Make sure to fill in all the required variables before running the script (lines 273-275)
#Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    User.ReadBasic.All (needed for the translateExchangeIds call)
#    Exchange.ManageAsApp and an Exchange administrator role assigned to the service principal (View-Only Recipients or equivalent)

[CmdletBinding()]
Param([Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][Alias("Identity")][String]$Mailbox,[switch]$IncludeNonIPM = $false)

#For details on what the script does and how to run it, check: https://michev.info/blog/post/8282/convert-mailbox-folder-identifiers-non-interactively

#region Helper functions
#Obtain an access token(s) or renew it if needed
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
        $authenticationResult = Invoke-RestMethod -Method Post -Uri $url -Body $body -ErrorAction Stop -Verbose:$false
        $token = $authenticationResult.access_token
    }
    catch { throw $_ }

    if (!$token) { Write-Error "Failed to aquire token!" -ErrorAction Stop; return }
    else {
        Write-Verbose "Successfully acquired Access Token for $service"

        #Use the access token to set the authentication header
        if (!$Service -or $Service -eq "Graph") { Set-Variable -Name authHeaderGraph -Scope Global -Value @{'Authorization'="Bearer $token";'Content-Type'='application/json'} -Confirm:$false}
        elseif ($Service -eq "Exchange") { Set-Variable -Name authHeaderExchange -Scope Global -Value @{'Authorization'="Bearer $token";'Content-Type'='application/json'} -Confirm:$false }
        else { Write-Error "Invalid service specified, aborting..." -ErrorAction Stop; return }
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

        Renew-Token -Service "Graph"
        Renew-Token -Service "Exchange"

        if (!$authHeaderGraph) { Write-Error "Failed to renew token, aborting..." -ErrorAction Stop }
        if (!$authHeaderExchange) { Write-Error "Failed to renew token, aborting..." -ErrorAction Stop }
    }
    #The rest are non-terminal errors
    elseif ($ErrorMessage.ErrorDetails.Message -match "ManagementObjectNotFoundException|ADNoSuchObjectException|Couldn't find object|Unable to retrieve mailbox folder statistics for mailbox") { Write-Error "The specified object was not found, check the input values..." -ErrorAction Stop }
    elseif ($ErrorMessage.ErrorDetails.Message -match "Invalid object identifier|The requested user .* is invalid|does not exist or one of its queried reference-property|Unsupported referenced-object resource identifier") { Write-Error "The specified object was not found, this should not happen..." -ErrorAction Stop }
    else { $ErrorMessage | fl * -Force; return } #catch-all for any unhandled errors
}

function ReturnFolderList {
<#
.Synopsis
    Enumerates all user-accessible folders for the mailbox.
.DESCRIPTION
    The ReturnFolderList cmdlet enumerates the folders for the given mailbox.
.PARAMETER SMTPAddress
	Use the -SMTPAddress parameter to designate the mailbox where the desired folders reside.
.PARAMETER IncludeNonIPM
    Use the -IncludeNonIPM switch to include folders from the non-IPM tree.
.EXAMPLE
    ReturnFolderList user@domain.com
    This command will return a list of all user-accessible folders for the user@domain.com mailbox.
.INPUTS
    Identifier for the mailbox, UPN or GUID.
.OUTPUTS
    Array with information about the mailbox folders.
#>

    param(
	[Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$SMTPAddress, #Best use UPN
    [switch]$IncludeNonIPM) #whether to include folders from the non-IPM tree

    if ($IncludeNonIPM) { $folderScope = "NonIpmRoot" }
    else { $folderScope = "All" }

    #Prepare the payload for the REST API call to get mailbox folders
    $body = @{
        CmdletInput = @{
            CmdletName="Get-MailboxFolderStatistics"
            Parameters=@{"Identity"=$SMTPAddress;"FolderScope"=$FolderScope;"ResultSize"="Unlimited"}
        }
    }

    $uri = "https://outlook.office365.com/adminapi/beta/$($TenantID)/InvokeCommand"
    try {
        $result = Invoke-RestMethod -Method POST -Uri $uri -Headers $authHeaderExchange -Body ($body | ConvertTo-Json -Depth 5) -ContentType "application/json" -Verbose:$false -ErrorAction Stop #suppress the output
    }
    catch {
        Process-Error -ErrorMessage $_
    }

    $MBfolders = $result.value

    if (!$MBfolders -or $MBfolders.count -eq 0) { Write-Error "No folders found for the specified mailbox. The script will exit." -ErrorAction Stop; return }

    return ($MBfolders | select Name,FolderType,Identity,FolderId,@{n="eDiscoveryId";e={FolderIdEDiscovery $_.FolderId}},@{n="EntryId";e={FolderIdToEntryId $_.FolderId}})
}

function FolderIdToEntryId {
<#
.Synopsis
    Transforms folderId value to entryId format.
.DESCRIPTION
    The FolderIdToEntryId cmdlet transforms the folderId value obtained from Get-ExOMailboxFolderStatistics to the entryId format used by MAPI clients.
.PARAMETER -FolderId
	Use the -FolderId parameter to designate the original folderId value.
.EXAMPLE
    FolderIdToEntryId LgAAAAChKSJAhlnUTIHtKSso30ThAQBIPfDMxyP/RYhY8M8xmAPVAAAU1V6iAAAD
    This command will convert the given folderId value to the entryId format.
.INPUTS
    FolderId value obtained from Get-ExOMailboxFolderStatistics.
.OUTPUTS
    The converted entryId value in base64 format.
.LINK
    https://stackoverflow.com/a/75482631
#>

    param([Parameter(Mandatory=$true)]$folderId)

    # convert from base64 to bytes
    $folderIdBytes = [Convert]::FromBase64String($folderId)

    # convert byte array to string, remove '-' and ignore first byte
    $folderIdHexString = [System.BitConverter]::ToString($folderIdBytes).Replace('-','')
    $folderIdHexStringLength = $folderIdHexString.Length

    # get hex entry id string by removing first and last byte
    $entryIdHexString = $folderIdHexString.SubString(2,($folderIdHexStringLength-4))

    # convert to byte array - two chars represents one byte
    $entryIdBytes = [byte[]]::new($entryIdHexString.Length / 2)

    For($i=0; $i -lt $entryIdHexString.Length; $i+=2){
        $entryIdTwoChars = $entryIdHexString.Substring($i, 2)
        $entryIdBytes[$i/2] = [convert]::ToByte($entryIdTwoChars, 16)
    }

    # convert bytes to base64 string
    $entryIdBase64 = [Convert]::ToBase64String($entryIdBytes)

    # count how many '=' contains base64 entry id
    $equalCharCount = $entryIdBase64.Length - $entryIdBase64.Replace('=','').Length

    # trim '=', replace '/' with '-', replace '+' with '_' and add number of '=' at the end
    $entryId = $entryIdBase64.TrimEnd('=').Replace('/','_').Replace('+','-')+$equalCharCount

    return $entryId
}

function FolderIdEDiscovery {
<#
.Synopsis
    Transforms folderId value to the format accepted by eDiscovery searches.
.DESCRIPTION
    The FolderIdEDiscovery cmdlet transforms the folderId value obtained from Get-ExOMailboxFolderStatistics to the format used by eDiscovery targeted collections feature.
.PARAMETER FolderId
	Use the -FolderId parameter to designate the original folderId value.
.EXAMPLE
    FolderIdEDiscovery LgAAAAChKSJAhlnUTIHtKSso30ThAQBIPfDMxyP/RYhY8M8xmAPVAAAU1V6iAAAD
    This command will convert the given folderId value to the eDiscovery format.
.INPUTS
    FolderId value obtained from Get-ExOMailboxFolderStatistics.
.OUTPUTS
    The converted FolderId value in the format accepted by eDiscovery searches.
.LINK
    https://www.enowsoftware.com/solutions-engine/m365-sharepoint-onedrive-center/performing-ediscovery-against-a-specific-folder
#>
    param([Parameter(Mandatory=$true)]$folderId)

    # convert from base64 to bytes
    $folderId = [Convert]::FromBase64String($folderId)

    #
    $encoding = [System.Text.Encoding]::GetEncoding("us-ascii")
    $nibbler = $encoding.GetBytes("0123456789ABCDEF")

    # the value is stored in the middle of the folderId
    $indexIdBytes = New-Object byte[] 48; $indexIdIdx = 0
    $folderId | select -skip 23 -first 24 | % { $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -shr 4]; $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -band 0x0F] }

    return $encoding.GetString($indexIdBytes)
}

function EntryIdToRestId {
<#
.Synopsis
    Transforms entryId value to RestId format.
.DESCRIPTION
    The EntryIdToRestId cmdlet transforms the entryId value obtained from FolderIdToEntryId to the RestId format used by the Graph API.
.PARAMETER Mailbox
    Use the -Mailbox parameter to designate the mailbox where the folders reside. Mandatory. The calling user must have User.ReadBasic.All permissions for the call to succeed.
.PARAMETER Ids
    Use the -Ids parameter to provide the entryId values to be converted. Multiple values can be provided as an array, but the total number of IDs in a single call must not exceed 1000.
.EXAMPLE
    EntryIdToRestId -Mailbox user@domain.com -Ids $entryId1, $entryId2, $entryId3
    This command will convert the given folderId value(s) to the RestId format.
.INPUTS
    entryId value obtained from FolderIdToEntryId and the mailbox SMTP address.
.OUTPUTS
    The converted restId value.
#>
    param([Parameter(Mandatory=$true)][string[]]$Ids, #max 1000
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$Mailbox) #don't cast as SMTP, as the Graph expects GUID/UPN

    #Hash table to store the translated IDs
    $RestIDs = @{}

    do {
        $batch = $Ids | Select-Object -First 1000
        $Ids = $Ids | Select-Object -Skip 1000

        #Prepare the request body for the batch of IDs
        $params = @{
            inputIds = @($batch)
            sourceIdType = "entryId"
            targetIdType = "restId"
        }

        # Execute the request
        try {
           $res = Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$Mailbox/translateExchangeIds" -Headers $authHeaderGraph -Body ($params | ConvertTo-Json -Depth 5) -ContentType "application/json" -Verbose:$false -ErrorAction Stop
        }
        catch {
            Process-Error -ErrorMessage $_
        }

        foreach ($item in $res.value) {
            if ($item.TargetId -and !$item.ErrorDetails.Code) { $RestIDs[$item.SourceId] = $item.TargetId }
            else { Write-Warning "Failed to translate ID: $($item.SourceId). Error details: $($item.ErrorDetails.Code)"; $RestIDs[$item.SourceId] = "N/A"; continue }
        }
    } while ($Ids.Count -gt 0)

    return $restIDs
}
#endregion

#==========================================================================
# Main script
#==========================================================================

#Variables to configure
$tenantID = "tenant.onmicrosoft.com" #your tenantID or tenant root domain
$appID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #the GUID of your app
$client_secret = "verylongsecurestring" #client secret for the app

Renew-Token -Service "Graph"
Renew-Token -Service "Exchange"

#Check if the input value is GUID and if not, try to resolve it to a valid mailbox identity. SMTP address will not work for /users either.
if (![Guid]::TryParse($Mailbox,[ref]([System.Guid]::empty))) {

    #We use Get-Mailbox here, as the user must have a mailbox anyway for the script to work
    $body = @{
        CmdletInput = @{
            CmdletName="Get-Mailbox"
            Parameters=@{"Identity"=$Mailbox}
        }
    }

    $uri = "https://outlook.office365.com/adminapi/beta/$($TenantID)/InvokeCommand"
    try {
        $result = Invoke-RestMethod -Method POST -Uri $uri -Headers $authHeaderExchange -Body ($body | ConvertTo-Json -Depth 5) -ContentType "application/json" -Verbose:$false -ErrorAction Stop #suppress the output
    }
    catch {
        Process-Error -ErrorMessage $_
    }

    if (!$result -or !$result.value -or $result.value.count -ne 1) { Write-Error "Failed to resolve the provided mailbox identity. The script will exit." -ErrorAction Stop; return }
    $MailboxID = $result.value[0].ExternalDirectoryObjectId
}
else { $MailboxID = $Mailbox }

#Get the folder list, output contains FolderId, eDiscoveryId and EntryId
$temp = ReturnFolderList $Mailbox -IncludeNonIPM:$IncludeNonIPM

#Convert EntryId to RestId using the Graph API, store in a hash table for easy retrieval
$RestIDs = EntryIdToRestId -Ids $temp.EntryId -Mailbox $MailboxID
#Prepare the final output by adding the RestId values from the hash table, and also an OWA-friendly version of the RestId for direct linking to OWA
$output = $temp | select Name,FolderType,Identity,FolderId,eDiscoveryId,EntryId,@{n="RestId";e={$RestIDs[$_.EntryId]}}

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
                <td><a href="https://developer.microsoft.com/graph/graph-explorer?request=users%2F$mailboxID%2FmailFolders%2F$($folder.RestId)&method=GET&version=v1.0&GraphUrl=https://graph.microsoft.com" class="button" target="_blank">Open in Graph explorer</a></td>
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
#Requires -Version 3.0
#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.0.0" }
#Requires -Modules @{ ModuleName="Microsoft.Graph.Users.Actions"; ModuleVersion="2.37.0" }

[CmdletBinding()]
Param([Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][Alias("Identity")][String]$Mailbox,[switch]$IncludeNonIPM = $false)

#For details on what the script does and how to run it, check: https://michev.info/blog/post/7984/converting-get-mailboxfolderstatistics-ids-for-use-with-the-graph-api

#region Helper functions
function Check-Connectivity {
    [cmdletbinding()]
    [OutputType([bool])]
    param()

    #Make sure we are connected to Exchange Remote PowerShell
    Write-Verbose "Checking connectivity..."

    #Check via Get-ConnectionInformation first
    if (Get-ConnectionInformation) { return $true }

    #Double-check and try to eastablish a session if needed
    try { Get-EXOMailbox -ResultSize 1 -ErrorAction Stop -Verbose:$false | Out-Null }
    catch {
        try { Connect-ExchangeOnline -CommandName Get-ExOMailboxFolderStatistics, Get-MailboxFolderStatistics -SkipLoadingFormatData -ShowBanner:$false -Verbose:$false } #custom for this script
        catch { Write-Error "No active Exchange Online session detected. To connect to ExO: https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps"; return $false }
    }

    # Make sure we are connected to Microsoft Graph and have the required permissions
    Write-Verbose "Checking connectivity to Graph PowerShell..."
    try {
        if (!(Get-MgContext) -or !((Get-MgContext).Scopes.Contains("User.ReadBasic.All"))) {
            Write-Verbose "Not connected to the Microsoft Graph or the required permissions are missing!"
            Connect-MgGraph -Scopes User.ReadBasic.All -ErrorAction Stop | Out-Null
        }
    }
    catch { Write-Error $_; return $false }
    #Double-check required permissions
    if (!((Get-MgContext).Scopes.Contains("Group.ReadWrite.All"))) { Write-Error "The required permissions are missing, please re-consent!"; return $false }

    return $true
}

function ReturnFolderList {
<#
.Synopsis
    Enumerates all user-accessible folders for the mailbox.
.DESCRIPTION
    The ReturnFolderList cmdlet enumerates the folders for the given mailbox.
.PARAMETER SMTPAddress
	Use the -SMTPAddress parameter to designate the mailbox where the desired folders reside.
.EXAMPLE
    ReturnFolderList user@domain.com
    This command will return a list of all user-accessible folders for the user@domain.com mailbox.
.INPUTS
    SMTP address of the mailbox, with optional parent folder (full path).
.OUTPUTS
    Array with information about the mailbox folders.
#>

    param(
	[Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()]$SMTPAddress, #Best use UPN
    [switch]$IncludeNonIPM) #whether to include folders from the non-IPM tree

    #Make sure we are connected to Exchange Remote PowerShell
    if (Check-Connectivity) { Write-Verbose "Connected to Exchange Remote PowerShell, processing..." }
    else { Write-Host "ERROR: Connectivity test failed, exiting the script..." -ForegroundColor Red; continue }

    if ($IncludeNonIPM) { $folderScope = "NonIpmRoot" }
    else { $folderScope = "All" }
    $MBfolders = Get-ExOMailboxFolderStatistics -FolderScope $folderScope -Identity $SMTPAddress -Verbose:$false | Select-Object Name,FolderType,FolderPath,Identity,FolderId

    if (!$MBfolders) { return }

    return ($MBfolders | select Name,FolderType,Identity,FolderId,@{n="eDiscoveryId";e={FolderIdEDiscovery $_.FolderId}},@{n="EntryId";e={FolderIdToEntryId $_.FolderId}})
}

function FolderIdToEntryId {
<#
.Synopsis
    Transforms folderId value to entryId format.
.DESCRIPTION
    The FolderIdToEntryId cmdlet transforms the folderId value obtained from Get-ExOMailboxFolderStatistics to the entryId format used by MAPI clients.
.PARAMETER SMTPAddress
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
.PARAMETER SMTPAddress
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

#Basic function to "fix" the RestId for use in OWA links
function Base64URLSafe {
    param([Parameter(Mandatory=$true)][string]$Base64String)
    return $Base64String.Replace('-','%2F').Replace('=','%3D').Replace('_','%2B')
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
    param([Parameter(Mandatory=$true)][string[]]$Ids, #max 1000, add a check?
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$Mailbox) #don't cast as SMTP, as the Graph expects GUID/UPN

    #Make sure we are connected to Exchange Remote PowerShell
    if (Check-Connectivity) { Write-Verbose "Connected to the Graph, processing..." }
    else { Write-Host "ERROR: Connectivity test failed, exiting the script..." -ForegroundColor Red; continue }

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
         $res = Invoke-MgTranslateUserExchangeId -UserId $Mailbox -BodyParameter $params -ErrorAction Stop
         foreach ($item in $res) {
             if ($item.TargetId -and !$item.ErrorDetails.Code) { $RestIDs[$item.SourceId] = $item.TargetId }
             else { Write-Warning "Failed to translate ID: $($item.SourceId). Error details: $($item.ErrorDetails.Code)"; $RestIDs[$item.SourceId] = "N/A"; continue }
         }
    } while ($Ids.Count -gt 0)

    return $restIDs
}
#endregion

#Main script starts here

#Get the folder list, output contains FolderId, eDiscoveryId and EntryId
$temp = ReturnFolderList $Mailbox -IncludeNonIPM:$IncludeNonIPM
#Convert EntryId to RestId using the Graph API, store in a hash table for easy retrieval
$RestIDs = EntryIdToRestId -Ids $temp.EntryId -Mailbox $Mailbox
#Prepare the final output by adding the RestId values from the hash table, and also an OWA-friendly version of the RestId for direct linking to OWA
$output = $temp | select Name,FolderType,Identity,FolderId,eDiscoveryId,EntryId,@{n="RestId";e={$RestIDs[$_.EntryId]}},@{n="OWAId";e={Base64URLSafe $RestIDs[$_.EntryId]}}

#Export the output to CSV
$output | Select * | Export-Csv -Path "$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_MailboxFolderIDs.csv" -NoTypeInformation -Encoding UTF8

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
                <th>OWAId Link</th>
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
                <td><a href="https://developer.microsoft.com/graph/graph-explorer?request=users%2F$mailbox%2FmailFolders%2F$($folder.RestId)&method=GET&version=v1.0&GraphUrl=https://graph.microsoft.com" class="button" target="_blank">Open in Graph explorer</a></td>
                <td><a href="https://outlook.cloud.microsoft/mail/$mailbox/$($folder.OWAId)" class="button" target="_blank">Open in OWA</a></td>
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
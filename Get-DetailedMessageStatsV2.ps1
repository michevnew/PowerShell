[CmdletBinding()] #Make sure we can use -Verbose
Param()

#Confirm connectivity to Exchange Online.
Write-Verbose "Connecting to Exchange Online..."
try { Get-EXORecipient -ResultSize 1 -ErrorAction Stop -Verbose:$false | Out-Null }
catch {
    try { Connect-ExchangeOnline -CommandName Get-MessageTraceV2 -SkipLoadingFormatData -ShowBanner:$false -Verbose:$false } #needs to be non-REST cmdlet
    catch { Write-Error "No active Exchange Online session detected. To connect to ExO: https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps"; return }
}

Write-Verbose "Collecting Recipients..."

#Collect all Exchange Online recipient aliases
$Recipients = Get-ExORecipient -ResultSize Unlimited -Verbose:$false | Select PrimarySMTPAddress,RecipientTypeDetails,EmailAddresses
if ($Recipients.Count -eq 0) {
    Write-Error "No recipients found in the tenant. Please check your connectivity and permissions."
    return
}

$MailTraffic = @{}
foreach($Recipient in $Recipients)
{
    $MailTraffic[$Recipient.PrimarySMTPAddress.ToLower()] = @{}
    $MailTraffic[$Recipient.PrimarySMTPAddress.ToLower()]["Aliases"] = @($Recipient.EmailAddresses | ? {$_ -match "smtp:"} | % { $_.Split(":")[1]})
}
Remove-Variable Recipients

#Collect Message Trace data
$StartDate = (Get-Date).AddDays(-10) #max period we can cover in a single query is 10 days, if needed rerun multiple times to cover up to 90
$EndDate = (Get-Date)

#Get the first "page"
$Messages = $null
$cMessages = Get-MessageTraceV2 -ResultSize 5000 -StartDate $StartDate -EndDate $EndDate -WarningVariable MoreResultsAvailable -Verbose:$false 3>$null
$Messages += $cMessages | Select Received,SenderAddress,RecipientAddress,Size,Status

#If more results are available, as indicated by the presence of the WarningVariable, we need to loop until we get all results
if ($MoreResultsAvailable) {
    do {
        #As we don't have a clue how many pages we will get, proper progress indicator is not feasible.
        Write-Host "." -NoNewline

        #Handling this via Warning output is beyong annoying...
        $NextPage = ($MoreResultsAvailable -join "").TrimStart("There are more results, use the following command to get more. ")
        $ScriptBlock = [ScriptBlock]::Create($NextPage)
        $cMessages = Invoke-Command -ScriptBlock $ScriptBlock -WarningVariable MoreResultsAvailable -Verbose:$false 3>$null #MUST PASS WarningVariable HERE OR IT WILL NOT WORK
        $Messages += $cMessages | Select Received,SenderAddress,RecipientAddress,Size,Status
    }
    until ($MoreResultsAvailable.Count -eq 0) #Arraylist
}
#If no messages were found, exit
if ($Messages.Count -eq 0) {
    Write-Error "No messages found for the specified date range. Please check your permissions or update the date range above."
    return
}

Write-Verbose "Crunching Results..."

#Read each message trace entry and add it to a hash table
foreach($Message in $Messages) {
    #Skip messages sent to plus addresses, we have duplicate entries for those. Or exclude "Resolved" status?
    if ($Message.SenderAddress.Contains("+") -or $Message.RecipientAddress.Contains("+")) { continue }

    #Process the semder address
    if ($null -ne $Message.SenderAddress) {
        # Normalize to lower case and make sure to account for aliases
        $Address = $Message.SenderAddress.ToLower()
        $Key = $MailTraffic.GetEnumerator() | ? {$_.Value.Values -match $Address } | select -ExpandProperty Name
        if ($Key -and ($Address -ne $Key)) {
            $Address = $Key
        }

        #If a valid recipient, add it to the output
        if ($MailTraffic.ContainsKey($Address)) {
            $MessageDate = Get-Date -Date $Message.Received -Format yyyy-MM-dd

            if ($MailTraffic[$Address].ContainsKey($MessageDate)) {
                $MailTraffic[$Address][$MessageDate]['Outbound']++
                $MailTraffic[$Address][$MessageDate]['OutboundSize'] += $Message.Size
            }
            else {
                $MailTraffic[$Address][$MessageDate] = @{}
                $MailTraffic[$Address][$MessageDate]['Outbound'] = 1
                $MailTraffic[$Address][$MessageDate]['Inbound'] = 0
				$MailTraffic[$Address][$MessageDate]['InboundSize'] = 0
				$MailTraffic[$Address][$MessageDate]['OutboundSize'] += $Message.Size
            }
        }
    }

    #Process the recipient address
    if ($null -ne $Message.RecipientAddress) {
        # Normalize to lower case and make sure to account for aliases
        $Address = $Message.RecipientAddress.ToLower()
        $Key = $MailTraffic.GetEnumerator() | ? {$_.Value.Values -match $Address } | select -ExpandProperty Name
        if ($Key -and ($Address -ne $Key)) {
            $Address = $Key
        }

        #If a valid recipient, add it to the output
        if ($MailTraffic.ContainsKey($Address)) {
            $MessageDate = Get-Date -Date $Message.Received -Format yyyy-MM-dd

            if ($MailTraffic[$Address].ContainsKey($MessageDate)) {
                $MailTraffic[$Address][$MessageDate]['Inbound']++
				$MailTraffic[$Address][$MessageDate]['InboundSize'] += $Message.Size
            }
            else {
                $MailTraffic[$Address][$MessageDate] = @{}
                $MailTraffic[$Address][$MessageDate]['Inbound'] = 1
                $MailTraffic[$Address][$MessageDate]['Outbound'] = 0
				$MailTraffic[$Address][$MessageDate]['OutboundSize'] = 0
				$MailTraffic[$Address][$MessageDate]['InboundSize'] += $Message.Size
			}
        }
    }
}

Write-Verbose "Formatting Results..."

$table = @()
#Transpose hashtable to PSObject
ForEach ($RecipientName in $MailTraffic.keys) {

    foreach($Date in ($MailTraffic[$RecipientName].keys | ? {$_ -ne "Aliases"})) {
        $row = [ordered]@{
            Date = $Date
            Recipient = [string]$RecipientName
            Inbound = [int32]$MailTraffic[$RecipientName][$Date].Inbound
            Outbound = [int32]$MailTraffic[$RecipientName][$Date].Outbound
            InboundSize = [int32]$MailTraffic[$RecipientName][$Date].InboundSize
            OutboundSize = [int32]$MailTraffic[$RecipientName][$Date].OutboundSize
        }
        $table += [PSCustomObject]$row
    }
}

#Export to CSV
$table | Sort-Object -Property Date | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_DetailedMessageStats.csv" -NoTypeInformation -Encoding UTF8 -UseCulture

# Generate sortable HTML table with type-aware sorting
$HtmlHeader = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Detailed Message Statistics</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; background: #f4f6f8; color: #222; }
h1 { background: #0078d4; color: #fff; padding: 16px; border-radius: 6px 6px 0 0; margin-bottom: 20px; }
table { width: 100%; background: #fff; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border-collapse: collapse; }
th, td { padding: 12px; text-align: left; }
th { background: #e5eaf1; cursor: pointer; position: relative; }
th:hover { background: #d0e7fa; }
th::after { content: '↕'; position: absolute; right: 8px; opacity: 0.5; }
tr:nth-child(even) { background: #f0f4fa; }
tr:hover { background: #d0e7fa; }
</style>
<script>
function parseValue(val, type) {
    if(type === 'number') return parseFloat(val.replace(/,/g,'')) || 0;
    if(type === 'date') return new Date(val);
    return val.toLowerCase();
}
function sortTable(n, type) {
    var table = document.getElementById('msgstats');
    var rows = Array.from(table.rows).slice(1);
    var dir = table.getAttribute('data-sortdir'+n) === 'asc' ? 'desc' : 'asc';
    rows.sort(function(a, b) {
        var x = parseValue(a.cells[n].innerText, type);
        var y = parseValue(b.cells[n].innerText, type);
        if(x < y) return dir === 'asc' ? -1 : 1;
        if(x > y) return dir === 'asc' ? 1 : -1;
        return 0;
    });
    rows.forEach(function(row) { table.tBodies[0].appendChild(row); });
    table.setAttribute('data-sortdir'+n, dir);
}
</script>
</head>
<body>
<h1>Detailed Message Statistics</h1>
<table id="msgstats">
<thead>
<tr>
<th onclick="sortTable(0,'date')">Date</th>
<th onclick="sortTable(1,'string')">Recipient</th>
<th onclick="sortTable(2,'number')">Inbound</th>
<th onclick="sortTable(3,'number')">Outbound</th>
<th onclick="sortTable(4,'number')">InboundSize</th>
<th onclick="sortTable(5,'number')">OutboundSize</th>
</tr>
</thead>
<tbody>
"@

$HtmlRows = foreach ($row in $table | Sort-Object Date,Recipient) {
    "<tr><td>$($row.Date)</td><td>$($row.Recipient)</td><td>$($row.Inbound)</td><td>$($row.Outbound)</td><td>$($row.InboundSize)</td><td>$($row.OutboundSize)</td></tr>"
}

$HtmlFooter = @"
</tbody>
</table>
</body>
</html>
"@

#Generate the full HTML content and save it to a file
$FullHtml = $HtmlHeader + ($HtmlRows -join "`n") + $HtmlFooter
$FullHtml | Set-Content -Encoding UTF8 "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_DetailedMessageStats.html"
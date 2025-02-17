[CmdletBinding()] #Make sure we can use common parameters
Param(
    [switch]$IncludeAll, #Include all users with OneDrives, not just the licensed ones
    [int][ValidateRange(1,3650)]$Days = 30 #Number of days to check for activity
)

# Connecting to Microsoft Graph & Exchange Online
Connect-MgGraph -Scopes "Directory.Read.All" -ErrorAction Stop -NoWelcome #LicenseAssignment.Read.All
if (!(Get-ConnectionInformation)) { Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop -SkipLoadingFormatData -SkipLoadingCmdletHelp -CommandName Search-UnifiedAuditLog }

Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -ErrorAction Stop
Connect-SPOService -Url "https://tenant.sharepoint.com" #variable

#Retrieve all users with OneDrive
$OneDriveUsers = Get-SPOSite -IncludePersonalSite $true -Limit All -Template "SPSPERS" | select LastContentModifiedDate,Status, ArchiveStatus, StorageUsageCurrent, Url, Owner, SharingCapability, OverrideSharingCapability
#The LIST method occasionally returns empty Owner values... repeat with GET
$OneDriveUsers | ? {!$_.Owner} | % { $_.Owner = (Get-SPOSite -Identity $_.Url).Owner }
if (!$OneDriveUsers) { Write-Warning "No OneDrive sites found..."; return }

#Generate the report file
if ($null -eq (Get-Module -Name ImportExcel -ListAvailable -Verbose:$false)) {
    Write-Verbose "The ImportExcel module was not found, skipping export to Excel file..."; return
}
#ImportExcel does NOT overwrite existing files, so we need to generate a unique filename
$excel = Export-Excel -Path "$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_InactiveOneDrives.xlsx" -WorksheetName Overview -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -PassThru

$endDate = Get-Date
$startDate = $endDate.AddDays(-$Days)
$OneDrivePlans = @("b4ac11a0-32ff-4e78-982d-e039fa803dec","f7e5b77d-f293-410a-bae8-f941f19fe680","13696edf-5a08-49f6-8134-03083ed8ba30","4495894f-534f-41ca-9d3b-0ebf1220a423","afcafa6a-d966-4462-918c-ec0b4e0fe642","da792a53-cbc0-4184-a10d-e544dd34b3c1","98709c2e-96b5-4244-95f5-a0ebe139fb8a")
$SharePointPlans = @("e95bec33-7c88-4a70-8e19-b10bd9d0c014","5dbe027f-2339-4123-9542-606e4d348a72","902b47e5-dcb2-4fdc-858b-c63a90a2bdb9","63038b2c-28d0-45f6-bc36-33062963b498","6b5b6a67-fc72-4a1f-a2b5-beecf05de761","c7699d2e-19aa-44de-8edf-1736da088ca1","0a4983bb-d3e5-4a09-95d8-b2d0127b3df5")

#Process entries for each object with role assignments
$UsersReport = @();
foreach ($user in $OneDriveUsers) {
    #Check the status of the drive
    if ($user.Status -ne "Active") { continue }

    #Get the user details
    $userGraph = Get-MgUser -UserId $user.Owner -Property Id,assignedPlans -ErrorAction SilentlyContinue | select Id, assignedPlans

    #Apparently SPO can return non-existent userIds, skip them
    if ($userGraph) {
        #Check if the user has a OneDrive license
        $userHasOneDriveLicense = (($userGraph.AssignedPlans | ? { ($_.ServicePlanId -in $OneDrivePlans) -or ($_.ServicePlanId -in $SharePointPlans) } | ? {$_.CapabilityStatus -eq "Enabled"})) ? $true : $false
    }
    else { $userHasOneDriveLicense = $false }

    if ($userHasOneDriveLicense) {
        #If the user has a license and we're not using the -IncludeAll switch, add a "basic" entry and continue
        if (!$IncludeAll) {
            $UsersReport += [PSCustomObject]@{
                SiteURL = $user.Url
                OwnerId = $userGraph.Id
                OwnerUpn = $user.Owner
                HasOneDriveLicense = "Yes"
                LastAction = "Not checked"
                ActionCount = "Not checked"
                "UsedStorage (MB)" = $user.StorageUsageCurrent
                AccessedByOtherUsers = "Not checked"
                AccessedByGuests = "Not checked"
            }
            continue
        }
    }

    $userUpn = $user.Owner
    #$userIds = ($userGraph.Id) ? @($userUpn, $userGraph.Id) : @($userUpn)
    $userURL = $user.Url + "/*"

    #Get UnifiedAuditLog events for each admin
    $userAuditLogsUALTemp = @()
    $sessionID = (New-Guid).Guid + "$userUpn"
    Write-Verbose "Collecting UAL entries for $userUpn..."
    do {
        #Don't pass UserIds, as we want to get all events for the site
        $userAuditLogsUAL = Search-UnifiedAuditLog -StartDate $startDate.Date -EndDate $endDate.AddDays(1).Date -SessionCommand ReturnLargeSet -ResultSize 5000 -SessionId $sessionID -RecordType SharePointFileOperation -ObjectIds $userUrl #`
        #`-Operations FileUploaded, FileAccessed, FileDeleted, FilePreviewed, FileModified, FileRenamed, FileModifiedExtended, FileCheckedIn, FileRecycled
        #`-Operations FileAccessed,FileDeleted,FolderModified,FileModifiedExtended,FileModified,FileSyncUploadedFull,FileSensitivityLabelApplied,FilePreviewed,FileSyncDownloadedFull,FolderRecycled,FolderCreated,FileDownloaded,FileMoved,FileCopied,FileVersionsAllDeleted,FileSensitivityLabelRemoved,FileSensitivityLabelChanged,FileUploaded,FileRecycled,FileRenamed,FileMalwareDetected,FileAccessedExtended

        #Trim duplicated records, filter out some noise
        $userAuditLogsUALTemp += $userAuditLogsUAL | Select-Object -ExpandProperty AuditData -Unique | ConvertFrom-Json | Where-Object { $_.UserId -notin @("SHAREPOINT\system","app@sharepoint","Microsoft\ServiceOperator") }

        if (!$userAuditLogsUAL -or ($userAuditLogsUAL[-1].ResultIndex -ge $userAuditLogsUAL[-1].ResultCount)) { break }
        Write-Host "." -NoNewline
    } while ($userAuditLogsUAL) #50k per admin max

    $UsersReport += [PSCustomObject]@{
        SiteURL = $user.Url
        OwnerId = ($userGraph.Id) ? $($userGraph.Id) : "Unknown"
        OwnerUpn = $user.Owner
        HasOneDriveLicense = ($userHasOneDriveLicense) ? "Yes" : "No"
        LastAction = ($userAuditLogsUALTemp) ? ($userAuditLogsUALTemp.CreationTime | Sort-Object -Descending | Select-Object -First 1) : "N/A" # UAL entries are NOT sorted
        ActionCount = ($userAuditLogsUALTemp) ? $userAuditLogsUALTemp.Count : 0
        "UsedStorage (MB)" = $user.StorageUsageCurrent
        AccessedByOtherUsers = ($userAuditLogsUALTemp.UserId | Sort-Object -Unique | Where-Object { $_ -ne $user.Owner }) ? "Yes" : "No"
        AccessedByGuests = ($userAuditLogsUALTemp.UserId | Sort-Object -Unique | Where-Object { $_ -match "#EXT#" }) ? "Yes" : "No"
    }

    if ($userAuditLogsUALTemp) {
        if ($userUpn.length -gt 31) { $userUpn = $userUpn.Substring(0,31) }

        if ($userAuditLogsUALTemp.Count -gt 0) { #No need to add empty sheets
            $userAuditLogsUALTemp | select CreationTime, Operation, UserId, ObjectId, ClientIP, UserAgent `
            ` | Export-Excel -ExcelPackage $excel -WorksheetName $userUpn -FreezeTopRow -AutoFilter -BoldTopRow -AutoSize -NoHyperLinkConversion TargetResource -PassThru > $null
        }
    }
}

#Add the remaining sheets to the XLSX file
$UsersReport | Export-Excel -ExcelPackage $excel -WorksheetName Overview -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -MoveToStart -PassThru > $null

#Make UPNs in the Overview sheet clickable
$sheet = $excel.Workbook.Worksheets["Overview"]

#Add a hyperlink to the MemberUpn/ActionCount columns
$cells = $sheet.Cells["B2:B"] #Gives just the populated cells
foreach ($cell in $cells) {
    #Process only rows corresponding to user objects
    $cellValue = $cell.Value
    if ($cell.Value.length -gt 31) { $cellValue = $cell.Value.Substring(0,31) }
    $otherCell = $sheet.Cells[$cell.Address.Replace("B","F")]
    if (!($otherCell.Value) -or ($otherCell.Value -eq "0")) { continue }

    if ($excel.Workbook.Worksheets[$cellValue]) {
        $targetAddress = $excel.Workbook.Worksheets[$cellValue].Cells["A1"].FullAddress
        $cell.Hyperlink = New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList $targetAddress, $cell.Value
        $otherCell.Hyperlink = New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList $targetAddress, $otherCell.Value
        $otherCell.Value = [int]$otherCell.Value #OMFG spent two hours on this

        $cell.Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
        $cell.Style.Font.Underline = $true
        $otherCell.Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
        $otherCell.Style.Font.Underline = $true
    }
}

Export-Excel -ExcelPackage $excel -WorksheetName "Overview" -Show > $null
$excel.Dispose()
Write-Verbose "Finished..."
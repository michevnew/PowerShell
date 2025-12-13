#Requires -Version 7.4
# Make sure to fill in all the required variables before running the script
# Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Sites.Read.All

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/6276/reporting-on-sharepoint-online-and-onedrive-for-business-item-size-with-version-history-included-using-the-graph-api

[CmdletBinding()] #Make sure we can use -Verbose
Param(
[string[]][ValidateNotNullOrEmpty()]$Sites, #Use the Sites parameter to specify a set of sites to process.
[switch]$IncludeODFBsites, #Use the IncludeODFBsites switch to specify whether to include personal OneDrive for Business sites in the output.
[switch]$NoItemLevelStats, #Use the NoItemLevelStats switch to include item-level statistics in the output.
[switch]$IncludeVersions, #Use the IncludeVersions switch to also include item versions statistics in the output.
[switch]$ExportToExcel #Use the ExportToExcel switch to specify whether to export the output to an Excel file.
)

function processChildren {
    Param(
    #Graph Site object
    [Parameter(Mandatory=$true)]$Site,
    #URI for the drive
    [Parameter(Mandatory=$true)][string]$URI)

    if ($tokenExp -lt [datetime]::Now.AddSeconds(360)) {
        Write-Verbose "Access token is about to expire, renewing..."
        Renew-Token
    }

    $children = @()
    #fetch children, make sure to handle multiple pages
    do {
        $result = Invoke-GraphApiRequest -Uri "$URI" -Verbose:$VerbosePreference
        $URI = $result.'@odata.nextLink'

        #If we are getting multiple pages, add some delay to avoid throttling
        Start-Sleep -Milliseconds 300
        $children += $result
    } while ($URI)
    if (!$children.value) { Write-Verbose "No items found for $($Site.webUrl), skipping..."; continue }

    $out = [System.Collections.Generic.List[object]]::new();$i=0
    Write-Verbose "Processing a total of $($children.value.count) items for $($Site.webUrl)"
    $children = $children.value
    if (!$children) { continue }

    #Process items
    foreach ($file in $children) {
        $out.Add($(processItem -Site $Site -file $file -Verbose:$VerbosePreference))

        if ($IncludeVersions) {
            #Anti-throttling control. We don't make any additional calls unless -IncludeVersions is specified, so only add delay here
            $i++
            if ($i % 100 -eq 0) { Start-Sleep -Milliseconds 300 }
        }
    }

    #Use the comma operator to force the output as actual list instead of array
    ,($out)
}

function processItem {
    Param(
    #Graph site object
    [Parameter(Mandatory=$true)]$site,
    #File object
    [Parameter(Mandatory=$true)]$file)

    #Determine the item type
    if ($file.driveItem.package.Type -eq "OneNote") { $itemType = "Notebook" }
    elseif ($file.driveItem.file) { $itemType = "File" }
    else { $itemType = "Folder" }

    #While we can fetch versions in the initial /lists/{id}/items query, you cannot return the version file size therein. So we do a separate query, per item
    if ($IncludeVersions) { #Include version details
        if ($file.versions.count -ge 2) {
            $versions = @()
            $uri = "https://graph.microsoft.com/v1.0/sites/$($Site.id)/drives/$($file.driveItem.parentReference.driveId)/items/$($file.driveitem.id)/versions?`$select=size&`$top=999" #Seems you can go over 999, but just in case...

            do {#handle pagination
                $result = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop
                $uri = $result.'@odata.nextLink'

                $versions += $result.Value
            } while ($uri)

            if ($versions) {
                $versionSize = ($versions.size | Measure-Object -Sum).Sum
                $versionCount = $versions.count
                $versionQuota = (&{If($site.quota) {[math]::Round(100*$versionSize / $site.quota,2)} Else {"N/A"}})
            }
            else { #No versions found, can happen because /lists/{id}/items DOES return versions for some Folder/Notebook items, whereas the corresponding /drive call does NOT
                $versionSize = (&{If($file.driveItem.file) {$file.driveItem.size} Else {$null}}) #Only stamp this on files
                $versionCount = (&{If($file.driveItem.file) {1} Else {$null}}) #Only stamp this on files
                $versionQuota = (&{If(($null -ne $file.driveItem.size) -and $site.Quota) {[math]::Round(100*$file.driveItem.size / $site.quota,2)} Else {"N/A"}})
            }
        }
        else { #single version only, no point in querying
            $versionSize = (&{If($file.driveItem.file) {$file.driveItem.size} Else {$null}}) #Only stamp this on files
            $versionCount = (&{If($file.driveItem.file) {1} Else {$null}}) #Only stamp this on files
            $versionQuota = (&{If(($null -ne $file.driveItem.size) -and $site.Quota) {[math]::Round(100*$file.driveItem.size / $site.quota,2)} Else {"N/A"}})
        }
    }
    else {
        $versionQuota = (&{If(($null -ne $file.driveItem.size) -and $site.Quota) {[math]::Round(100*$file.driveItem.size / $site.quota,2)} Else {"N/A"}})
    }

    #Prepare the output object
    $fileinfo = [ordered]@{
        Site = (&{If($site.displayName) {$site.displayName} Else {$site.Name}})
        SiteURL = $site.webUrl
        Name = $file.driveItem.name
        ItemType = $itemType
        Size = (&{If($null -ne $file.driveItem.size) {$file.driveItem.size} Else {"N/A"}})
        createdDateTime = (&{If($file.driveItem.createdDateTime) {$file.driveItem.createdDateTime} Else {"N/A"}})
        lastModifiedDateTime = (&{If($file.driveItem.lastModifiedDateTime) {$file.driveItem.lastModifiedDateTime} Else {"N/A"}})
        lastModifiedBy = (&{If($file.driveItem.lastModifiedBy) { Get-Identifier $file.driveItem.lastModifiedBy } Else {"N/A"}}) #Can be missing for some items??
        Shared = (&{If($file.driveItem.shared) {"Yes"} Else {"No"}})
        ID = $file.driveItem.Id #Hide column
        InFolder = $file.driveItem.parentReference.Id #Hide column
        ItemPath = $file.webUrl
        ItemID = "https://graph.microsoft.com/v1.0/sites/$($Site.id)/drives/$($file.driveItem.parentReference.driveId)/items/$($file.driveitem.id)"
        "% of Site quota" = $versionQuota
    }
    if ($IncludeVersions -and $versionCount) { $fileinfo."VersionCount" = $versionCount }
    if ($IncludeVersions -and $versionSize) { $fileinfo."VersionSize" = $versionSize }

    #handle the output
    return [PSCustomObject]$fileinfo
}

#"Borrowed" from https://stackoverflow.com/a/42275676
function buildIndex {
    Param($array,[string]$keyName)

    $index = @{}
    foreach ($row in $array) {
        $key = $row.($keyName)
        $data = $index[$key]
        if ($data -is [Collections.ArrayList]) {
            $data.add($row) >$null
        } elseif ($data) {
            $index[$key] = [Collections.ArrayList]@($data, $row)
        } else {
            $index[$key] = $row
        }
    }
    $index
}

function Get-Identifier {
    param([Parameter(Mandatory=$true)]$Id) #Whatever Graph returns for lastModifiedBy

    #Cover additional scenarios here
    if ($Id.user) {
        if ($Id.user.email) { return $Id.user.email }
        elseif ($Id.user.displayName) { return $Id.user.displayName }
        elseif ($Id.user.id) { return $Id.user.id }
        else { return "N/A" }
    }
    else { return $Id } #catch-all
}

function Renew-Token {
    #prepare the request
    $url = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'

    $Scopes = New-Object System.Collections.Generic.List[string]
    $Scope = "https://graph.microsoft.com/.default"
    $Scopes.Add($Scope)

    $body = @{
        grant_type = "client_credentials"
        client_id = $appID
        client_secret = $client_secret
        scope = $Scopes
    }

    try {
        $authenticationResult = Invoke-WebRequest -Method Post -Uri $url -UseBasicParsing -Verbose:$false -Body $body -ErrorAction Stop
        $token = ($authenticationResult.Content | ConvertFrom-Json).access_token
        Set-Variable -Name tokenExp -Scope Global -Value $([datetime]::Now.AddSeconds(($authenticationResult.Content | ConvertFrom-Json).expires_in))
        Set-Variable -Name authHeader -Scope Global -Value @{'Authorization'="Bearer $token";'Content-Type'='application\json'}
        Write-Verbose "Successfully obtained access token, valid until $tokenExp"
    }
    catch { $_; throw }

    if (!$token) { Write-Error "Failed to aquire token!" -ErrorAction Stop; throw }
}

function Invoke-GraphApiRequest {
    param(
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$Uri,
    [bool]$RetryOnce)

    if (!$AuthHeader) { Write-Verbose "No access token found, aborting..."; throw }

    if ($MyInvocation.BoundParameters.ContainsKey("ErrorAction")) { $ErrorActionPreference = $MyInvocation.BoundParameters["ErrorAction"] }
    else { $ErrorActionPreference = "Stop" }

    try { $result = Invoke-WebRequest -Headers $AuthHeader -Uri $uri -UseBasicParsing -Verbose:$false -ErrorAction $ErrorActionPreference -ConnectionTimeoutSeconds 300 } #still getting the occasional timeout :(
    catch {
        if ($null -eq $_.Exception.Response) { throw }

        switch ($_.Exception.Response.StatusCode) {
            "TooManyRequests" { #429, throttled (Too many requests)
                if ($_.Exception.Response.Headers.'Retry-After') {
                    Write-Verbose "The request was throttled, pausing for $($_.Exception.Response.Headers.'Retry-After') seconds..."
                    Start-Sleep -Seconds $_.Exception.Response.Headers.'Retry-After'
                }
                else { Write-Verbose "The request was throttled, pausing for 10 seconds"; Start-Sleep -Seconds 10 }

                #retry the query
                $result = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference
            }
            "ResourceNotFound|Request_ResourceNotFound" { Write-Verbose "Resource $uri not found, skipping..."; return } #404, continue
            "BadRequest" { #400, we should terminate... but stupid Graph sometimes returns 400 instead of 404. And sometimes for a valid request... so likely throttling related
                if ($RetryOnce) { throw } #We already retried, terminate
                Write-Verbose "Received a Bad Request reply, retry after 10 seconds just because Graph sucks..."
                Start-Sleep -Seconds 10
                $result = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -RetryOnce
            }
            "GatewayTimeout" {
                #Do NOT retry, the error is persistent and on the server side
                Write-Verbose "The request timed out, if this happens regularly, consider increasing the timeout or updating the query to retrieve less data per run"; throw
            }
            "ServiceUnavailable" { #Should be retriable, then again, it's Microsoft...
                if ($RetryOnce) { throw } #We already retried, terminate
                if ($_.Exception.Response.Headers.'Retry-After') {
                    Write-Verbose "The request was throttled, pausing for $($_.Exception.Response.Headers.'Retry-After') seconds..."
                    Start-Sleep -Seconds $_.Exception.Response.Headers.'Retry-After'
                }
                else {
                    Write-Verbose "The service is unavailable, pausing for 10 seconds..."
                    Start-Sleep -Seconds 10
                    $result = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -RetryOnce
                }
            }
            "Forbidden" { Write-Verbose "Insufficient permissions to run the Graph API call, aborting..."; throw } #403, terminate
            "InvalidAuthenticationToken" { #Access token has expired
                if ($_.ErrorDetails.Message -match "Lifetime validation failed, the token is expired|Access token has expired") { #renew token, continue
                Write-Verbose "Access token is invalid, trying to renew..."
                Renew-Token

                if (!$AuthHeader) { Write-Verbose "Failed to renew token, aborting..."; throw }
                #Token is renewed, retry the query
                $result = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference
            }}
            default { throw }
        }
    }

    if ($result) {
        if ($result.Content) { return ($result.Content | ConvertFrom-Json) }
        else { return $result }
    }
    else { throw }
}

#==========================================================================
#Main script starts here
#==========================================================================

if ($NoItemLevelStats -and $IncludeVersions) { $Includeversions = $false; Write-Warning "The NoItemLevelStats switch is specified, disabling the IncludeVersions switch..." }

#Variables to configure
$tenantID = "tenant.onmicrosoft.com" #your tenantID or tenant root domain
$appID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #the GUID of your app. For best result, use app with Sites.ReadWrite.All scope granted.
$client_secret = "verylongsecurestring" #client secret for the app

Renew-Token

#Get a list of SPO/ODFB sites
$GraphSites = @()
if ($Sites) {#Process the list of sites provided as input
    Write-Verbose "Processing the list of sites provided as input..."
    foreach ($Site in $Sites) {
        if ($Site.Contains("/")) { $Site = $Site.Replace("https://", "").Replace("sharepoint.com", "sharepoint.com:/").TrimEnd("/") }
        $uri = "https://graph.microsoft.com/v1.0/sites/$($Site)?`$expand=drives(`$select=id,quota)"
        $result = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction SilentlyContinue
        if (!$result) { Write-Warning "Site $Site not found, skipping..."; continue }
        $GraphSites += $result
    }
}
else {#Get all sites
    Write-Verbose "Obtaining a list of all sites..."
    $uri = 'https://graph.microsoft.com/v1.0/sites?$top=999&$expand=drives($select=id,quota)' #The LIST method DOES NOT allow to expand drives :(
    do {
        $result = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop
        $uri = $result.'@odata.nextLink'

        $GraphSites += $result.Value
    } while ($uri)

    if (!$IncludeODFBsites) {#Not relevant when we use -Sites parameter
        Write-Warning "Not processing personal ODFB sites, if you want to include them, use the -IncludeODFBsites switch..."
        $GraphSites = $GraphSites | ? {$_.isPersonalSite -eq $false}
    }
}
if (!$GraphSites) { throw "No sites found, aborting..." }
$GraphSites = $GraphSites | select * -Unique #Remove duplicates
Write-Verbose "Obtained a total of $($GraphSites.count) sites"

#Processing sites
$Output = @()
$count = 1; $PercentComplete = 0;
foreach ($GraphSite in $GraphSites) {
    #Progress message
    $ActivityMessage = "Retrieving data for site $($GraphSite.displayName). Please wait..."
    $StatusMessage = ("Processing site {0} of {1}: {2}" -f $count, @($GraphSites).count, $GraphSite.webUrl)
    $PercentComplete = ($count / @($GraphSites).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #Check for any subsites and add them to the list to process later
    $cSite = @($GraphSite) #current site
    if (!$site.isPersonalSite) { #Personal sites cannot have subsites anymore... should we still check?
        $uri = "https://graph.microsoft.com/v1.0/sites/$($GraphSite.id)/sites?`$top=999&`$expand=drives(`$select=id,quota)" #Do we need pagination?
        $SubSites = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop
        if ($SubSites.value) {
            foreach ($SubSite in $SubSites.value) {
                $cSite += $SubSite
            }
        }
    }

    #Process each site
    foreach ($site in $cSite) {
        #Check for the presence of drives facet. If we used the LIST method, such will NOT be present, so we need to fetch drives separately
        #lastModifiedDateTime will also NOT be present if we used the LIST method, so make sure to populate it as well
        if ($null -eq $site.drives) {
            $uri = "https://graph.microsoft.com/v1.0/sites/$($site.id)?`$expand=drives(`$select=id,quota)"
            $siteDrives = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop

            #Add the missing properties
            $site | Add-Member -MemberType NoteProperty -Name "lastModifiedDateTime" -Value (&{If($siteDrives.lastModifiedDateTime) {$siteDrives.lastModifiedDateTime} Else {"N/A"}})
            if ($null -ne $siteDrives.drives) {
                $site | Add-Member -MemberType NoteProperty -Name "drives" -Value $siteDrives.drives
            }
            else { Write-Verbose "No drives found for site $($site.webUrl), skipping..." }
            #Fix for missing Name property
            if ($null -eq $site.Name) { $site | Add-Member -MemberType NoteProperty -Name "name" -Value (&{If($siteDrives.name) {$siteDrives.name} Else {"N/A"}}) }
        }
        #if (!$site.drives) { Write-Verbose "No drives found for site $($site.webUrl), skipping..."; continue }

        Write-Verbose "Processing site $($site.webUrl)..."
        if ($site.drives) {
            $site | Add-Member -MemberType NoteProperty -Name "Quota" -Value ($site.drives.quota.total | Sort-Object -Unique -Descending | select -First 1) #Do we need the last part?
            $site | Add-Member -MemberType NoteProperty -Name "QuotaRemaining" -Value ($site.drives.quota.remaining | Sort-Object -Unique -Descending | select -First 1) #Do we need the last part?
            $site | Add-Member -MemberType NoteProperty -Name "Size" -Value ($site.drives.quota.used | Sort-Object -Unique -Descending | select -First 1) #Do we need the last part?
        }
        else { #How to handle sites without drives?
            $site | Add-Member -MemberType NoteProperty -Name "Quota" -Value $null
            $site | Add-Member -MemberType NoteProperty -Name "QuotaRemaining" -Value $null
            $site | Add-Member -MemberType NoteProperty -Name "Size" -Value "N/A" #$null
        }

        #Add site-level details to the output object
        $siteinfo = [ordered]@{
            Site = (&{If($site.displayName) {$site.displayName} Else {$site.Name}})
            SiteURL = $site.webUrl
            Name = $site.Name
            ItemType = (&{If($site.root) {"Site (Root)"} Else {"Site"}})
            Size = $site.Size
            "% of Site quota" = (&{If($site.quota -and ($null -ne $site.size)) {[math]::Round(100 * ($site.Size) / ($site.quota),2)} else {"N/A"} })
            createdDateTime = (&{If($site.createdDateTime) {$site.createdDateTime} Else {"N/A"}})
            lastModifiedDateTime = (&{If($site.lastModifiedDateTime) {$site.lastModifiedDateTime} Else {"N/A"}})
            Shared = "N/A"
            ItemPath = $site.webUrl
            ItemID = "https://graph.microsoft.com/v1.0/sites/$($Site.id)"
        }
        $Output += [PSCustomObject]$siteinfo

        if ($NoItemLevelStats) { continue }

        #Get the set of LISTS, filter out hidden ones and those that are not document libraries
        $uri = "https://graph.microsoft.com/v1.0/sites/$(($site.id).TrimEnd("/"))/lists?`$expand=drive(`$select=id)&`$top=999" #Do we need pagination?
        $SiteLists = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop
        #No server-side filtering, so we do it here
        $SiteLists = $SiteLists.value | ? {$_.list.hidden -eq $false -and ($_.list.template -eq "documentLibrary" -or $_.list.template -eq "mySiteDocumentLibrary")}
        if (!$SiteLists) { Write-Verbose "No lists found for site $($site.webUrl), skipping..."; continue }

        #Process each list
        foreach ($list in $SiteLists) {#max page size is 5000
            Write-Verbose "Processing items for $($site.webUrl)/$($list.displayName)..."
            if (!$list.drive.id) { Write-Verbose "No drive resource returned for list $($list.id), skipping..."; continue }
            if ($IncludeVersions) { #$top is not supported within select, so we get up to 200 versions. Ideally we want $top=2, but that's not possible
                $uri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($list.id)/items?`$expand=driveItem(`$select=id,name,webUrl,parentReference,file,folder,package,shared,size,createdDateTime,lastModifiedDateTime,lastModifiedBy),versions(`$select=id)&`$select=id,driveItem,versions,webUrl&`$top=100"
            }
            else {
                $uri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($list.id)/items?`$expand=driveItem(`$select=id,name,webUrl,parentReference,file,folder,package,shared,size,createdDateTime,lastModifiedDateTime,lastModifiedBy)&`$select=id,driveItem,versions,webUrl&`$top=5000"
            }
            $pOutput = processChildren -Site $site -URI $uri

            #Correct folder size where necessary
            #Either need to process items hierarchy, or run this twice?
            if ($IncludeVersions) {#Only makes sense when we include versions
                $varIndex = buildIndex -array $pOutput -keyName "InFolder" #build index for faster lookup

                #process each folder
                $pOutput | ? {$_.ItemType -in @("Folder","Notebook")} | Sort-Object -Property {$_.ItemPath.Split("/").Count} -Descending | % {
                    $Items = $varIndex[$_.ID] #Get all items with the same path as the folder
                    $totalItemSize = $Items | % { if ($_.VersionSize) {$_.VersionSize} else {$_.Size} } | Measure-Object -Sum | Select-Object -ExpandProperty Sum

                    if ($totalItemSize) {#Check for and correct the folder size
                        if (($_.size -eq "N/A") -or ($totalItemSize -gt $_.Size)) { #maybe change to -ne... as example see the PowerShell folder
                            Write-Verbose "Correcting folder size for $($_.Name)..."
                            $_.Size = $totalItemSize
                        }
                    }

                    #Redo the '% of Site quota' calculation
                    $_."% of Site quota" = (&{If($site.Quota -and ($null -ne $_.size)) {[math]::Round(100 * ($_.Size) / ($site.Quota),2)} else {"N/A"} })
                }
            }

            #Add the updated output to the main object
            $Output += $pOutput
        }
    }
    #simple anti-throttling control
    Start-Sleep -Milliseconds 300
}

#Return the output
if (!$Output) { Write-Warning "No items found, exiting..."; return }

if ($IncludeVersions) { $Output = $Output | select Site,SiteURL,Name,ItemType,Shared,Size,VersionCount,VersionSize,'% of Site quota',createdDateTime,lastModifiedDateTime,lastModifiedBy,ItemPath,ItemID }
else { $Output = $Output | select Site,SiteURL,Name,ItemType,Shared,Size,'% of Site quota',createdDateTime,lastModifiedDateTime,lastModifiedBy,ItemPath,ItemID }

$global:varSPOSharedItems = $Output

if ($ExportToExcel) {
    Write-Verbose "Exporting the results to an Excel file..."
    # Verify module exists
    if ($null -eq (Get-Module -Name ImportExcel -ListAvailable -Verbose:$false)) {
        Write-Verbose "The ImportExcel module was not found, skipping export to Excel file..."; return
    }

    $excel = $Output `
    ` | Export-Excel -Path "$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_SPOStorageReport.xlsx" -WorksheetName StorageReport -FreezeTopRow -AutoFilter -BoldTopRow -NoHyperLinkConversion ItemID -AutoSize -PassThru

    $sheet = $excel.Workbook.Worksheets["StorageReport"]
    $sheet.Column(2).Hidden = $true #SiteURL

    #Add a hyperlink to the SiteURL
    $cells = $sheet.Cells["A:A"] #Gives just the populated cells
    foreach ($cell in $cells) {
        $cell.Hyperlink = $sheet.Cells[$cell.Address.Replace("A","B")].Text
        $lastcell = $cell.Address #needed for styling/formatting, otherwise the whole column is changed and file size goes boom
    }

    $styles = @(
        New-ExcelStyle -FontColor Blue -Underline -Range "A2:$lastcell"
    )

    #Add the Insights sheet
    $topSites = $Output | ? {$_.ItemType -in @("Site (Root)","Site")} | ? {$_.Size -ne "N/A"} | Sort-Object -Property Size -Descending | select -First 10
    $topFiles = $Output | ? {$_.ItemType -eq "File"} | Sort-Object -Property Size -Descending | select -First 10
    $topFilesV = $Output | ? {$_.ItemType -eq "File"} | Sort-Object -Property VersionSize -Descending | select -First 10
    $topFilesVC = $Output | ? {$_.ItemType -eq "File"} | Sort-Object -Property VersionCount -Descending | select -First 10

    $topSites | select Site,Size,'% of Site quota',SiteURL | Export-Excel -ExcelPackage $excel -WorksheetName "Insights" -TableName "TopSites" -TableStyle Dark8 -StartRow 2 -AutoSize -PassThru > $null
    $sheet2 = $excel.Workbook.Worksheets["Insights"]
    Set-Format -Worksheet $sheet2 -Range A1 -Value "Top 10 Sites" -Bold
    if (!$NoItemLevelStats) {
        if ($topFiles) {
            $topFiles | select Name,Size,'% of Site quota',ItemPath | Export-Excel -ExcelPackage $excel -WorksheetName "Insights" -TableName "TopFiles" -TableStyle Dark8 -StartRow 15 -AutoSize -PassThru > $null
            Set-Format -Worksheet $sheet2 -Range A14 -Value "Top 10 Files by size" -Bold
        }
        if ($topFilesV) {
            $topFilesV | select Name,VersionSize,'% of Site quota',ItemPath | Export-Excel -ExcelPackage $excel -WorksheetName "Insights" -TableName "TopFilesWithVersions" -TableStyle Dark8 -StartRow 28 -PassThru > $null
            Set-Format -Worksheet $sheet2 -Range A27 -Value "Top 10 Files by size with versions included" -Bold
        }
        if ($topFilesVC) {
            $topFilesVC | select Name,VersionCount,'% of Site quota',ItemPath | Export-Excel -ExcelPackage $excel -WorksheetName "Insights" -TableName "TopFilesVersionCount" -TableStyle Dark8 -StartRow 41 -PassThru > $null
            Set-Format -Worksheet $sheet2 -Range A40 -Value "Top 10 Files by number of versions" -Bold
        }
    }

    #Save the changes
    Export-Excel -ExcelPackage $excel -WorksheetName "StorageReport" -Style $styles -Show
    Write-Verbose "Excel file exported successfully..."
}
else {
    $Output | Export-Csv -Path "$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_SPOStorageReport.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
    Write-Verbose "Results exported to ""$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_SPOStorageReport.csv""."
}
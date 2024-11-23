#Requires -Version 7.4
# Make sure to fill in all the required variables before running the script
# Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Sites.Read.All to return all the item sharing details
#    (optional) Directory.Read.All to obtain a domain list and check whether an item is shared externally

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/6154/report-on-externally-shared-files-via-the-graph-api

[CmdletBinding()] #Make sure we can use -Verbose
Param(
[string[]][ValidateNotNullOrEmpty()]$Sites, #Use the Sites parameter to specify a set of sites to process.
[switch]$IncludeODFBsites, #Use the IncludeODFBsites switch to specify whether to include personal OneDrive for Business sites in the output.
[switch]$IncludeExpired, #Use the IncludeExpired switch to include expired sharing links in the output.
[switch]$IncludeOwner, #Use the IncludeOwner switch to include Site collection admin/secondary admin entries in the output.
[switch]$ExpandFolders, #Use the ExpandFolders switch to specify whether to expand folders recursively, otherwise covers only the root.
[int]$Depth, #Use the Depth parameter to specify the folder depth for expansion/inclusion of items. 1 is just root, 2 is root+children, etc.
[switch]$ExportToExcel #Use the ExportToExcel switch to specify whether to export the output to an Excel file.
)

function processChildren {
    Param(
    #Graph Site object
    [Parameter(Mandatory=$true)]$Site,
    #URI for the drive
    [Parameter(Mandatory=$true)][string]$URI,
    #Use the ExpandFolders switch to specify whether to expand folders and include their items in the output.
    [switch]$ExpandFolders,
    #Use the Depth parameter to specify the folder depth for expansion/inclusion of items.
    [int]$depth)

    $URI = "$URI/children?`$top=5000"
    $children = @()
    Write-Verbose "Processing children for $($Site.webUrl)..."
    #fetch children, make sure to handle multiple pages
    do {
        $result = Invoke-GraphApiRequest -Uri "$URI" -Verbose:$VerbosePreference
        $URI = $result.'@odata.nextLink'
        #If we are getting multiple pages, add some delay to avoid throttling
        Start-Sleep -Milliseconds 500
        $children += $result
    } while ($URI)
    #if (!$children.value) { Write-Verbose "No items found for $($Site.webUrl), skipping..."; continue } #DO NOT SKIP, we might need to handle the parent folder if called recursively!

    #Process items
    $output = @()
    Write-Verbose "Processing a total of $($children.value.count) items for $($Site.webUrl), of which $(($children.value | ? {$_.shared}).count) are shared..."
    #We do NOT filter for the shared facet here, as the parent folder might not be shared, while the child items are. Example, "Shared with everyone" folder
    $children = $children.value #| ? {$_.shared}

    #handle different children types
    $cFolders = $children | ? {$_.Folder} #no filter for shared facet here!
    $cFiles = $children | ? {$_.File -and $_.shared} #doesnt return notebooks
    $cNotebooks = $children | ? {$_.package.type -eq "OneNote" -and $_.shared}

    #Process Folders
    foreach ($folder in $cFolders) {
        $output += (processFolder -Site $Site -folder $folder -ExpandFolders:$ExpandFolders -depth $depth -Verbose:$VerbosePreference)
    }

    #Process Files
    foreach ($file in $cFiles) {
        $output += (processFile -Site $Site -file $file -Verbose:$VerbosePreference)
    }

    #Process Notebooks
    foreach ($notebook in $cNotebooks) {
        $output += (processFile -site $Site -file $notebook -Verbose:$VerbosePreference)
    }

    return $output
}

function processFolder {
    Param(
    #Graph Site object
    [Parameter(Mandatory=$true)]$Site,
    #Folder object
    [Parameter(Mandatory=$true)]$folder,
    #Use the ExpandFolders switch to specify whether to expand folders and include their items in the output.
    [switch]$ExpandFolders,
    #Use the Depth parameter to specify the folder depth for expansion/inclusion of items.
    [int]$depth)

    #prepare the output object
    $fileinfo = New-Object psobject
    $fileinfo | Add-Member -MemberType NoteProperty -Name "Site" -Value $Site.displayName
    $fileinfo | Add-Member -MemberType NoteProperty -Name "SiteURL" -Value $Site.webUrl
    $fileinfo | Add-Member -MemberType NoteProperty -Name "Name" -Value $folder.name
    $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemType" -Value "Folder"
    $fileinfo | Add-Member -MemberType NoteProperty -Name "Shared" -Value (&{If($folder.shared) {"Yes"} Else {"No"}})

    #if the Shared property is set, fetch permissions
    if ($folder.shared) {
        $permlist = getPermissions -SiteId $Site.id -DriveId $folder.parentReference.driveId -ItemId $folder.id -Verbose:$VerbosePreference
        if (!$permlist) { return } #No permissions found, skip the item. Or try with -IncludeOwner

        #Match entries against the list of domains in the tenant to populate the ExternallyShared value
        $regexmatches = $permlist | % {if ($_ -match "\(?\w+([-+.']\w+)*(#EXT#)?@\w+([-.]\w+)*\.\w+([-.]\w+)*\)?") {$Matches[0]}} #Updated to match Guest user UPNs
        if ($permlist -match "anonymous") { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "Yes" }
        else {
            if (!$domains) { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "No domain info" }
            elseif ($regexmatches -match "#EXT#") { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "Yes" }
            elseif ($regexmatches -notmatch ($domains -join "|")) { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "Yes" }
            else { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "No" }
        }
        $fileinfo | Add-Member -MemberType NoteProperty -Name "Permissions" -Value ($permlist -join ",")
    }
    $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemPath" -Value $folder.webUrl
    $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemID" -Value "https://graph.microsoft.com/v1.0/sites/$($Site.id)/drives/$($folder.parentReference.driveId)/items/$($folder.id)"

    #Since this is a folder item, check for any children, depending on the script parameters
    if (($folder.folder.childCount -gt 0) -and $ExpandFolders -and ($depth + 4 - $folder.parentReference.path.Split("/").Count) -gt 1) { #/drives/{driveId}/root: gives the 4
        Write-Verbose "Folder $($folder.Name) has child items"
        $uri = "https://graph.microsoft.com/v1.0/sites/$($Site.id)/drives/$($folder.parentReference.driveId)/items/$($folder.id)"
        $folderItems = processChildren -Site $Site -URI $uri -ExpandFolders:$ExpandFolders -depth $depth -Verbose:$VerbosePreference
    }

    #handle the output
    if ($folderItems) { $f = @(); $f += $fileinfo; $f += $folderItems; return $f }
    else { return $fileinfo }
}

function processFile {
    Param(
    #Graph site object
    [Parameter(Mandatory=$true)]$site,
    #File object
    [Parameter(Mandatory=$true)]$file)

    #prepare the output object
    $fileinfo = New-Object psobject
    $fileinfo | Add-Member -MemberType NoteProperty -Name "Site" -Value $site.displayName
    $fileinfo | Add-Member -MemberType NoteProperty -Name "SiteURL" -Value $Site.webUrl
    $fileinfo | Add-Member -MemberType NoteProperty -Name "Name" -Value $file.name
    $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemType" -Value (&{If($file.package.Type -eq "OneNote") {"Notebook"} Else {"File"}})
    $fileinfo | Add-Member -MemberType NoteProperty -Name "Shared" -Value (&{If($file.shared) {"Yes"} Else {"No"}})

    #if the Shared property is set, fetch permissions
    if ($file.shared) {
        $permlist = getPermissions -SiteId $site.id -DriveId $file.parentReference.driveId -ItemId $file.id -Verbose:$VerbosePreference
        if (!$permlist) { return } #No permissions found, skip the item. Or try with -IncludeOwner

        #Match entries against the list of domains in the tenant to populate the ExternallyShared value
        $regexmatches = $permlist | % {if ($_ -match "\(?\w+([-+.']\w+)*(#EXT#)?@\w+([-.]\w+)*\.\w+([-.]\w+)*\)?") {$Matches[0]}} #Updated to match Guest user UPNs
        if ($permlist -match "anonymous") { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "Yes" }
        else {
            if (!$domains) { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "No domain info" }
            elseif ($regexmatches -match "#EXT#") { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "Yes" }
            elseif ($regexmatches -notmatch ($domains -join "|")) { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "Yes" }
            else { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "No" }
        }
        $fileinfo | Add-Member -MemberType NoteProperty -Name "Permissions" -Value ($permlist -join ",")
    }
    $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemPath" -Value $file.webUrl
    $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemID" -Value "https://graph.microsoft.com/v1.0/sites/$($Site.id)/drives/$($file.parentReference.driveId)/items/$($file.id)"

    #handle the output
    return $fileinfo
}

function getPermissions {
    Param(
    #Use the SiteId parameter to provide an unique identifier for the site object.
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$SiteId,
    #Use the DriveId parameter to provide an unique identifier for the drive object.
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$DriveId,
    #Use the ItemId parameter to provide an unique identifier for the item object.
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$ItemId)

    #Check if the token is about to expire and renew if needed
    if ($tokenExp -lt [datetime]::Now.AddSeconds(360)) {
        Write-Verbose "Access token is about to expire, renewing..."
        Renew-Token
    }

    #fetch permissions for the given item. Add pagination support?
    $uri = "https://graph.microsoft.com/beta/sites/$($SiteId)/drives/$($DriveId)/items/$($ItemId)/permissions?`$top=999"
    $permissions = (Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference).Value

    #build the permissions string
    $permlist = @()
    foreach ($entry in $permissions) {
        #Sharing link
        if ($entry.link) {
            #blocksDownload is undocumented type, not sure if working as expected. In any case, we need the actual role reflected
            if ($entry.link.type -eq "blocksDownload") { $strPermissions = $($entry.roles) + ":" + $($entry.link.scope) }
            else { $strPermissions = $($entry.link.type) + ":" + $($entry.link.scope) }
            if ($entry.grantedToIdentitiesV2) { $strPermissions = $strPermissions + " (" + (((&{If($entry.grantedToIdentitiesV2.siteUser.email) {$entry.grantedToIdentitiesV2.siteUser.email} else {$entry.grantedToIdentitiesV2.User.email}}) | select -Unique) -join ",") + ")" }
            if ($entry.hasPassword) { $strPermissions = $strPermissions + "[PasswordProtected]" }
            if ($entry.link.preventsDownload) { $strPermissions = $strPermissions + "[BlockDownloads]" }
            if ($entry.expirationDateTime) {
                if ($entry.expirationDateTime -lt [datetime]::Now) {
                    if ($IncludeExpired) { $strPermissions = $strPermissions + " (Expired on: $($entry.expirationDateTime))" }
                    else { continue }
                }
                else { $strPermissions = $strPermissions + " (Expires on: $($entry.expirationDateTime))" }
            }
            $permlist += $strPermissions
        }
        #Invitation
        elseif ($entry.invitation) { $permlist += $($entry.roles) + ":" + $($entry.invitation.email) }
        #Direct permissions
        elseif ($null -ne $entry.roles) {
            if (!$entry.grantedToV2) { $roleentry = "Unknown"; continue }

            $facets = $entry.grantedToV2.psobject.properties | ? {$_.MemberType -eq "NoteProperty"} | select Name,value
            foreach ($t in $facets) {
                $roleentry = switch ($t.Name) {
                    "siteUser" {
                        if ($t.value.email) { $t.value.email }
                        elseif ($t.value.loginName) { $t.value.loginName }
                        else { $t.value.displayName }
                        break
                    }
                    "User" { $t.value.email; break }
                    "siteGroup" { $t.value.displayName; break }
                    "group" { $t.value.displayName; break }
                    "application" { $t.value.Id; break }
                    default { $t.value.loginName; break }
                }
                if ($roleentry) { break }
            }

            #This one seems new... and stupid
            if ("" -eq $entry.roles) { $permlist += "Restricted view:" + $roleentry } #Restricted view/View Only/God knows what else
            else { $permlist += $($entry.Roles) + ':' + $roleentry }
        }
        #Inherited permissions. Useless...
        elseif ($entry.inheritedFrom.path) { $permlist += "[Inherited from: $($entry.inheritedFrom.path)]" } #If only Graph populated these...
        #ShareId... perhaps add it to the output?
        elseif ($entry.shareId) {} #do nothing
        #some other permissions?
        else { Write-Verbose "Permission $entry not covered by the script!"; $permlist += $entry }
    }

    #handle the output
    if (!$IncludeOwner) { return ($permlist | ? {$_ -notmatch 'owner:'}) }
    else { return $permlist }
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
        $authenticationResult = Invoke-WebRequest -Method Post -Uri $url -Debug -Verbose:$false -Body $body -ErrorAction Stop
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

    try { $result = Invoke-WebRequest -Headers $AuthHeader -Uri $uri -Verbose:$false -ErrorAction $ErrorActionPreference -ConnectionTimeoutSeconds 300 } #still getting the occasional timeout :(
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

#Variables to configure
$tenantID = "tenant.onmicrosoft.com" #your tenantID or tenant root domain
$appID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #the GUID of your app. For best result, use app with Sites.ReadWrite.All scope granted.
$client_secret = "verylongsecurestring" #client secret for the app

Renew-Token

Write-Verbose "Obtaining a list of verified domains..."
#Used to determine whether sharing is done externally, needs Directory.Read.All scope.
$domains = (Invoke-GraphApiRequest -Uri "https://graph.microsoft.com/v1.0/domains" -Verbose:$VerbosePreference).Value | ? {$_.IsVerified -eq "True"} | select -ExpandProperty Id
#$domains = @("xxx.com","yyy.com")
if (!$domains) { Write-Verbose "No verified domains found, skipping external sharing checks..." }
else { Write-Verbose "The following list of domains will be used for external sharing checks: $($domains -join ", ")" }

#Adjust the input parameters
if ($ExpandFolders -and ($depth -le 1)) { $depth = 1 }

#Get a list of SPO/ODFB sites
$GraphSites = @()
if ($Sites) {#Process the list of sites provided as input
    Write-Verbose "Processing the list of sites provided as input..."
    foreach ($Site in $Sites) {
        if ($Site.Contains("/")) { $Site = $Site.Replace("https://", "").Replace("sharepoint.com", "sharepoint.com:/").TrimEnd("/") }
        $uri = "https://graph.microsoft.com/v1.0/sites/$Site"
        $result = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction SilentlyContinue
        if (!$result) { Write-Warning "Site $Site not found, skipping..."; continue }
        $GraphSites += $result
    }
}
else {#Get all sites
    Write-Verbose "Obtaining a list of all sites..."
    $uri = 'https://graph.microsoft.com/v1.0/sites?$top=999'
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
        $uri = "https://graph.microsoft.com/v1.0/sites/$($GraphSite.id)/sites?`$top=999" #Do we need pagination?
        $SubSites = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop
        if ($SubSites.value) {
            foreach ($SubSite in $SubSites.value) {
                $cSite += $SubSite
            }
        }
    }

    #Process each site
    foreach ($site in $cSite) {
        Write-Verbose "Processing site $($site.webUrl)..."
        #Get the set of Drives, filter out hidden ones and those that are not document libraries
        $uri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives?`$expand=list(`$select=id,list)&`$select=id,webUrl&`$top=999" #Do we need pagination?
        $SiteDrives = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop
        #No server-side filtering, so we do it here
        $SiteDrives = $SiteDrives.value | ? {$_.list.list.hidden -eq $false -and ($_.list.list.template -eq "documentLibrary" -or $_.list.list.template -eq "mySiteDocumentLibrary")}
        if (!$SiteDrives) { Write-Verbose "No lists found for site $($site.webUrl), skipping..."; continue }

        #Process each drive
        foreach ($SiteDrive in $SiteDrives) {#max page size is 5000
            #Get the root folder on the drive
            $uri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/drives/$($SiteDrive.id)/root"
            $SiteDriveRoot = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop

            #If no items in the drive, skip
            if (!$SiteDriveRoot -or ($SiteDriveRoot.folder.childCount -eq 0)) { Write-Verbose "No items to report on for site $($SiteDrive.webUrl), skipping..."; continue }

            #Enumerate items in the drive and prepare the output
            $pOutput = processChildren -Site $site -URI $uri -ExpandFolders:$ExpandFolders -depth $depth
            $Output += $pOutput
        }
    }

    #simple anti-throttling control
    Start-Sleep -Milliseconds 300
}

#Return the output
if (!$Output) { Write-Warning "No shared items found, exiting..."; return }
$global:varSPOSharedItems = $Output | select Site,SiteURL,Name,ItemType,Shared,ExternallyShared,Permissions,ItemPath,ItemID | ? {$_.Shared -eq "Yes"}

if ($ExportToExcel) {
    Write-Verbose "Exporting the results to an Excel file..."
    # Verify module exists
    if ($null -eq (Get-Module -Name ImportExcel -ListAvailable -Verbose:$false)) {
        Write-Verbose "The ImportExcel module was not found, skipping export to Excel file..."; return
    }

    $excel = $Output | ? {$_.Shared -eq "Yes"} | select Site,SiteURL,Name,ItemType,Shared,ExternallyShared,Permissions,ItemPath,ItemID `
    ` | Export-Excel -Path "$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_SPOSharedItems.xlsx" -WorksheetName SharedFiles -FreezeTopRow -AutoFilter -BoldTopRow -NoHyperLinkConversion Permissions,ItemID -AutoSize -PassThru

    $sheet = $excel.Workbook.Worksheets["SharedFiles"]
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

    #Add conditional formatting for the ExternallyShared column
    Add-ConditionalFormatting -Worksheet $sheet -Range "F2:$($lastcell.Replace("A","F"))" -RuleType Equal -ConditionValue "Yes" -ForegroundColor White -BackgroundColor Red

    #Add the summary sheet
    $Output | ? {$_.Shared -eq "Yes"} | group Site | select @{Name="Site";Expression={$_.Name}}, @{Name="Shared files";Expression={$_.Count}}, @{Name="Externally shared";e={($_.Group | ? {$_.ExternallyShared -eq "Yes"}).count}} `
    ` | Export-Excel -ExcelPackage $excel -WorksheetName "Summary" -AutoSize -FreezeTopRow -BoldTopRow -PassThru

    #Save the changes
    Export-Excel -ExcelPackage $excel -WorksheetName "SharedFiles" -Style $styles -Show
    Write-Verbose "Excel file exported successfully..."
}
else {
    $Output | ? {$_.Shared -eq "Yes"} | select Site,SiteURL,Name,ItemType,Shared,ExternallyShared,Permissions,ItemPath,ItemID | Export-Csv -Path "$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_SPOSharedItems.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
    Write-Verbose "Results exported to ""$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_SPOSharedItems.csv""."
}
#Requires -Version 3.0
# Make sure to fill in all the required variables before running the script
# Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Sites.Read.All to return all the item sharing details
#    (optional) Directory.Read.All to obtain a domain list and check whether an item is shared externally

[CmdletBinding()] #Make sure we can use -Verbose
Param(
[string[]][ValidateNotNullOrEmpty()]$Sites, #Use the Sites parameter to specify a set of sites to process.
[switch]$IncludeODFBsites, #Use the IncludeODFBsites switch to specify whether to include personal OneDrive for Business sites in the output.
[switch]$IncludeExpired, #Use the IncludeExpired switch to include expired sharing links in the output.
[switch]$IncludeOwner, #Use the IncludeOwner switch to include Site collection admin/secondary admin entries in the output.
[switch]$ExportToExcel #Use the ExportToExcel switch to specify whether to export the output to an Excel file.
)

function processChildren {
    Param(
    #Graph Site object
    [Parameter(Mandatory=$true)]$Site,
    #URI for the drive
    [Parameter(Mandatory=$true)][string]$URI)

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
    if (!$children.value) { Write-Verbose "No items found for $($Site.webUrl), skipping..."; continue }

    $output = @();$i=0
    Write-Verbose "Processing a total of $($children.value.count) items for $($Site.webUrl), of which $(($children.value.driveItem | ? {$_.shared}).count) are shared..."
    $children = $children.value | ? {$_.driveItem.shared}
    if (!$children) { continue }

    #Process items
    foreach ($file in $children) {
        $output += (processItem -Site $Site -file $file -Verbose:$VerbosePreference)

        #Anti-throttling control
        $i++
        if ($i % 100 -eq 0) { Start-Sleep -Milliseconds 500 }
    }

    return $output
}

function processItem {
    Param(
    #Graph site object
    [Parameter(Mandatory=$true)]$site,
    #File object
    [Parameter(Mandatory=$true)]$file)

    #Prepare the output object
    $fileinfo = New-Object psobject
    $fileinfo | Add-Member -MemberType NoteProperty -Name "Site" -Value $site.displayName
    $fileinfo | Add-Member -MemberType NoteProperty -Name "SiteURL" -Value $Site.webUrl
    $fileinfo | Add-Member -MemberType NoteProperty -Name "Name" -Value $file.driveItem.name
    #Determine the item type
    if ($file.driveItem.package.Type -eq "OneNote") { $itemType = "Notebook" }
    elseif ($file.driveItem.file) { $itemType = "File" }
    else { $itemType = "Folder" }
    $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemType" -Value $itemType
    $fileinfo | Add-Member -MemberType NoteProperty -Name "Shared" -Value (&{If($file.driveItem.shared) {"Yes"} Else {"No"}})

    #If the Shared property is set, fetch permissions
    if ($file.driveItem.shared) {
        $permlist = getPermissions ("https://graph.microsoft.com/beta/sites/{0}/drives/{1}/items/{2}" -f $site.id, $file.driveItem.parentReference.driveId, $file.driveItem.id) -Verbose:$VerbosePreference
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
    $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemID" -Value "https://graph.microsoft.com/v1.0/sites/$($Site.id)/drives/$($file.driveItem.parentReference.driveId)/items/$($file.driveitem.id)"

    #handle the output
    return $fileinfo
}

function getPermissions {
    Param(
    #Use the ItemId parameter to provide an unique identifier for the item object.
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$ItemURI)

    #Check if the token is about to expire and renew if needed
    if ($tokenExp -lt [datetime]::Now.AddSeconds(360)) {
        Write-Verbose "Access token is about to expire, renewing..."
        Renew-Token
    }

    #fetch permissions for the given item. Add pagination support?
    $uri = "$ItemURI/permissions?`$top=999"
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
        elseif ($entry.roles) {
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

            $permlist += $($entry.Roles) + ':' + $roleentry
        }
        #Inherited permissions. Useless...
        elseif ($entry.inheritedFrom) { $permlist += "[Inherited from: $($entry.inheritedFrom.path)]" } #If only Graph populated these...
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
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$Uri
    )

    if (!$AuthHeader) { Write-Verbose "No access token found, aborting..."; throw }

    try { $result = Invoke-WebRequest -Headers $AuthHeader -Uri $uri -Verbose:$false -ErrorAction Stop }
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
            "BadRequest" { throw } #400, we should terminate... but stupid Graph sometimes returns 400 instead of 404
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

#Get a list of SPO/ODFB sites
$GraphSites = @()
if ($Sites) {#Process the list of sites provided as input
    Write-Verbose "Processing the list of sites provided as input..."
    foreach ($Site in $Sites) {
        if ($Site.Contains("/")) { $Site = $Site.Replace("https://", "").Replace("sharepoint.com", "sharepoint.com:/").TrimEnd("/") }
        $uri = "https://graph.microsoft.com/v1.0/sites/$Site"
        $result = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop
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
Write-Verbose "Obtained a total of $($GraphSites.count) sites"

#Processing sites
$Output = @()
$count = 1; $PercentComplete = 0;
foreach ($site in $GraphSites) {
    #Progress message
    $ActivityMessage = "Retrieving data for site $($site.displayName). Please wait..."
    $StatusMessage = ("Processing site {0} of {1}: {2}" -f $count, @($GraphSites).count, $site.webUrl)
    $PercentComplete = ($count / @($GraphSites).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    Write-Verbose "Processing site $($site.webUrl)..."
    $uri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists?`$expand=drive(`$select=id)&`$top=999" #Do we need pagination?
    $SiteLists = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop
    #No server-side filtering, so we do it here
    $SiteLists = $SiteLists.value | ? {$_.list.hidden -eq $false -and ($_.list.template -eq "documentLibrary" -or $_.list.template -eq "mySiteDocumentLibrary")}
    if (!$SiteLists) { Write-Verbose "No lists found for site $($site.webUrl), skipping..."; continue }

    #Process each list
    foreach ($list in $SiteLists) {#max page size is 5000
        Write-Verbose "Processing items for $($Site.webUrl)/$($list.displayName)..."
        if (!$list.drive.id) { Write-Verbose "No drive resource returned for list $($list.id), skipping..."; continue }
        $uri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($list.id)/items?`$expand=driveItem(`$select=id,name,webUrl,parentReference,file,folder,package,shared)&`$select=id,webUrl&`$top=5000"
        $pOutput = processChildren -Site $site -URI $uri
        $Output += $pOutput
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
    if ($null -eq (Get-Module -Name ImportExcel -ListAvailable)) {
        Write-Verbose "The ImportExcel module was not found, skipping export to Excel file..."; return
    }

    $excel = $Output | select Site,SiteURL,Name,ItemType,Shared,ExternallyShared,Permissions,ItemPath,ItemID `
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
    $Output | group Site | select @{Name="Site";Expression={$_.Name}}, @{Name="Shared files";Expression={$_.Count}}, @{Name="Externally shared";e={($_.Group | ? {$_.ExternallyShared -eq "Yes"}).count}} `
    ` | Export-Excel -ExcelPackage $excel -WorksheetName "Summary" -AutoSize -FreezeTopRow -BoldTopRow -PassThru

    #Save the changes
    Export-Excel -ExcelPackage $excel -WorksheetName "SharedFiles" -Style $styles -Show
    Write-Verbose "Excel file exported successfully..."
}
else {
    $Output | select Site,SiteURL,Name,ItemType,Shared,ExternallyShared,Permissions,ItemPath,ItemID | Export-Csv -Path "$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_SPOSharedItems.csv" -NoTypeInformation -Encoding UTF8 -UseCulture 
    Write-Verbose "Results exported to ""$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_SPOSharedItems.csv""."
}
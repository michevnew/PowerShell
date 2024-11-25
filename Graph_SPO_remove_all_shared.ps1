#Requires -Version 7.4
# Make sure to fill in all the required variables before running the script
# Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Sites.ReadWrite.All to return all the item sharing details and remove permissions

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/6308/remove-all-sharing-links-and-permissions-for-items-in-sharepoint-online-or-onedrive-for-business

[CmdletBinding()] #Make sure we can use -Verbose
Param(
[string[]][ValidateNotNullOrEmpty()]$Sites, #Use the Sites parameter to specify a set of sites to process.
[switch]$IncludeODFBsites #Use the IncludeODFBsites switch to specify whether to include personal OneDrive for Business sites in the output.
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
    $children = $children.value | ? {$_.driveItem.shared} #Only process shared items
    if (!$children) { continue }
    Write-Verbose "Processing a total of $(($children).count) shared items for $($Site.webUrl)..."

    #Process items
    foreach ($file in $children) {
        Write-Verbose "Found shared file ($($file.driveItem.name)), removing permissions..."
        RemovePermissions ("https://graph.microsoft.com/beta/sites/{0}/drives/{1}/items/{2}" -f $site.id, $file.driveItem.parentReference.driveId, $file.driveItem.id) -Verbose:$VerbosePreference

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
        $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemPath" -Value $file.webUrl
        $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemID" -Value "https://graph.microsoft.com/v1.0/sites/$($Site.id)/drives/$($file.driveItem.parentReference.driveId)/items/$($file.driveitem.id)"
        $output += $fileinfo

        #Anti-throttling control
        $i++
        if ($i % 100 -eq 0) { Start-Sleep -Milliseconds 500 }
    }

    return $output
}

function RemovePermissions {

    Param(
    #Use the ItemId parameter to provide an unique identifier for the item object.
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$ItemURI)

    #Fetch permissions for the given item. Pagination?
    $uri = "$ItemURI/permissions?`$top=999"
    $permissions = (Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference).Value

    foreach ($entry in $permissions) {
        #Check if the permission is external
        #if (($entry.link.scope -eq "anonymous") -or (($entry | ConvertFrom-Json -Depth 5) -match "#EXT#")) { #do stuff }

        $uri = "$ItemURI/permissions/$($entry.id)"

        try { Invoke-WebRequest -Method DELETE -Verbose:$false -Uri $uri -Headers $authHeader -SkipHeaderValidation -ErrorAction Stop | Out-Null }
        catch { Write-Verbose "Failed to remove permission entry $entry for $ItemURI"; continue }

        #Anti-throttling control
        $i++
        if ($i % 100 -eq 0) { Start-Sleep -Milliseconds 500 }
    }
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

    try { $result = Invoke-WebRequest -Headers $AuthHeader -Uri $uri -Verbose:$false -ErrorAction $ErrorActionPreference -ConnectionTimeoutSeconds 300 }
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
        #Get the set of LISTS, filter out hidden ones and those that are not document libraries
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
    }

    #simple anti-throttling control
    Start-Sleep -Milliseconds 300
}

#Return the output
if (!$Output) { Write-Warning "No shared items found, exiting..."; return }
$global:varSPOSharedItems = $Output | select Site,SiteURL,Name,ItemType,ItemPath,ItemID | ? {$_.Shared -eq "Yes"}

$Output | select Site,SiteURL,Name,ItemType,ItemPath,ItemID | Export-Csv -Path "$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_SPOSharedItems.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
Write-Verbose "Results exported to ""$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_SPOSharedItems.csv""."
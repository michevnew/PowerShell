#Requires -Version 7.4
# Make sure to fill in all the required variables before running the script
# Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    User.Read.All to enumerate all users in the tenant
#    Sites.ReadWrite.All to return all the item sharing details
#    (optional) Directory.Read.All to obtain a domain list and check whether an item is shared externally


[CmdletBinding()] #Make sure we can use -Verbose
Param([switch]$ExpandFolders,[int]$depth)

function processChildren {

    Param(
    #Graph User object
    [Parameter(Mandatory=$true)]$User,
    #URI for the drive
    [Parameter(Mandatory=$true)][string]$URI,
    #Use the ExpandFolders switch to specify whether to expand folders and include their items in the output.
    [switch]$ExpandFolders,
    #Use the Depth parameter to specify the folder depth for expansion/inclusion of items.
    [int]$depth)

    $URI = "$URI/children"
    $children = @()
    #fetch children, make sure to handle multiple pages
    do {
        $result = Invoke-GraphApiRequest -Uri "$URI" -Verbose:$VerbosePreference
        $URI = $result.'@odata.nextLink'
        #If we are getting multiple pages, add some delay to avoid throttling
        Start-Sleep -Milliseconds 500
        $children += $result
    } while ($URI)
    if (!$children) { Write-Verbose "No items found for $($user.userPrincipalName), skipping..."; continue }

    #handle different children types
    $output = @()
    $cFolders = $children.value | ? {$_.Folder}
    $cFiles = $children.value | ? {$_.File} #doesnt return notebooks
    $cNotebooks = $children.value | ? {$_.package.type -eq "OneNote"}

    #Process Folders
    foreach ($folder in $cFolders) {
        $output += (processFolder -User $User -folder $folder -ExpandFolders:$ExpandFolders -depth $depth -Verbose:$VerbosePreference)
    }

    #Process Files
    foreach ($file in $cFiles) {
        $output += (processFile -User $User -file $file -Verbose:$VerbosePreference)
    }

    #Process Notebooks
    foreach ($notebook in $cNotebooks) {
        $output += (processFile -User $User -file $notebook -Verbose:$VerbosePreference)
    }

    return $output
}

function processFolder {

    Param(
    #Graph User object
    [Parameter(Mandatory=$true)]$User,
    #Folder object
    [Parameter(Mandatory=$true)]$folder,
    #Use the ExpandFolders switch to specify whether to expand folders and include their items in the output.
    [switch]$ExpandFolders,
    #Use the Depth parameter to specify the folder depth for expansion/inclusion of items.
    [int]$depth)

    #prepare the output object
    $fileinfo = New-Object psobject
    $fileinfo | Add-Member -MemberType NoteProperty -Name "OneDriveOwner" -Value $user.userPrincipalName
    $fileinfo | Add-Member -MemberType NoteProperty -Name "Name" -Value $folder.name
    $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemType" -Value "Folder"
    $fileinfo | Add-Member -MemberType NoteProperty -Name "Shared" -Value (&{If($folder.shared) {"Yes"} Else {"No"}})

    #if the Shared property is set, fetch permissions
    if ($folder.shared) {
        $permlist = getPermissions $user.id $folder.id -Verbose:$VerbosePreference

        #Match user entries against the list of domains in the tenant to populate the ExternallyShared value
        $regexmatches = $permlist | % {if ($_ -match "\(?\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*\)?") {$Matches[0]}}
        if ($permlist -match "anonymous") { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "Yes" }
        else {
            if (!$domains) { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "No domain info" }
            elseif ($regexmatches -notmatch ($domains -join "|")) { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "Yes" }
            else { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "No" }
        }
        $fileinfo | Add-Member -MemberType NoteProperty -Name "Permissions" -Value ($permlist -join ",")
    }
    $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemPath" -Value $folder.webUrl

    #Since this is a folder item, check for any children, depending on the script parameters
    if (($folder.folder.childCount -gt 0) -and $ExpandFolders -and ((3 - $folder.parentReference.path.Split("/").Count + $depth) -gt 0)) {
        Write-Verbose "Folder $($folder.Name) has child items"
        $uri = "https://graph.microsoft.com/v1.0/users/$($user.id)/drive/items/$($folder.id)"
        $folderItems = processChildren -User $user -URI $uri -ExpandFolders:$ExpandFolders -depth $depth -Verbose:$VerbosePreference
    }

    #handle the output
    if ($folderItems) { $f = @(); $f += $fileinfo; $f += $folderItems; return $f }
    else { return $fileinfo }
}

function processFile {

    Param(
    #Graph User object
    [Parameter(Mandatory=$true)]$User,
    #File object
    [Parameter(Mandatory=$true)]$file)

    #prepare the output object
    $fileinfo = New-Object psobject
    $fileinfo | Add-Member -MemberType NoteProperty -Name "OneDriveOwner" -Value $user.userPrincipalName
    $fileinfo | Add-Member -MemberType NoteProperty -Name "Name" -Value $file.name
    $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemType" -Value (&{If($file.package.Type -eq "OneNote") {"Notebook"} Else {"File"}})
    $fileinfo | Add-Member -MemberType NoteProperty -Name "Shared" -Value (&{If($file.shared) {"Yes"} Else {"No"}})

    #if the Shared property is set, fetch permissions
    if ($file.shared) {
        $permlist = getPermissions $user.id $file.id -Verbose:$VerbosePreference

        #Match user entries against the list of domains in the tenant to populate the ExternallyShared value
        $regexmatches = $permlist | % {if ($_ -match "\(?\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*\)?") {$Matches[0]}}
        if ($permlist -match "anonymous") { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "Yes" }
        else {
            if (!$domains) { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "No domain info" }
            elseif ($regexmatches -notmatch ($domains -join "|")) { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "Yes" }
            else { $fileinfo | Add-Member -MemberType NoteProperty -Name "ExternallyShared" -Value "No" }
        }
        $fileinfo | Add-Member -MemberType NoteProperty -Name "Permissions" -Value ($permlist -join ",")
    }
    $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemPath" -Value $file.webUrl

    #handle the output
    return $fileinfo
}

function getPermissions {

    Param(
    #Use the UserId parameter to provide an unique identifier for the user object.
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$UserId,
    #Use the ItemId parameter to provide an unique identifier for the item object.
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$ItemId)

    #fetch permissions for the given item
    $uri = "https://graph.microsoft.com/beta/users/$($UserId)/drive/items/$($ItemId)/permissions"
    $permissions = (Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference).Value

    #build the permissions string
    $permlist = @()
    foreach ($entry in $permissions) {
        #Sharing link
        if ($entry.link) {
            $strPermissions = $($entry.link.type) + ":" + $($entry.link.scope)
            if ($entry.grantedToIdentitiesV2) { $strPermissions = $strPermissions + " (" + (((&{If($entry.grantedToIdentitiesV2.siteUser.email) {$entry.grantedToIdentitiesV2.siteUser.email} else {$entry.grantedToIdentitiesV2.User.email}}) | select -Unique) -join ",") + ")" }
            if ($entry.hasPassword) { $strPermissions = $strPermissions + "[PasswordProtected]" }
            if ($entry.link.preventsDownload) { $strPermissions = $strPermissions + "[BlockDownloads]" }
            if ($entry.expirationDateTime) { $strPermissions = $strPermissions + " (Expires on: $($entry.expirationDateTime))" }
            $permlist += $strPermissions
        }
        #Invitation
        elseif ($entry.invitation) { $permlist += $($entry.roles) + ":" + $($entry.invitation.email) }
        #Direct permissions
        elseif ($entry.roles) {
            if ($entry.grantedToV2.siteUser.Email) { $roleentry = $entry.grantedToV2.siteUser.Email }
            elseif ($entry.grantedToV2.User.Email) { $roleentry = $entry.grantedToV2.User.Email }
            #else { $roleentry = $entry.grantedToV2.siteUser.DisplayName }
            else { $roleentry = $entry.grantedToV2.siteUser.loginName } #user claim
            $permlist += $($entry.Roles) + ':' + $roleentry #apparently the email property can be empty...
        }
        #Inherited permissions
        elseif ($entry.inheritedFrom) { $permlist += "[Inherited from: $($entry.inheritedFrom.path)]" } #Should have a Roles facet, thus covered above
        #some other permissions?
        else { Write-Verbose "Permission $entry not covered by the script!"; $permlist += $entry }
    }

    #handle the output
    return $permlist
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
        Set-Variable -Name authenticationResult -Scope Global -Value (Invoke-WebRequest -Method Post -Uri $url -Debug -Verbose -Body $body -ErrorAction Stop)
        $token = ($authenticationResult.Content | ConvertFrom-Json).access_token
    }
    catch { $_; return }

    if (!$token) { Write-Host "Failed to aquire token!"; return }
    else {
        Write-Verbose "Successfully acquired Access Token"

        #Use the access token to set the authentication header
        Set-Variable -Name authHeader -Scope Global -Value @{'Authorization'="Bearer $token";'Content-Type'='application\json'}
    }
}

function Invoke-GraphApiRequest {
    param(
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$Uri
    )

    if (!$AuthHeader) { Write-Verbose "No access token found, aborting..."; throw }

    try { $result = Invoke-WebRequest -Headers $AuthHeader -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop }
    catch [System.Net.WebException] {
        if ($_.Exception.Response -eq $null) { throw }

        #Get the full error response
        $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
        $streamReader.BaseStream.Position = 0
        $errResp = $streamReader.ReadToEnd() | ConvertFrom-Json
        $streamReader.Close()

        if ($errResp.error.code -match "ResourceNotFound|Request_ResourceNotFound") { Write-Verbose "Resource $uri not found, skipping..."; return } #404, continue
        #also handle 429, throttled (Too many requests)
        elseif ($errResp.error.code -eq "BadRequest") { return } #400, we should terminate... but stupid Graph sometimes returns 400 instead of 404
        elseif ($errResp.error.code -eq "Forbidden") { Write-Verbose "Insufficient permissions to run the Graph API call, aborting..."; throw } #403, terminate
        elseif ($errResp.error.code -eq "InvalidAuthenticationToken") {
            if ($errResp.error.message -eq "Access token has expired.") { #renew token, continue
                Write-Verbose "Access token has expired, trying to renew..."
                Renew-Token

                if (!$AuthHeader) { Write-Verbose "Failed to renew token, aborting..."; throw }
                #Token is renewed, retry the query
                $result = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference
            }
            else { Write-Verbose "Access token is invalid, exiting the script." ; throw } #terminate
        }
        else { $errResp ; throw }
    }
    catch { $_ ; return }

    if ($result) {
        if ($result.Content) { ($result.Content | ConvertFrom-Json) }
        else { return $result }
    }
    else { return }
}

#==========================================================================
#Main script starts here
#==========================================================================

#Variables to configure
$tenantID = "tenant.onmicrosoft.com" #your tenantID or tenant root domain
$appID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #the GUID of your app. For best result, use app with Sites.ReadWrite.All scope granted.
$client_secret = "verylongsecurestring" #client secret for the app

Renew-Token

#Used to determine whether sharing is done externally, needs Directory.Read.All scope.
$domains = (Invoke-GraphApiRequest -uri "https://graph.microsoft.com/v1.0/domains" -Verbose:$VerbosePreference).Value | ? {$_.IsVerified -eq "True"} | select -ExpandProperty Id
#$domains = @("xxx.com","yyy.com")

#Adjust the input parameters
if ($ExpandFolders -and ($depth -le 0)) { $depth = 0 }

#Get a list of all users, make sure to handle multiple pages
$GraphUsers = @()
$uri = "https://graph.microsoft.com/v1.0/users/?`$select=displayName,mail,userPrincipalName,id,userType&`$top=999&`$filter=userType eq 'Member'"
do {
    $result = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop
    $uri = $result.'@odata.nextLink'
    #If we are getting multiple pages, best add some delay to avoid throttling
    Start-Sleep -Milliseconds 500
    $GraphUsers += $result.Value
} while ($uri)
if (!$GraphUsers) { throw "No users found, aborting..." }

#Get the drive for each user and enumerate files
$Output = @()
$count = 1; $PercentComplete = 0;
foreach ($user in $GraphUsers) {
    #Progress message
    $ActivityMessage = "Retrieving data for user $($user.displayName). Please wait..."
    $StatusMessage = ("Processing user {0} of {1}: {2}" -f $count, @($GraphUsers).count, $user.userPrincipalName)
    $PercentComplete = ($count / @($GraphUsers).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #simple anti-throttling control
    Start-Sleep -Milliseconds 500
    Write-Verbose "Processing user $($user.userPrincipalName)..."

    #Check whether the user has ODFB drive provisioned
    $uri = "https://graph.microsoft.com/v1.0/users/$($user.id)/drive/root"
    $UserDrive = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop

    #If no items in the drive, skip
    if (!$UserDrive -or ($UserDrive.folder.childCount -eq 0)) { Write-Verbose "No items to report on for user $($user.userPrincipalName), skipping..."; continue }

    #enumerate items in the drive and prepare the output
    $pOutput = processChildren -User $user -URI $uri -ExpandFolders:$ExpandFolders -depth $depth
    $Output += $pOutput
}

#Return the output
#$Output | select OneDriveOwner,Name,ItemType,Shared,ExternallyShared,Permissions,ItemPath | ? {$_.Shared -eq "Yes"} | Ogv -PassThru
$global:varODFBSharedItems = $Output | select OneDriveOwner,Name,ItemType,Shared,ExternallyShared,Permissions,ItemPath | ? {$_.Shared -eq "Yes"}
#$Output | select OneDriveOwner,Name,ItemType,Shared,ExternallyShared,Permissions,ItemPath | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_ODFBSharedItems.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
return $global:varODFBSharedItems
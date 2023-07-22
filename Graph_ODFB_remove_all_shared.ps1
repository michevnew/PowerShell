#Requires -Version 3.0
# Make sure to fill in all the required variables before running the script
# Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    User.Read.All to enumerate all users in the tenant
#    Sites.ReadWrite.All to return all the item sharing details
# Help file: https://github.com/michevnew/PowerShell/blob/master/Graph_ODFB_remove_all_shared.md
# More info at: https://www.michev.info/blog/post/3018/remove-sharing-permissions-on-all-files-in-users-onedrive-for-business

[CmdletBinding()] #Make sure we can use -Verbose
Param([switch]$ExpandFolders=$true,[int]$Depth=2,[string]$User)

#==========================================================================
# Helper functions
#==========================================================================
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
        $result = (Invoke-WebRequest -Uri "$URI" -Verbose:$VerbosePreference -Headers $authHeader -ErrorAction Stop).Content | ConvertFrom-Json
        $URI = $result.'@odata.nextLink'
        #If we are getting multiple pages, add some delay to avoid throttling
        Start-Sleep -Milliseconds 500
        $children += $result
    } while ($URI)
    if (!$children) { Write-Verbose "No child items found..."; continue }

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
        if ($file.shared) {
            Write-Verbose "Found shared file ($($file.name)), removing permissions..."
            RemovePermissions $User.id $file.id -Verbose:$VerbosePreference
            $fileinfo = New-Object psobject
            $fileinfo | Add-Member -MemberType NoteProperty -Name "Name" -Value $file.name
            $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemType" -Value "File"
            $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemPath" -Value $file.webUrl
            $output += $fileinfo
        }
        else { continue }
    }
    
    #Process Notebooks
    foreach ($notebook in $cNotebooks) {
        if ($notebook.shared) {
            Write-Verbose "Found shared notebook ($($notebook.name)), removing permissions..."
            RemovePermissions $User.id $notebook.id -Verbose:$VerbosePreference
            $fileinfo = New-Object psobject
            $fileinfo | Add-Member -MemberType NoteProperty -Name "Name" -Value $notebook.name
            $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemType" -Value "Notebook"
            $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemPath" -Value $notebook.webUrl
            $output += $fileinfo
        }
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

    #if the Shared property is set, fetch permissions
    if ($folder.shared) {
        Write-Verbose "Found shared folder ($($folder.name)), removing permissions..."
        RemovePermissions $User.id $folder.id -Verbose:$VerbosePreference
        $fileinfo = New-Object psobject
        $fileinfo | Add-Member -MemberType NoteProperty -Name "Name" -Value $folder.name
        $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemType" -Value "Folder"
        $fileinfo | Add-Member -MemberType NoteProperty -Name "ItemPath" -Value $folder.webUrl
    }

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

function RemovePermissions {
    
    Param(
    #Use the UserId parameter to provide an unique identifier for the user object.
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$UserId,
    #Use the ItemId parameter to provide an unique identifier for the item object.
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$ItemId)

    #fetch permissions for the given item
    $permissions = @()
    $uri = "https://graph.microsoft.com/beta/users/$($UserId)/drive/items/$($ItemId)/permissions"
    $result = Invoke-WebRequest -Uri $uri -Verbose:$VerbosePreference -Headers $authHeader -ErrorAction Stop
    if ($result) { $permissions = ($result.content | ConvertFrom-Json).Value }
    else { continue }
  
    foreach ($entry in $permissions) {
        if ($entry.inheritedFrom) { Write-Verbose "Skipping inherited permissions..." ; continue }
        Invoke-WebRequest -Method DELETE -Verbose:$VerbosePreference -Uri "$uri/$($entry.id)" -Headers $authHeader -ErrorAction Stop | Out-Null
    }
    #check for sp. prefix on permission entries 
    #SC admin permissions are skipped, not covered via the "shared" property
}

#==========================================================================
# Main script starts here
#==========================================================================

#Variables to configure
$tenantID = "tenant.onmicrosoft.com" #your tenantID or tenant root domain
$appID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #the GUID of your app. For best result, use app with Sites.ReadWrite.All scope granted.
$client_secret = "verylongsecurestring" #client secret for the app

$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $client_secret
    grant_type    = "client_credentials"
}

#Simple code to get an access token, add your own handlers as needed
Write-Verbose "Acquiring token..."
try { $global:authenticationResult = Invoke-WebRequest -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body $body -ErrorAction Stop -Verbose:$VerbosePreference }
catch { Write-Host "Unable to obtain access token, aborting..."; return }

$token = ($authenticationResult.Content | ConvertFrom-Json).access_token

#prepare auth header
$global:authHeader = @{
   'Content-Type'='application\json'
   'Authorization'="Bearer $token"
}
Write-Verbose "Successfully acquired Access Token..."

if (!$user) { Write-Error "No user specified, aborting..." -ErrorAction Stop }
#Check the user object
Write-Verbose "Checking user $user ..."
$uri = "https://graph.microsoft.com/v1.0/users/$user"
try { $result = Invoke-WebRequest -Uri $uri -Verbose:$VerbosePreference -Headers $authHeader -ErrorAction Stop }
catch { Write-Error 'No matching user found, check the value of the $user parameter' -ErrorAction Stop }
$GraphUser = $result.Content | ConvertFrom-Json


$Output = @()
Write-Verbose "Processing user $($GraphUser.userPrincipalName) ODFB drive..."
#Check whether the user has ODFB drive provisioned
$uri = "https://graph.microsoft.com/v1.0/users/$($GraphUser.id)/drive/root"
try { $result = Invoke-WebRequest -Uri $uri -Verbose:$VerbosePreference -Headers $authHeader -ErrorAction Stop }
catch { Write-Error "User $user doenst have OneDrive provisioned, aborting..." -ErrorAction Stop }
$UserDrive = $result.Content | ConvertFrom-Json

#If no items in the drive, skip
if (!$UserDrive -or ($UserDrive.folder.childCount -eq 0)) { Write-Verbose "No items found for user $$user" }
else {
    #enumerate items in the drive and prepare the output
    Write-Verbose "Processing drive items..."
    $pOutput = (processChildren -User $GraphUser -URI $uri -ExpandFolders:$ExpandFolders -depth $depth)
    $Output += $pOutput
}

#Return the output
$global:varODFBSharedItems = $Output | select Name,ItemType,ItemPath
#$Output | select Name,ItemType,ItemPath | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_ODFBSharedItems.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
if ($varODFBSharedItems) {
    Write-Output "The following shared items were found and permissions removed where possible:" 
    return $global:varODFBSharedItems
}
else { Write-Output "No shared items found for $user" }
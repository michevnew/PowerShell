#Requires -Version 3.0
#Make sure to fill in all the required variables before running the script
#Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    User.ReadWrite.All

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/3490/remove-all-office-365-licenses-for-a-group-of-users-from-csv-file-via-graph

param([string[]]$UserList)
[CmdletBinding()] #Make sure we can use -Verbose

#region Authentication
#We use the client credentials flow as an example. For production use, REPLACE the code below wiht your preferred auth method. NEVER STORE CREDENTIALS IN PLAIN TEXT!!!

#Variables to configure
$tenantID = "tenant.onmicrosoft.com" #your tenantID or tenant root domain
$appID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #the GUID of your app. For best result, use app with Directory.Read.All scope granted.
$client_secret = "verylongsecurestring" #client secret for the app

#Prepare token request
$url = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'

$body = @{
    grant_type = "client_credentials"
    client_id = $appID
    client_secret = $client_secret
    scope = "https://graph.microsoft.com/.default"
}

#Obtain the token
Write-Verbose "Authenticating..."
try {
    $tokenRequest = Invoke-WebRequest -Method Post -Uri $url -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing -ErrorAction Stop
    $token = ($tokenRequest.Content | ConvertFrom-Json).access_token

    $authHeader = @{
       'Content-Type'='application\json'
       'Authorization'="Bearer $token"
    }
}
catch { Write-Error "Unable to obtain access token, aborting..." -ErrorAction Stop; return }
#endregion Authentication

#If a list of users was provided via the -UserList parameter
if (!$UserList) {
    Write-Verbose "No user list provided, please enter (a list of) UPN(s):"
    $UserList = Read-Host "Enter UPN"
}

$UserList = $UserList.Split(",") | ? {$_ -match ".+@.+\..+|.{8}-.{4}-.{4}-.{4}-.{12}"} #lazy regex to match either UPNs or GUIDs

if (!$UserList) { Write-Verbose "No valid UPN provided, aborting..."; return }

Write-Verbose "The following list of users will be used: $($UserList -join ",")"

#Loop over each entry
$count = 1; $PercentComplete = 0;
foreach ($u in $UserList) {
    #Simple progress indicator
    $ActivityMessage = "Retrieving data for user $($u). Please wait..."
    $StatusMessage = ("Processing user {0} of {1}: {2}" -f $count, @($UserList).count, $u)
    $PercentComplete = ($count / @($UserList).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    Write-Verbose "Processing user $($u)"
    try {
        $uri = "https://graph.microsoft.com/v1.0/users/$($u)?`$select=id,userPrincipalName,assignedLicenses,licenseAssignmentStates"
        $res = Invoke-WebRequest -Method Get -Headers $authHeader -Uri $uri -UseBasicParsing -ErrorAction Stop -Verbose:$VerbosePreference
        $user = ($res.Content | ConvertFrom-Json)
    }
    catch {
        Write-Verbose "No match found for provided user entry $u, skipping..."; continue
    }

    #Check if the user has any licenses applied, skip to the next user if not
    if (!$user.assignedLicenses) { Write-Verbose "No Licenses found for user $($user.UserPrincipalName), skipping..." ; continue }

    #Process licenses
    $SKUs = @()
    foreach ($SKU in $user.licenseAssignmentStates) {
        if (!$SKU.assignedByGroup) { $SKUs += $SKU.SkuId }
        else { Write-Verbose "License $($SKU.skuId) is assigned via the group-based licensing feature and cannot be removed. Either remove the user from the group or unassign the group license, as needed."; continue }
    }
    if (!$SKUs) { Write-Verbose "No direct-assigned licenses found for user $($user.UserPrincipalName), skipping..." ; continue }

    #Remove licenses
    Write-Verbose "Removing license(s) $($SKUs -join ",") from user $($user.UserPrincipalName)"
    #prepare the request
    $uri = "https://graph.microsoft.com/v1.0/users/$($user.UserPrincipalName)/assignLicense"
    $body = @{
        "addLicenses" = @()
        "removeLicenses" = $SKUs
    }

    #try to remove the licenses
    try {
        Invoke-WebRequest -Headers $authHeader -Uri $uri -Body ($body | ConvertTo-Json) -Method Post -ErrorAction Stop -Verbose -ContentType 'application/json' -UseBasicParsing | Out-Null
    }
    catch {
        $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
        $errResp = $streamReader.ReadToEnd() | ConvertFrom-Json
        $streamReader.Close()
        if ($errResp.error.message -match "User license is inherited from a group membership and it cannot be removed directly from the user.") {
            Write-Verbose "At least one license is assigned via the group-based licensing feature, either remove the user from the group or unassign the group license, as needed."
            continue
        }
        elseif ($errResp.error.message -match "License assignment failed because service plan .{8}-.{4}-.{4}-.{4}-.{12} depends on the service") {
            Write-Verbose "At least one license has a dependency on another license and cannot be removed (for example, because it's assigned by a group)."
            continue
        }
        else {$errResp | fl * -Force; continue} #catch-all for any unhandled errors
    }

    Write-Verbose "Removed direct-assigned licenses from user $($user.UserPrincipalName)"
}
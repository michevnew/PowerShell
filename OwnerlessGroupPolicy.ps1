######################################################################################################## 
# For details on what the script does and how to run it, check: https://www.michev.info/Blog/Post/4148 #
########################################################################################################

#Simple function to get an access token via the MSAL library. Replace with your preferred method!
function Get-MSALTokenForDefaultApp {
    param(
    #Tenant identifier
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][string]$TenantId
    )

    #try loading the MSAL binaries
    try { Add-Type -Path "C:\Program Files\WindowsPowerShell\Modules\MSAL\Microsoft.Identity.Client.dll" -ErrorAction Stop }
    catch { Write-Error "Unable to load the MSAL library, aborting..." -ErrorAction Stop; return }

    #build an app
    try {
        #We can use tenant.onmicrosoft.com values here too
        ($null -ne $TenantId -and (($TenantId -match ".+\.onmicrosoft\.com") -or ([System.Guid]::Parse($TenantId).Guid))) | Out-Null
        Write-Verbose "Creating MSAL application with Tenant ID value: $TenantId"
        $app = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create("fb78d390-0c51-40cd-8e17-fdbfab77341b").WithRedirectUri("https://login.microsoftonline.com/common/oauth2/nativeclient").WithTenantId($tenantId).WithBroker().Build()
        }
    catch {
        Write-Verbose "No valid value of `$TenantId provided, ignoring the parameter..."
        $app = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create("fb78d390-0c51-40cd-8e17-fdbfab77341b").WithRedirectUri("https://login.microsoftonline.com/common/oauth2/nativeclient").WithBroker().Build() 
        }

    #Use default value for scopes
    $Scopes = New-Object System.Collections.Generic.List[string]
    $Scope = "https://outlook.office365.com/.default"
    $Scopes.Add($Scope)

    #try fetching an access token
    try { $token = $app.AcquireTokenInteractive($Scopes).ExecuteAsync().Result }
    catch { Write-Error $_ -ErrorAction Stop }

    if ($token) { return $token }
    else { Write-Error "No access token acquired, exiting..." -ErrorAction Stop; return }
}

#Function to query the /Policy/OwnerlessGroupPolicy endpoint. If no authentication data is passed, it will try to obtain a token via the Get-MSALTokenForDefaultApp function.
function Get-OwnerlessGroupPolicy {
    param(
    #Access token
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][string]$AccessToken,
    #Tenant identifier
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][string]$TenantId
    )

    #Maybe add a proper validation...
    if (!$AccessToken) {
        Write-Verbose "No access token provided, trying to obtain one via the Get-MSALTokenForDefaultApp function..."
        if ($TenantId) { $Token = Get-MSALTokenForDefaultApp $TenantId }
        else  { $Token = Get-MSALTokenForDefaultApp }

        if ($Token.AccessToken) {
            Write-Verbose "Succesfully obtained an Access token"
            $AccessToken = $Token.AccessToken
            }
        else { Write-Error "Unable to obtain an Access token, aborting..." -ErrorAction Stop; return }
        if ($Token.TenantId) { $TenantId = $Token.TenantId }
    }

    $authHeader = @{
        'Authorization'="Bearer $($AccessToken)"
    }
  
    #Validate the Tenant ID
    try {
        [System.Guid]::Parse($TenantId) | Out-Null
        if ([System.Guid]::Parse($TenantId) -eq [System.Guid]::Empty) { throw "Not a valid GUID!" }
        }
    catch { Write-Error "The provided value for the -TenantId parameter is not a valid GUID!" -ErrorAction Stop; return }

    #Get Ownerless group policy data
    $uri = "https://outlook.office.com/ows/groupsapi/v0.1/organizations('TID:$($TenantId)')/Policy/OwnerlessGroupPolicy"
    try { $res = Invoke-WebRequest -Uri $uri -Headers $authHeader -Verbose -Debug }
    catch { Write-Error $_ -ErrorAction Stop; return }
    return ($res.Content | ConvertFrom-Json)
}

function Set-OwnerlessGroupPolicy {
    param(
    #Access token
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][string]$AccessToken,
    #Tenant identifier
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][string]$TenantId,
    #Policy status
    [Parameter(Mandatory=$false)][bool]$Enabled=$true,
    #Email address to send messages from
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][string]$SenderEmailAddress,
    #Send notifications for how many weeks?
    [Parameter(Mandatory=$false)][ValidateRange(1,7)][int]$NoOfWeeksToNotify=4,
    #Notify how many of the active members?
    [Parameter(Mandatory=$false)][ValidateRange(1,90)][int]$MaxNoOfMembersToNotify=5,
    #List of groups to cover. Accepts GUIDs, SMTP addresses or any other identifier recognizable by Get-Recipient. To reset the value to "All groups", pass an empty array @()
    [Parameter(Mandatory=$false)][System.Collections.Generic.List[string]]$EnabledGroupIds,
    #Whether to allow or exclude the groups specified. Must be used together with the SecurityGroups parameter
    [Parameter(Mandatory=$false)][bool]$IsRuleAllowType,
    #List of group whose members can be nominated as owners. SINGLE group only. Must be security-enabled. Must be used together with IsRuleAllowType
    #Even though the API definition seems to imply an array value, you can only designate a SINGLE group. To reset the value to "All group members", pass an empty array @()
    [Parameter(Mandatory=$false)][ValidatePattern('(?im)^[0-9A-F]{8}-?(?:[0-9A-F]{4}-){3}[0-9A-F]{12}$')][System.Collections.Generic.List[string]]$SecurityGroups,
    #Subject of the notification email
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][string]$EmailSubject,
    #Body of the notification email. Pass a here string for multi-line body, or add `n as needed.
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][string]$EmailBody,
    #Link to the Policy Guideliness URL
    [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][ValidatePattern('^(https:\/\/)[\w.-]+(.[\w.-]+)+[\w\-_~:/?#[\]@!\$&\(\)\*\+,;=.]+$')][string]$PolicyUrl
    )

    #region ParameterValidation
    #Auth details. Maybe add a proper validation for the Access token (aud, scopes, etc)
    if (!$AccessToken) {
        Write-Verbose "No access token provided, trying to obtain one via the Get-MSALTokenForDefaultApp function..."
        if ($TenantId) { $Token = Get-MSALTokenForDefaultApp $TenantId }
        else  { $Token = Get-MSALTokenForDefaultApp }

        if ($Token.AccessToken) {
            Write-Verbose "Succesfully obtained an Access token"
            $AccessToken = $Token.AccessToken
            }
        else { Write-Error "Unable to obtain an Access token, aborting..." -ErrorAction Stop; return }
        if ($Token.TenantId) { $TenantId = $Token.TenantId }
        if ($Token.Account) { $AccountId = $token.Account.Username }
    }

    $authHeader = @{
        'Authorization'="Bearer $($AccessToken)"
    }
  
    #Validate the Tenant ID. As we DO NOT parse the JWT, you must ensure the tenantID value matches the corresponding claim in the access token!
    try {
        [System.Guid]::Parse($TenantId) | Out-Null
        if ([System.Guid]::Parse($TenantId) -eq [System.Guid]::Empty) { throw "Not a valid GUID!" }
        }
    catch { Write-Error "Plese provide a valid GUID value for the `$TenantId parameter!" -ErrorAction Stop; return }

    #Enabled needs at least NoOfWeeksToNotify,MaxNoOfMembersToNotify,SenderEmailAddress!
    if ($true -eq $Enabled) {
        Write-Verbose "Enabling Ownerless group policy"
        $parametersJson = [ordered]@{ "enabled" = $true }

        #SenderEmailAddress shennanigans. We can actually validate the value provided via a query to the EXO REST endpoint, with the same token!
        if ($PSBoundParameters.ContainsKey('SenderEmailAddress')) {
            try { 
            $sender = Invoke-RestMethod -Method Get -Uri "https://outlook.office.com/adminApi/beta/$($TenantId)/Recipient('$SenderEmailAddress')" -Headers $AuthHeader -Verbose -Debug -ErrorAction Stop
            $SenderEmailAddress = $sender.PrimarySmtpAddress.ToString()
            }
            catch { Write-Error "Unable to find a matching recpient for the provided value of the -SenderEmailAddress parameter, aborting..." -ErrorAction Stop; return }
        }
        else {
            Write-Verbose "No value provided for the -SenderEmailAddress parameter, try to leverage the current user instead..."
            if ($AccountId) { $SenderEmailAddress = $AccountId }
            else { Write-Error "The -SenderEmailAddress parameter is mandatory when enabling the policy, aborting..." -ErrorAction Stop; return }
        }
        $parametersJson["senderEmailAddress"] = $SenderEmailAddress
        $parametersJson["noOfWeeksToNotify"] = $NoOfWeeksToNotify
        $parametersJson["maxNoOfMembersToNotify"] = $MaxNoOfMembersToNotify

        #The rest are all optional parameters
        #try validating each group ID out of $EnabledGroupIds
        if ($PSBoundParameters.ContainsKey('EnabledGroupIds')) {
            $EnabledGroupsList = @()
            foreach ($GroupId in $EnabledGroupIds) {
                try {
                    $Group = Invoke-RestMethod -Method Get -Uri "https://outlook.office.com/adminApi/beta/$($TenantId)/Recipient('$GroupId')" -Headers $AuthHeader -Verbose -Debug -ErrorAction Stop
                    Write-Verbose "Found a match for group id $($GroupId)"
                    $EnabledGroupsList += $Group.ExternalDirectoryObjectId
                }
                catch { Write-Verbose "Unable to find a match for group id $($GroupId), removing from list..." }
            }
            #We don't want to override the current list if no valid values are provided
            if ($PSBoundParameters["EnabledGroupIds"].Count -eq 0) {
                Write-Verbose "Null value provided, resetting the list of groups to cover by the policy."
                $parametersJson["enabledGroupIds"] = @()
            }
            else {
                if ($EnabledGroupsList) {
                    Write-Verbose "The following list of groups will be covered by the policy: $($EnabledGroupsList -join ",")"
                    $parametersJson["enabledGroupIds"] = $EnabledGroupsList 
                }
                else { Write-Verbose "No matching Groups found for the value provided for the -EnabledGroupIds parameter, skipping..." }
            }
        }

        #IsRuleAllowType should only be called together with SecurityGroups
        if ($PSBoundParameters.ContainsKey('SecurityGroups')) {
            if (!$SecurityGroups) { Write-Verbose "Provided an empty/null value for the -SecurityGroups parameter, resetting the list of allowed owners to All members..." }
            else {
                if (!$PSBoundParameters.ContainsKey('IsRuleAllowType')) {
                    Write-Error "Non-null value provided for the -SecurityGroups parameter, a value for the -IsRuleAllowType parameter must also be specified!" -ErrorAction Stop; return
            }}
        }
        else {
            if ($PSBoundParameters.ContainsKey('IsRuleAllowType')) {
                Write-Error "A value for the -SecurityGroups parameter must also be specified!" -ErrorAction Stop; return
        }}

        #$SecurityGroups cannot be validated, as we can only fetch mail-enabled SGs
        if ($PSBoundParameters.ContainsKey('isRuleAllowType')) { $parametersJson["isRuleAllowType"] = $IsRuleAllowType }
        if ($PSBoundParameters.ContainsKey('SecurityGroups')) {
            if ($SecurityGroups) { $parametersJson["securityGroups"] = @($SecurityGroups[0]) }
            else { $parametersJson["securityGroups"] = @() }
            } #take only the first one, NOT a multi-value property
        if ($EmailSubject) { $parametersJson["emailSubject"] = $EmailSubject }
        if ($EmailBody) { $parametersJson["emailBody"] = $EmailBody }
        if ($PolicyUrl) { $parametersJson["policyUrl"] = $PolicyUrl }

    }
    else {
        Write-Verbose "Disabling Ownerless group policy"
        $parametersJson = [ordered]@{ "enabled" = $false }
    }
    #Enabled=false doesn't need any other parameters, and clears their values. Even if you specify a value, it will be cleared in any subsequent GET requests, regardless of what the output of the POST request shows!
    #endregion

    #verify the mandatory parameters are present
    if ($parametersJson["enabled"] -and (!$parametersJson.Contains("maxNoOfMembersToNotify") -or !$parametersJson.Contains("noOfWeeksToNotify") -or !$parametersJson.Contains("senderEmailAddress"))) {
        Write-Error "Insufficient data. Please provide valid values for all the following parameteres: MaxNoOfMembersToNotify, NoOfWeeksToNotify, SenderEmailAddress" -ErrorAction Stop
    }
    Write-Verbose "The following policy settings will be used:`n $($parametersJson | Out-String)"

    #Set Ownerless group policy data
    $uri = "https://outlook.office.com/ows/groupsapi/v0.1/organizations('TID:$($TenantId)')/Policy/OwnerlessGroupPolicy"
    try { $res = Invoke-WebRequest -Uri $uri -Headers $authHeader -Verbose -Debug -Method POST -ContentType 'application/json' -Body ($parametersJson | ConvertTo-Json) }
    catch { Write-Error $_ -ErrorAction Stop; return }
    return ($res.Content | ConvertFrom-Json)

}
#Load the MSAL binaries
Add-Type -LiteralPath "C:\Program Files\PackageManagement\NuGet\Packages\Microsoft.IdentityModel.Abstractions.6.22.0\lib\net45\Microsoft.IdentityModel.Abstractions.dll"
Add-Type -LiteralPath "C:\Program Files\PackageManagement\NuGet\Packages\Microsoft.Identity.Client.4.54.1\lib\net45\Microsoft.Identity.Client.dll"

#region "Helper functions"
function Get-AccessTokens {
    $global:hashtokens = @{}
    $app =  [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create("1950a258-227b-4e31-a9cf-717495945fc2").WithRedirectUri("https://login.microsoftonline.com/common/oauth2/nativeclient").Build()

    $scopes = @("https://graph.microsoft.com/.default","https://service.flow.microsoft.com/.default","https://service.powerapps.com/.default")
    foreach ($scope in $scopes) {
        $TokenScopes = New-Object System.Collections.Generic.List[string]
        $TokenScopes.Add($Scope)

        try {
            if ($hashTokens.Count) { $token = $app.AcquireTokenSilent($TokenScopes,$app.GetAccountsAsync().Result.Username).ExecuteAsync().Result }
            else { $token = $app.AcquireTokenInteractive($TokenScopes).ExecuteAsync().Result }
        }
        catch { $_; return }

        if (!$token) { Write-Host "Failed to aquire token!"; return }
        else {
            Write-Verbose "Successfully acquired Access Token with scope $scope"
            $hashTokens[$scope] = $token.AccessToken
        }
    }
}

function Invoke-GraphApiRequest {
    param(
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$Uri
    )

    if (!$hashTokens) { Write-Verbose "No access token found, aborting..."; throw }

    $authHeader = $null
    switch -Wildcard ($Uri) {
        "https://graph.microsoft.com/*" { $authHeader = @{'Authorization'="Bearer $($hashTokens['https://graph.microsoft.com/.default'])";'Content-Type'='application\json'} }
        "https://api.bap.microsoft.com/*" { $authHeader = @{'Authorization'="Bearer $($hashTokens['https://service.powerapps.com/.default'])";'Content-Type'='application\json'} }
        "https://api.flow.microsoft.com/*" { $authHeader = @{'Authorization'="Bearer $($hashTokens['https://service.flow.microsoft.com/.default'])";'Content-Type'='application\json'} }
    }

    Write-Verbose "Processing request $Uri"
    try { $result = Invoke-WebRequest -Headers $authHeader -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop }
    catch [System.Net.WebException] {
        if ($_.Exception.Response -eq $null) { throw }

        #Get the full error response
        $streamReader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
        $streamReader.BaseStream.Position = 0
        $global:errResp = $streamReader.ReadToEnd() | ConvertFrom-Json
        $streamReader.Close()

        if ($errResp.error.code -match "ResourceNotFound|Request_ResourceNotFound|FlowNotFound") { Write-Verbose "Resource $uri not found, skipping..."; return } #404, continue
        #also handle 429, throttled (Too many requests)
        elseif ($errResp.error.code -eq "BadRequest") { return } #400, we should terminate...
        elseif ($errResp.error.code -match "Forbidden|InvalidPath|AuthenticationFailed") { Write-Verbose "Insufficient permissions to run the Graph API call, aborting..."; throw } #403, terminate
        elseif ($errResp.error.code -match "InvalidAuthenticationToken|ExpiredAuthenticationToken") {
            if ($errResp.error.message -match "Access token has expired|The access token expiry|The received access token has expiry") { #renew token, continue
                Write-Verbose "Access token has expired, trying to renew..."
                Get-AccessTokens

                if (!$hashTokens) { Write-Verbose "Failed to renew token, aborting..."; throw }
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

function StateToStatus ([string]$state) {
    switch ($state) {
        "Suspended" { "Suspended" }
        "Started" { "Enabled" }
        "Stopped" { "Disabled" }
        Default { "Unknown" }
    }
}

function TypeToType ([string]$type) {
    switch ($type) {
        "Request" { "Instant" }
        "ApiConnection" { "Automated" }
        "ApiConnectionWebhook" { "Automated" }
        "ApiConnectionNotification" { "Automated" }
        "OpenApiConnection" { "Automated" }
        "Recurrence" { "Scheduled" }
        Default { "Unknown" }
    }
}

function processConnections ($connections) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($connection in $connections) {
        #Replace "shared_" for any default connectors, leave other values?
        $out.Add("[" + $connection.apiId.Substring(36).Replace("shared_","") + "]" + $connection.displayName + (&{If($connection.statuses[0].error) {"*"}}))
        #* indicates connection with Error status (expired, etc)
    }
    return $out -join ";"
}

#Works for both Actions and Triggers
function processActions ($actions) {
    $hash = @{};$out = New-Object System.Collections.Generic.List[string]
    foreach ($action in $actions) {
        if (!$action.swaggerOperationId) { $action.swaggerOperationId = $action.type }
        if (!$action.api) { $action.api = "" }

        if ($hash.ContainsKey($action.api)) { $hash[$action.api] = $hash[$action.api] + "+" + $action.swaggerOperationId }
        else { $hash[$action.api] = $action.swaggerOperationId }
    }
    foreach ($o in $hash.GetEnumerator()) {
        $out.Add("[" + $o.name.Replace("shared_","") + "]" + $o.value)
    }
    return ($out -join ";").Replace("[]","")
}

#endregion "Helper functions"

#Start by obtaining the required access tokens
Get-AccessTokens

#Get the list of environments
$environments = (Invoke-GraphApiRequest -Uri "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments?api-version=2016-11-01").Value

#Get the list of Flows. #/scopes/admin/ DOES NOT return ALL flows as it has no support for include=includeSolutionCloudFlows
$flows = @();$hashUsers = @{};$hashTemplates = @{};
foreach ($env in $environments) {
    $uri = "https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/$($env.name)/v2/flows?api-version=2016-11-01&`$top=250"
    do {
        $result = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop
        $uri = $result.'nextLink'

        $flows += $result.value
    } while ($uri)
}

$count = 1; $PercentComplete = 0;
#Loop over each flow to gather additional details
foreach ($flow in $flows) {
    #Progress message
    $ActivityMessage = "Retrieving data for flow $($flow.Name). Please wait..."
    $StatusMessage = ("Processing flow {0} of {1}: {2}" -f $count, @($flows).count, $flow.properties.displayName)
    $PercentComplete = ($count / @($flows).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #Get flow details
    $flowID = $flow.id.Replace("/environments","/scopes/admin/environments")
    $uri = "https://api.flow.microsoft.com" + $flowID + "?api-version=2016-11-01"
    $FlowDetails = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop

    $flow | Add-Member -NotePropertyName FlowName -NotePropertyValue $FlowDetails.properties.displayName
    $flow | Add-Member -NotePropertyName Description -NotePropertyValue $FlowDetails.properties.definitionSummary.description
    $flow | Add-Member -NotePropertyName Status -NotePropertyValue (StateToStatus $FlowDetails.properties.state)
    $flow | Add-Member -NotePropertyName Created -NotePropertyValue $FlowDetails.properties.createdTime
    $flow | Add-Member -NotePropertyName Modified -NotePropertyValue $FlowDetails.properties.lastModifiedTime
    $flow | Add-Member -NotePropertyName FlowType -NotePropertyValue (TypeToType $FlowDetails.properties.definitionSummary.triggers.type)
    $flow | Add-Member -NotePropertyName Properties -NotePropertyValue $FlowDetails.properties -Force #dump the properties blob
    #$flow | Add-Member -NotePropertyName Plan -NotePropertyValue $FlowDetails.properties.plan #plan not available on the admin endpoint?

    #Get data on flow connections
    $uri = "https://api.flow.microsoft.com" + $flowID + "/connections?api-version=2016-11-01"
    $FlowConnections = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop
    $flow | Add-Member -NotePropertyName Connections -NotePropertyValue (processConnections ($FlowConnections.properties | select apiId,displayName,statuses))

    #Triggers
    $flow | Add-Member -NotePropertyName Triggers -NotePropertyValue (processActions ($FlowDetails.properties.definitionSummary.triggers | select type,swaggerOperationId,@{n="api";e={$_.api.name}} -Unique))

    #Reference resources
    $flow | Add-Member -NotePropertyName Resources -NotePropertyValue (($FlowDetails.properties.referencedResources.service | sort -Unique) -join ";")

    #Actions
    $flow | Add-Member -NotePropertyName Actions -NotePropertyValue (processActions ($FlowDetails.properties.definitionSummary.actions | select type,swaggerOperationId,@{n="api";e={$_.api.name}} -Unique))

    #Get data on flow runs
    $uri = "https://api.flow.microsoft.com" + $flowID + "/runs?api-version=2016-11-01"
    $LastSuccessRun = Invoke-GraphApiRequest -Uri ($uri + "&`$filter=Status eq 'succeeded'&`$top=1") -Verbose:$VerbosePreference -ErrorAction Stop
    $LastFailedRun = Invoke-GraphApiRequest -Uri ($uri + "&`$filter=Status eq 'failed'&`$top=1") -Verbose:$VerbosePreference -ErrorAction Stop
    $flow | Add-Member -NotePropertyName LastSuccessRun -NotePropertyValue (&{If($LastSuccessRun.value) {$LastSuccessRun.value.properties.endTime} Else {"N/A"}})
    $flow | Add-Member -NotePropertyName LastFailRun -NotePropertyValue (&{If($LastFailedRun.value) {$LastFailedRun.value.properties.endTime} Else {"N/A"}})

    #Process owner/creator data
    if (!$hashUsers[$FlowDetails.properties.creator.userId]) {
        $createdBy = Invoke-GraphApiRequest -Uri ("https://graph.microsoft.com/v1.0/users/" + $FlowDetails.properties.creator.userId) -Verbose:$VerbosePreference -ErrorAction Stop
        $hashUsers[$createdBy.id] = $createdBy.userPrincipalName
    }
    $flow | Add-Member -NotePropertyName CreatedBy -NotePropertyValue $hashUsers[$FlowDetails.properties.creator.userId]

    if ($FlowDetails.properties.sharingType) { #DelegatedAuth or CommonDataService
        $flow | Add-Member -NotePropertyName Shared -NotePropertyValue "Yes"

        #Process each owner
        $uri = "https://api.flow.microsoft.com" + $flowID + "/owners?api-version=2016-11-01"
        $owners = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop
        foreach ($owner in $owners.value) {
            if ($owner.properties.permissionType -eq "AuthorizationDelegate") {
                #Shared with SPO list or library. Should only be present for flows with Connection with type "sharepointonline". And dynamics?
                if ($owner.properties.authorizationDelegate.delegatedAuthResource.resourceCollection -match "sharepoint.com") { $uri = "https://graph.microsoft.com/v1.0/sites/" + $owner.properties.authorizationDelegate.delegatedAuthResource.resourceCollection.Replace("https://","") + "/lists/" + $owner.properties.authorizationDelegate.delegatedAuthResource.resourceId }
                elseif ($owner.properties.authorizationDelegate.delegatedAuthResource.resourceCollection -match "crm.dynamics.com") { $uri = "yourorg.crm.dynamics.com/" + $owner.properties.authorizationDelegate.delegatedAuthResource.resourceId }
                else { $uri = $owner.properties.authorizationDelegate.delegatedAuthResource.resourceCollection + "/" + $owner.properties.authorizationDelegate.delegatedAuthResource.resourceId}

                $owner | Add-Member -NotePropertyName OwnerId -NotePropertyValue $uri
                continue #avoid Else below
            }
            if (!$hashUsers[$owner.name]) {
                if ($owner.properties.principal.type -eq "User") {
                    $createdBy = Invoke-GraphApiRequest -Uri ("https://graph.microsoft.com/v1.0/users/" + $owner.name) -Verbose:$VerbosePreference -ErrorAction Stop
                    $hashUsers[$createdBy.id] = $createdBy.userPrincipalName
                }
                else {
                    $createdBy = Invoke-GraphApiRequest -Uri ("https://graph.microsoft.com/v1.0/groups/" + $owner.name) -Verbose:$VerbosePreference -ErrorAction Stop
                    $hashUsers[$createdBy.id] = $createdBy.displayName
                }
            }
            $owner | Add-Member -NotePropertyName OwnerId -NotePropertyValue $hashUsers[$owner.name]
        }
        $flow | Add-Member -NotePropertyName Owners -NotePropertyValue ($owners.Value.OwnerId -join ";")
    }
    else {
        $flow | Add-Member -NotePropertyName Shared -NotePropertyValue "No"
        $flow | Add-Member -NotePropertyName Owners -NotePropertyValue $hashUsers[$FlowDetails.properties.creator.userId]
    }

    <#
    #Run Only Users flow
    #Trigger type should be Request and kind should be Button or APIConnection? Best we can do with /scopes/admin?
    #https://powerusers.microsoft.com/t5/General-Power-Automate/All-about-quot-Manage-Run-Only-Users-quot/td-p/122821
    if ($FlowDetails.properties.definitionSummary.triggers.type -eq "Request" -and $FlowDetails.properties.definitionSummary.triggers.kind -in @("Button","ApiConnection")) {
        #this should be a flow with trigger that supports Run Only Users
        #we can only get the /users details via the non-admin endpoint...
        $uri = "https://api.flow.microsoft.com" + $flowID + "/users?api-version=2016-11-01"
        Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop
    }
    #>

    #Process template. Does NOT use the admin endpoint!
    if (!$FlowDetails.properties.templateName) { continue } #If we created the flow from scratch, no tempalate property will be present
    if (!$hashTemplates[$FlowDetails.properties.templateName]) {
        $uri = "https://api.flow.microsoft.com" + $flow.properties.environment.id + "/galleries/public/templates/" + $($FlowDetails.properties.templateName) + "?api-version=2016-11-01"
        $template = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction SilentlyContinue
        $hashTemplates[$template.name] = $template.properties.displayName
    }
    $flow | Add-Member -NotePropertyName Template -NotePropertyValue (&{If($template) {$hashTemplates[$FlowDetails.properties.templateName]} Else {"$FlowDetails.properties.templateName"}})
}

$flows | select * -ExcludeProperty properties | Export-Csv -Path "$PSScriptRoot\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_FlowReport.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
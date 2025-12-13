#Requires -Version 3.0
#The script requires the following permissions:
#    Application.Read.All (required)
#    AuditLog.Read.All (required)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/5986/build-your-own-entra-id-application-credential-activity-report

[CmdletBinding(SupportsShouldProcess)] #Make sure we can use -Verbose
Param()

#==========================================================================
#Main script starts here
#==========================================================================

#Get an Access token. Make sure to fill in all the variable values here. Or replace with your own preferred method to obtain token.
$tenantId = "tenant.onmicrosoft.com"
$uri = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'
$clientId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$client_secret = "verylongstring"

$Scopes = New-Object System.Collections.Generic.List[string]
$Scope = "https://graph.microsoft.com/.default"
$Scopes.Add($Scope)

$body = @{
    grant_type = "client_credentials"
    client_id = $clientId
    client_secret = $client_secret
    scope = $Scopes
}

try {
    Write-Verbose "Obtaining token..."
    $res = Invoke-WebRequest -Method Post -Uri $uri -Body $body -UseBasicParsing -ErrorAction Stop -Verbose:$false
    $token = ($res.Content | ConvertFrom-Json).access_token

    $authHeader = @{
       'Authorization'="Bearer $token"
    }}
catch { Write-Output "Failed to obtain token, aborting..." ; return }

#Get the list of application objects within the tenant.
$Apps = @()

Write-Verbose "Retrieving list of applications..."
$uri = "https://graph.microsoft.com/beta/applications?`$top=999"
try {
    do {
        $result = Invoke-WebRequest -Method Get -Uri $uri -Headers $authHeader -UseBasicParsing -ErrorAction Stop -Verbose:$false
        $uri = ($result.Content | ConvertFrom-Json).'@odata.nextLink'

        #If we are getting multiple pages, best add some delay to avoid throttling
        Start-Sleep -Milliseconds 100
        $Apps += ($result.Content | ConvertFrom-Json).Value
    } while ($uri)
}
catch {
    Write-Output "Failed to retrieve the list of applications, aborting..."
    Write-Error $_ -ErrorAction Stop
    return
}

#Prepare variables
$output = [System.Collections.Generic.List[Object]]::new() #output variable
$i=0; $count = 1; $PercentComplete = 0;

#Process the list of applications
foreach ($App in $Apps) {
    #Progress message
    $ActivityMessage = "Retrieving data for application $($App.DisplayName). Please wait..."
    $StatusMessage = ("Processing application {0} of {1}: {2}" -f $count, @($Apps).count, $App.AppId)
    $PercentComplete = ($count / @($Apps).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #simple anti-throttling control
    Start-Sleep -Milliseconds 200
    Write-Verbose "Processing application $($App.id)..."

    #Get the service principal sign-in logs for the application
    #Filters SUCCESS events only - automatically done when we include the servicePrincipalCredentialKeyId filter. Consider NOT doing it?
    foreach ($cred in @($App.KeyCredentials + $App.PasswordCredentials)) {
        $KeyLastLogin = $null
        try {
            $uri = "https://graph.microsoft.com/beta/auditLogs/signIns?`$filter=(signInEventTypes/any(t:t eq 'servicePrincipal')) and appId eq `'$($App.AppId)`' and servicePrincipalCredentialKeyId eq `'$($cred.KeyId)`'&`$top=1"
            $KeyLastLogin = Invoke-WebRequest -Method Get -Uri $uri -Headers $authHeader -UseBasicParsing -ErrorAction Stop -Verbose:$false
            $KeyLastLogin = ($KeyLastLogin.Content | ConvertFrom-Json).value
        }
        catch { Write-Warning "Failed to retrieve sign-in logs for application $($App.id)"; $_ }

        #Prepare the output
        $i++;$objPermissions = [PSCustomObject][ordered]@{
            "Number" = $i
            "AppId" = $app.AppId
            "AppObjectId" = $app.Id
            "AppDisplayName" = $app.DisplayName
            "KeyId" = $cred.KeyId
            "KeyDisplayName" = & { if ($cred.DisplayName) { $cred.DisplayName } else { "N/A" } } #can be null. Portal returns Description, but it's not available via Graph?
            "KeyType" = & { if ($cred.Type) { $cred.Type } else { "Client secret" } }
            "KeyUsage" = & { if ($cred.Usage) { $cred.Usage } else { "N/A" } }
            "LastUsed" = & { if ($KeyLastLogin) { Get-Date($KeyLastLogin.CreatedDateTime.DateTime) -Format g } else { "N/A" } }
            "KeyExpirationDate" = & { if ($cred.EndDateTime) { Get-Date($cred.EndDateTime) -Format g } else { "N/A" } }
            "CredentialOrigin" = "application"
            "ServicePrincipalObjectId" = & { if ($KeyLastLogin.ServicePrincipalId) { $KeyLastLogin.ServicePrincipalId } else { "N/A" } } #Can be null, can we use it to differentiate between cert and client secret?
            "ServicePrincipalDisplayName" = & { if ($KeyLastLogin.ServicePrincipalName) { $KeyLastLogin.ServicePrincipalName } else { "N/A" } }
            "ResourceId" = & { if ($KeyLastLogin.ResourceId) { $KeyLastLogin.ResourceId } else { "N/A" } }
            "ResourceDisplayName" = & { if ($KeyLastLogin.ResourceDisplayName) { $KeyLastLogin.ResourceDisplayName } else { "N/A" } }
        }

        $output.Add($objPermissions)
    }
}

#Export the result to CSV file
$output | select * -ExcludeProperty Number | Export-CSV -nti -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_GraphAppRegInventory.csv"
Write-Verbose "Output exported to $($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_GraphAppRegInventory.csv"
#Requires -Version 3.0
#Make sure to fill in all the required variables before running the script
#Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Application.Read.All to read the app registrations
#    Application.ReadWrite.All for the remediation part

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/5894/script-to-review-and-remove-service-principal-credentials

[CmdletBinding(SupportsShouldProcess)] #Make sure we can use -WhatIf and -Verbose
Param([switch]$IncludeAllApps=$false,[switch]$Remediate=$false)

#==========================================================================
#Main script starts here
#==========================================================================

#Get MSAL token. Make sure to fill in all the variable values here. Or replace with your own preferred method to obtain token.
$tenantId = "tenant.onmicrosoft.com"
$url = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'

$Scopes = New-Object System.Collections.Generic.List[string]
$Scope = "https://graph.microsoft.com/.default"
$Scopes.Add($Scope)

$body = @{
    grant_type = "client_credentials"
    client_id = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    client_secret = "verylongstring"
    scope = $Scopes
}

try {
    $res = Invoke-WebRequest -Method Post -Uri $url -Verbose:$false -Body $body
    $token = ($res.Content | ConvertFrom-Json).access_token

    $authHeader = @{
       'Authorization'="Bearer $token"
    }}
catch { Write-Error "Failed to obtain token, aborting..." ; return }

#Get the list of application objects within the tenant.
$Apps = @()

if ($IncludeAllApps) { $uri = "https://graph.microsoft.com/beta/applications?`$top=999" }
else { $uri = "https://graph.microsoft.com/beta/applications?`$top=999&`$filter=signInAudience eq 'AzureADMultipleOrgs'" }

do {
    $result = Invoke-WebRequest -Method Get -Uri $uri -Headers $authHeader -Verbose:$false
    $uri = ($result.Content | ConvertFrom-Json).'@odata.nextLink'

    #If we are getting multiple pages, best add some delay to avoid throttling
    Start-Sleep -Milliseconds 500
    $Apps += ($result.Content | ConvertFrom-Json).Value
} while ($uri)

$output = [System.Collections.Generic.List[Object]]::new() #output variable
$i=0;

if (!$Apps) { Write-Warning "No applications found, aborting..." ; return }

foreach ($App in $Apps) {

    Write-Verbose "Processing application $($App.id)..."

    if ($Remediate) {
        if (!$app.servicePrincipalLockConfiguration.isEnabled -and !$app.servicePrincipalLockConfiguration.allProperties) {
            Write-Verbose "Application $($App.id) needs remediation, processing..."
            $body = @{"servicePrincipalLockConfiguration" = @{"isEnabled" = $true; "allProperties" = $true}}
            $uri = "https://graph.microsoft.com/beta/applications/$($App.id)"
            try {
                $res = Invoke-WebRequest -Method Patch -Uri $uri -Headers $authHeader -Body ($body | ConvertTo-Json) -ContentType "application/json" -Verbose:$false -ErrorAction Stop
                Write-Verbose "Application $($App.id) remediated successfully."
                Add-Member -InputObject $App -MemberType NoteProperty -Name "Remediated" -Value $true
            }
            catch {
                Write-Verbose "Application $($App.id) remediation failed."
                $_ | Out-Default
                Add-Member -InputObject $App -MemberType NoteProperty -Name "Remediated" -Value $false
            }
        }
        else {
            Write-Verbose "Application $($App.id) does not need remediation, skipping..."
            Add-Member -InputObject $App -MemberType NoteProperty -Name "Remediated" -Value $false
        }
    }
    else { Add-Member -InputObject $App -MemberType NoteProperty -Name "Remediated" -Value $false }

    #prepare the output object
    $i++;$objPermissions = [PSCustomObject][ordered]@{
        "Number" = $i
        "Application Name" = $App.DisplayName
        "ApplicationId" = $App.AppId
        "SignInAudience" = $app.signInAudience
        "ObjectId" = $App.id
        "Created on" = (&{if ($app.createdDateTime) {(Get-Date($App.createdDateTime) -format g)} else {"N/A"}})
        "Remediation needed" = (&{if (!$app.servicePrincipalLockConfiguration.isEnabled -and !$app.servicePrincipalLockConfiguration.allProperties) {$true} else {$false}})
        "Remediated" = $App.Remediated
    }

    $output.Add($objPermissions)
}

#Export the result to CSV file
$output | select * -ExcludeProperty Number | Export-CSV -nti -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_GraphAppRegInventory.csv"
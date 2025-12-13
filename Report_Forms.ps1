#Requires -Version 7.0

#Make sure to fill in all the required variables before running the script
#Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    (application) User.Read.All (Graph API resource, required)
#    (application) Forms.Read.All (Microsoft Forms API resource, required)
#    (delegate) Forms.Read (Microsoft Forms API resource, needed for covering group forms)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/6419/primer-reporting-on-microsoft-forms-in-use-within-your-organization

[CmdletBinding()] #Make sure we can use -Verbose
Param([switch]$IncludeGroupForms)

function Renew-Token {
    param(
    [ValidateNotNullOrEmpty()][string]$Service, #The service for which to get a token
    [switch]$ROPC #Use ROPC flow
    )

    #prepare the request
    $url = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'

    #Define the scope based on the service value provided
    if (!$Service -or $Service -eq "Graph") { $Scope = "https://graph.microsoft.com/.default" }
    elseif ($Service -eq "Forms") { $Scope = "https://forms.office.com/.default" }
    else { Write-Error "Invalid service specified, aborting..." -ErrorAction Stop; return }

    $Scopes = New-Object System.Collections.Generic.List[string]
    $Scopes.Add($Scope)

    $body = @{
        grant_type = "client_credentials"
        client_id = $appID
        client_secret = $client_secret
        scope = $Scopes
    }
    if ($ROPC) {
        $body = @{
            grant_type = "password"
            client_id = $appID
            username = $AuthCred.UserName
            password = $AuthCred.GetNetworkCredential().Password
            scope = $Scopes
        }
    }

    try {
        $authenticationResult = Invoke-WebRequest -Method Post -Uri $url -Body $body -UseBasicParsing -ErrorAction Stop -Verbose:$false
        $token = ($authenticationResult.Content | ConvertFrom-Json).access_token
    }
    catch { throw $_ }

    if (!$token) { Write-Error "Failed to aquire token!" -ErrorAction Stop; return }
    else {
        Write-Verbose "Successfully acquired Access Token for $service"

        #Use the access token to set the authentication header
        if (!$Service -or $Service -eq "Graph") { Set-Variable -Name (($ROPC) ? "authHeaderGraphROPC" : "authHeaderGraph") -Scope Global -Value @{'Authorization'="Bearer $token";'Content-Type'='application/json'} -Confirm:$false -WhatIf:$false }
        elseif ($Service -eq "Forms") { Set-Variable -Name (($ROPC) ? "authHeaderFormsROPC" : "authHeaderForms") -Scope Global -Value @{'Authorization'="Bearer $token";'Content-Type'='application/json'} -Confirm:$false -WhatIf:$false }
        else { Write-Error "Invalid service specified, aborting..." -ErrorAction Stop; return }
    }
}

#==========================================================================
# Main script
#==========================================================================

#Variables to configure
$tenantID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #Use tenantID here, NOT the tenant root domain
$appID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #the GUID of your app
$client_secret = "verylongsecurestring" #client secret for the app

Renew-Token -Service "Graph"
Renew-Token -Service "Forms"

#Get the list of all users
$Users = @()
$uri = "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,userPrincipalName&`$top=999"
try {
    do {
        $result = Invoke-WebRequest -Uri $uri -Headers $authHeaderGraph -UseBasicParsing -ErrorAction Stop -Verbose:$false
        $uri = $result.'@odata.nextLink'

        $Users += ($result.Content | ConvertFrom-Json).value
    } while ($uri)
}
catch {
    Write-Error "Failed to users, aborting..." -ErrorAction Stop
    return
}

#Generate the report file
if ($null -eq (Get-Module -Name ImportExcel -ListAvailable -Verbose:$false)) {
    Write-Verbose "The ImportExcel module was not found, skipping export to Excel file..."; return
}
#ImportExcel does NOT overwrite existing files, so we need to generate a unique filename
$excel = Export-Excel -Path "$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_Forms.xlsx" -WorksheetName Overview -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -PassThru

#Get the forms for each user
$ReportForms = @()
foreach ($User in $Users) {
    #Does this even support pagination?
    $uri = "https://forms.office.com/formapi/api/$tenantId/users/$($User.Id)/light/forms?`$select=id,status,title,createdDate,modifiedDate,version,rowCount,softDeleted,type" #CASE SESNITIVE!
    try { $result = Invoke-WebRequest -Uri $uri -Headers $authHeaderForms -UseBasicParsing -ErrorAction Stop -Verbose:$false }
    catch { Write-Error "Failed to retrieve forms for user $($User.UserPrincipalName), skipping..." -ErrorAction SilentlyContinue; $_ | fl * -Force; continue }
    $result = $result.Content | ConvertFrom-Json

    if (!$result.value) { Write-Verbose "No forms found for user $($user.UserPrincipalName), skipping..." ;continue }
    $UserForms = $result.value
    $ReportForms += [PSCustomObject]@{
        Name = $User.displayName
        Type = "User"
        Identifier = $User.userPrincipalName
        FormsCount = $UserForms.count
    }

    #Fetch the details for each form. GetRespCounts() only works in delegate context, so we iterate each form
    foreach ($form in $UserForms) {
        $uri = "https://forms.office.com/formapi/api/$tenantId/users/$($User.Id)/light/forms(`'$($form.id)`')"
        $result = (Invoke-WebRequest -Uri $uri -Headers $authHeaderForms -UseBasicParsing -ErrorAction Continue -Verbose:$false).Content | ConvertFrom-Json
        $form.rowCount = $result.rowCount
    }

    $UserForms | select Id, Status, Title, CreatedDate, ModifiedDate, version, @{n="ResponseCount";e={$_.rowCount}}, @{n="IsDeleted";e={($_.softDeleted) ? "Yes" : "No"}}, type `
    ` | Export-Excel -ExcelPackage $excel -WorksheetName $($User.UserPrincipalName) -FreezeTopRow -AutoFilter -BoldTopRow -AutoSize -NoHyperLinkConversion TargetResource -PassThru > $null

    Start-Sleep -Milliseconds 100
}

#Process group forms
if ($IncludeGroupForms) {
    $AuthCred = Get-Credential -Message "Enter credentials to use with the ROPC flow" -ErrorAction Stop
    Renew-Token -Service "Forms" -ROPC

    #Get the list of groups for the current user
    $uri = "https://forms.office.com/formapi/api/groups"
    $result = Invoke-WebRequest -Uri $uri -Headers $authHeaderFormsROPC -UseBasicParsing -ErrorAction Stop -Verbose:$false
    $result = $result.Content | ConvertFrom-Json
    if (!$result.value) { Write-Verbose "No groups found, skipping..." }

    $Groups = $result.value
    #Get the forms for each group
    foreach ($Group in $Groups) {
        #Does this even support pagination?
        $uri = "https://forms.office.com/formapi/api/$tenantId/groups/$($Group.Id)/light/forms?`$select=id,status,title,createdDate,modifiedDate,version,rowCount,softDeleted,type" #CASE SESNITIVE!
        try { $result = Invoke-WebRequest -Uri $uri -Headers $authHeaderFormsROPC -UseBasicParsing -ErrorAction Stop -Verbose:$false }
        catch { Write-Error "Failed to retrieve forms for group $($Group.displayName), skipping..." -ErrorAction SilentlyContinue; $_ | fl * -Force; continue }
        $result = $result.Content | ConvertFrom-Json

        if (!$result.value) { Write-Verbose "No forms found for group $($Group.displayName), skipping..." ; continue }
        $GroupForms = $result.value
        $ReportForms += [PSCustomObject]@{
            Name = $Group.displayName
            Type = "Group"
            Identifier = $Group.emailAddress
            FormsCount = $GroupForms.count
        }

        #Fetch the details for each form
        foreach ($form in $GroupForms) {
            $uri = "https://forms.office.com/formapi/api/$tenantId/groups/$($Group.Id)/light/forms(`'$($form.id)`')"
            $result = (Invoke-WebRequest -Uri $uri -Headers $authHeaderFormsROPC -UseBasicParsing -ErrorAction Continue -Verbose:$false).Content | ConvertFrom-Json
            $form.rowCount = $result.rowCount
        }

        $GroupForms | select Id, Status, Title, CreatedDate, ModifiedDate, version, @{n="ResponseCount";e={$_.rowCount}}, @{n="IsDeleted";e={($_.softDeleted) ? "Yes" : "No"}}, type `
        ` | Export-Excel -ExcelPackage $excel -WorksheetName $($Group.emailAddress) -FreezeTopRow -AutoFilter -BoldTopRow -AutoSize -NoHyperLinkConversion TargetResource -PassThru > $null

        Start-Sleep -Milliseconds 100
    }
}

#Add the summary sheet to the XLSX file
$ReportForms | Export-Excel -ExcelPackage $excel -WorksheetName Overview -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -MoveToStart -PassThru > $null

#Make UPNs in the Overview sheet clickable
$sheet = $excel.Workbook.Worksheets["Overview"]

#Add a hyperlink to the MemberUpn/ActionCount columns
$cells = $sheet.Cells["C2:C"] #Gives just the populated cells
foreach ($cell in $cells) {
    #Process only rows corresponding to user objects
    $cellValue = $cell.Value
    if ($cell.Value.length -gt 31) { $cellValue = $cell.Value.Substring(0,31) }
    $otherCell = $sheet.Cells[$cell.Address.Replace("C","D")]
    if (!($otherCell.Value) -or ($otherCell.Value -eq "0")) { continue }

    if ($excel.Workbook.Worksheets[$cellValue]) {
        $targetAddress = $excel.Workbook.Worksheets[$cellValue].Cells["A1"].FullAddress
        $cell.Hyperlink = New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList $targetAddress, $cell.Value
        $otherCell.Hyperlink = New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList $targetAddress, $otherCell.Value
        $otherCell.Value = [int]$otherCell.Value #OMFG spent two hours on this

        $cell.Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
        $cell.Style.Font.Underline = $true
        $otherCell.Style.Font.Color.SetColor([System.Drawing.Color]::Blue)
        $otherCell.Style.Font.Underline = $true
    }
}

Export-Excel -ExcelPackage $excel -WorksheetName "Overview" -Show > $null
$excel.Dispose()
Write-Verbose "Finished..."
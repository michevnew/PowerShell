$AppId = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$client_secret = Get-Content .\ReportingAPIsecret.txt | ConvertTo-SecureString
$app_cred = New-Object System.Management.Automation.PsCredential($AppId, $client_secret)
$TenantId = "tenant.onmicrosoft.com"
$SPOURL = "https://tenant-admin.sharepoint.com"

$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $app_cred.GetNetworkCredential().Password
    grant_type    = "client_credentials"
}

try { $tokenRequest = Invoke-WebRequest -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing -ErrorAction Stop }
catch { Write-Host "Unable to obtain access token, aborting..."; return }

$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

$authHeader1 = @{
   'Content-Type'='application\json'
   'Authorization'="Bearer $token"
}

try {
    Connect-SPOService -Url $SPOURL
    $Sites = (Get-SPOSite -Template "GROUP#0" -IncludePersonalSite:$False -Limit All) + (Get-SPOSite -Template "TEAMCHANNEL#0" -IncludePersonalSite:$False -Limit All)
}
catch { Write-Host "Unable to connect to SharePoint Online, aborting..."; return }

$Report = @();$count = 1
Foreach ($Site in $Sites) {
    #Progress message
    $ActivityMessage = "Retrieving data for site $($Site.Title). Please wait..."
    $StatusMessage = ("Processing site {0} of {1}: {2}" -f $count, @($Sites).count, $Site.Url)
    $PercentComplete = ($count / @($Sites).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #Simple anti-throttling control
    Start-Sleep -Milliseconds 200

    #Get detailed info about the site in order to get the GroupID
    $Site = Get-SPOSite -Identity $Site.Url -Detailed
    if ($Site.Template -eq "GROUP#0") { $GroupId = $site.GroupId.Guid }
    else { $GroupId = $site.RelatedGroupId.Guid }

    # Check if the Office 365 Group exists
    try {
        $O365Group = Invoke-WebRequest -Headers $AuthHeader1 -Uri "https://graph.microsoft.com/beta/groups/$($GroupId)?$`expand=members" -ErrorAction Stop
        $O365Group = ($O365Group.Content | ConvertFrom-Json)
        }
    catch { $O365Group = $Null }

    #Get the Group details, members, etc
    if ($O365Group) {
        #If this is a Private channel, we need to treat it differently
        if ($Site.Template -eq "TEAMCHANNEL#0") {
            $TeamsEnabled = $true
            $channelName = $site.Title.Split([string[]]@(" - "),[System.StringSplitOptions]::RemoveEmptyEntries)[-1]
            $TeamPCs = Invoke-WebRequest -Headers $AuthHeader1 -Uri "https://graph.microsoft.com/beta/Teams/$GroupId/channels"
            $channelID = ($TeamPCs.Content | ConvertFrom-Json).value | ? {$_.MembershipType -eq "Private" -and $_.DisplayName -eq $channelName }

            $Teammembers = Invoke-WebRequest -Headers $authHeader1 -Uri "https://graph.microsoft.com/beta/Teams/$GroupId/channels/$($channelID.Id)/members"
            $Teammembers = ($Teammembers.Content | ConvertFrom-Json).Value
            $members = ($Teammembers | ? {"guest" -ne $_.roles}) | measure | select -ExpandProperty Count
            $owners = ($Teammembers | ? {$_.Roles -eq "owner"}) | measure | select -ExpandProperty Count
            $guests = $Teammembers | ? {$_.Roles -eq "Guest"} | measure | select -ExpandProperty Count
        }
        else {
            $members = $O365Group.members | ? {$_.UserType -eq "Member"} | measure | select -ExpandProperty count
            $owners = Invoke-WebRequest -Headers $AuthHeader1 -Uri "https://graph.microsoft.com/beta/groups/$GroupId/owners"
            $owners = ($owners.Content | ConvertFrom-Json).value.Count
            $guests = $O365Group.members | ? {$_.UserType -ne "Member"} | measure | select -ExpandProperty count

            # Check if the Group is Team-enabled
            Try { $Team = Invoke-WebRequest -Headers $AuthHeader1 -Uri "https://graph.microsoft.com/beta/Teams/$GroupId" -ErrorAction Stop; $TeamsEnabled = $true }
            Catch { $TeamsEnabled = $False }
    }}

    #Prepare the output
    $objSite = New-Object PSObject
    $objSite | Add-Member -MemberType NoteProperty -Name "Site" -Value $Site.Title
    $objSite | Add-Member -MemberType NoteProperty -Name "URL" -Value $Site.Url
    $objSite | Add-Member -MemberType NoteProperty -Name "LastContentModifiedDate" -Value $Site.LastContentModifiedDate
    if ($O365Group) {
        $objSite | Add-Member -MemberType NoteProperty -Name "GroupName" -Value (&{If($O365Group.DisplayName) {$O365Group.DisplayName} Else {"N/A"}})
        $objSite | Add-Member -MemberType NoteProperty -Name "Created" -Value (&{If($site.Template -eq "GROUP#0" -and $O365Group) {$O365Group.CreatedDateTime} Else {"N/A"}}) #Group creation time does not match PC creation time
        $objSite | Add-Member -MemberType NoteProperty -Name "Expires" -Value (&{If($site.Template -eq "GROUP#0" -and $O365Group) {$O365Group.expirationDateTime} Else {"N/A"}}) #Group expiration time does not match PC expiration time
        $objSite | Add-Member -MemberType NoteProperty -Name "Privacy" -Value $O365Group.visibility
        $objSite | Add-Member -MemberType NoteProperty -Name "Owners" -Value $owners
        $objSite | Add-Member -MemberType NoteProperty -Name "Members" -Value $members
        $objSite | Add-Member -MemberType NoteProperty -Name "Guests" -Value $guests
        $objSite | Add-Member -MemberType NoteProperty -Name "TeamEnabled" -Value $TeamsEnabled
    }
    $objSite | Add-Member -MemberType NoteProperty -Name "Error" -Value (&{If(!$O365Group) {"Failed to find Office 365 Group for site: $($Site.Title) ($GroupId)"} Else {""}})
    $Report += $objSite
}

$Report | Sort Site | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_TeamsSprawlReport.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
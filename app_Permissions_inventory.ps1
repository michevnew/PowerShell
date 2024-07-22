if (!(Get-Module AzureAD -ListAvailable | ? {($_.Version.Major -eq 2 -and $_.Version.Build -eq 0 -and $_.Version.Revision -gt 55) -or ($_.Version.Major -eq 2 -and $_.Version.Build -eq 1)})) { Write-Host -BackgroundColor Red "This script requires a recent version of the AzureAD PowerShell module. Download it here: https://www.powershellgallery.com/packages/AzureAD/"; return}
try { Get-AzureADTenantDetail | Out-Null }
catch { Connect-AzureAD | Out-Null }

Write-Host "Gathering information about Azure AD integrated applications..." -ForegroundColor Yellow
try { $ServicePrincipals = Get-AzureADServicePrincipal -All:$true | ? {$_.Tags -eq "WindowsAzureActiveDirectoryIntegratedApp"} }
catch { Write-Host "You must connect to Azure AD first!" -ForegroundColor Red -ErrorAction Stop }
$appPermissions = @();$i=0;

foreach ($ServicePrincipal in $ServicePrincipals) {
    $SPperm = Get-AzureADServicePrincipalOAuth2PermissionGrant -ObjectId $ServicePrincipal.ObjectId -All:$true

    $OAuthperm = @{};
    $assignedto = @();$resID = $null; $userId = $null;
    $objPermissions = New-Object PSObject

    $i++;Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Number" -Value $i
    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Application Name" -Value $ServicePrincipal.DisplayName
    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "ApplicationId" -Value $ServicePrincipal.AppId
    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Publisher" -Value $ServicePrincipal.PublisherName
    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Homepage" -Value $ServicePrincipal.Homepage
    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "ObjectId" -Value $ServicePrincipal.ObjectId
    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Enabled" -Value $ServicePrincipal.AccountEnabled
    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Valid until" -Value ($SPperm.ExpiryTime | select -Unique | Sort-Object -Descending | select -First 1)

    Write-Host $SPperm.Scope
    $SPperm | % {#CAN BE DIFFERNT FOR DIFFERENT USERS!
        $resID = (Get-AzureADObjectByObjectId -ObjectIds $_.ResourceId).DisplayName
        if ($_.PrincipalId) { $userId = "(" + (Get-AzureADObjectByObjectId -ObjectIds $_.PrincipalId).UserPrincipalName + ")" }
        $OAuthperm["[" + $resID + $userId + "]"] = (($_.Scope.Split(" ") | select -Unique) -join ",")
    }
    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Permissions" -Value (($OAuthperm.GetEnumerator()  | % { "$($_.Name):$($_.Value)" }) -join ";")

    if (($SPperm.ConsentType | select -Unique) -eq "AllPrincipals") { $assignedto += "All users (admin consent)" }
    try { $assignedto += (Get-AzureADObjectByObjectId -ObjectIds ($SPperm.PrincipalId | select -Unique)).UserPrincipalName }
    catch {}
    Add-Member -InputObject $objPermissions -MemberType NoteProperty -Name "Authorized By" -Value ($assignedto -join ", ")

    $appPermissions += $objPermissions
}

$appPermissions | select 'Application name', 'ApplicationId', 'Publisher', 'Homepage', 'ObjectId', 'Enabled', 'Authorized By', 'Permissions', 'Valid until' #| Export-CSV -nti -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_AppInventory.csv"
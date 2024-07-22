if (!(Get-Module AzureAD -ListAvailable | ? {($_.Version.Major -eq 2 -and $_.Version.Build -eq 0 -and $_.Version.Revision -ge 55) -or ($_.Version.Major -eq 2 -and $_.Version.Build -ge 1)})) { Write-Host -BackgroundColor Red "This script requires a recent version of the AzureAD PowerShell module. Download it here: https://www.powershellgallery.com/packages/AzureAD/"; return}
try { Get-AzureADTenantDetail | Out-Null }
catch { Connect-AzureAD | Out-Null }

$Users = Get-AzureADUser -All:$true -Filter "DirSyncEnabled eq true"
$arrOUs = @();$i=0;

foreach ($user in $Users) {
    $extprop = $User.ExtensionProperty
    $userDN = $extprop["onPremisesDistinguishedName"]
    if (!$userDN) { continue }
    $userOU = $userDN.Substring($userDN.IndexOf(",") +1)

    $objUser = New-Object PSObject
    $i++;Add-Member -InputObject $objUser -MemberType NoteProperty -Name "Number" -Value $i
    Add-Member -InputObject $objUser -MemberType NoteProperty -Name "User Name" -Value $User.DisplayName
    Add-Member -InputObject $objUser -MemberType NoteProperty -Name "UserID" -Value $User.UserPrincipalName
    Add-Member -InputObject $objUser -MemberType NoteProperty -Name "OU" -Value $userOU

    $arrOUs += $objUser
}

$arrOUs | select 'User Name',UserId,OU | Sort-Object OU | Export-CSV -nti -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_O365UsersOUReport.csv"
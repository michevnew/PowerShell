#Requires -Version 3.0
#Requires -Modules @{ ModuleName="Microsoft.Graph.Users"; ModuleVersion="1.19.0" }

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/5716/reporting-on-synchronized-users-ou-via-the-graph-sdk-for-powershell

Write-Verbose "Checking connectivity to Graph PowerShell..."
try { 
    if (!(Get-MgContext) -or !((Get-MgContext).Scopes.Contains("User.Read.All"))) {
        Write-Verbose "Not connected to the Microsoft Graph or the required permissions are missing!"
        Connect-MgGraph -Scopes User.Read.All -ErrorAction Stop | Out-Null #Directory.Read.All
    }
}
catch { Write-Error $_; return }
#Double-check required permissions
if (!((Get-MgContext).Scopes.Contains("User.Read.All"))) { Write-Error "The required permissions are missing, please re-consent!"; return }

#Cannot filter by onPremisesDistinguishedName server-side, just client-side.
$Users = Get-MgUser -All -Filter "onPremisesSyncEnabled eq true" -Property displayName,userPrincipalName,onPremisesDistinguishedName | select displayName,userPrincipalName,onPremisesDistinguishedName

$arrOUs = @();$i=0;

foreach ($user in $Users) {
    if (!$user.onPremisesDistinguishedName) { continue }
    $userOU = $user.onPremisesDistinguishedName -replace '^(?:.+?(?<!\\),){1}(.+)$', '$1' #regex courtesy of Stanvy

    $objUser = New-Object PSObject
    $i++;Add-Member -InputObject $objUser -MemberType NoteProperty -Name "Number" -Value $i
    Add-Member -InputObject $objUser -MemberType NoteProperty -Name "Display Name" -Value $User.DisplayName
    Add-Member -InputObject $objUser -MemberType NoteProperty -Name "UPN" -Value $User.UserPrincipalName
    Add-Member -InputObject $objUser -MemberType NoteProperty -Name "OU" -Value $userOU

    $arrOUs += $objUser
}

$arrOUs | select 'Display Name',UPN,OU | sort OU | Export-CSV -nti -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_O365UsersOUReport.csv"
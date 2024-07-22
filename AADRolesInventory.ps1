#Do a check for the AzureAD module
if (!(Get-Module AzureAD -ListAvailable)) { Write-Host -BackgroundColor Red "This script requires a recent version of the AzureAD PowerShell module. Download it here: https://www.powershellgallery.com/packages/AzureAD/"; return }

#Do a connectivity check
try { Get-AzureADTenantDetail | Out-Null }
catch { Connect-AzureAD | Out-Null }

#Collect a list of all active roles in the tenant
$AADRoles = Get-AzureADDirectoryRole
$RolesHash = @{}

#Cycle each role and gather a list of users and service principals assigned
foreach ($AADRole in $AADRoles) {
    $AADRoleMembers = Get-AzureADDirectoryRoleMember -ObjectId $AADRole.ObjectId
    #if no role members assigned, skip
    if (!$AADRoleMembers) { continue }

    foreach ($AADRoleMember in $AADRoleMembers) {
        #prepare the output
        if (!$RolesHash[$AADRoleMember.ObjectId]) {
            $RolesHash[$AADRoleMember.ObjectId] = @{
                "UserPrincipalName" = (&{If($AADRoleMember.ObjectType -eq "User") {$AADRoleMember.UserPrincipalName} Else {$AADRoleMember.AppId}})
                "DisplayName" = $AADRoleMember.DisplayName
                "Roles" = $AADRole.DisplayName
                }
            }
        #if the same object was returned as a member of previous role(s)
        else { $RolesHash[$AADRoleMember.ObjectId].Roles += $(", " + $AADRole.DisplayName)  }
    }
}

#format and export
$report = foreach ($key in ($RolesHash.Keys)) { $RolesHash[$key] | % { [PSCustomObject]$_ } }
$report | Sort-Object DisplayName | Export-CSV -nti -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_AzureADRoleInventory.csv"
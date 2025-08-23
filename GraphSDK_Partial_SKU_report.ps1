#region Variables
#List of free SKUs
#Update from https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference as needed
$freeSKUs = @("90d8b3f8-712e-4f7b-aa1e-62e7ae6cbe96","f30db892-07e9-47e9-837c-80727f46fd3d","a403ebcc-fae0-4ca2-8c8c-7a907fd6c235","16ddbbfc-09ea-4de2-b1d7-312db6112d70","87bbbc60-4754-4998-8c88-227dca264858","6470687e-a428-4b7a-bef2-8a291ad947c9","dcb1a3ae-b33f-4487-846a-a640262fadf4","3f9f06f5-3c31-472c-985f-62d9c10ec167","440eaaa8-b3e0-484b-a8be-62870b9ba70a") #"47794cd0-f0e5-45c5-9033-2eb6b5fc84e0"

#List of service plans to check for "partial" license
$ExOSP = @("efb87545-963c-4e0d-99df-69c6916d9eb0","4a82b400-a79f-41a4-b4e2-e94f5787b113","1126bef5-da20-4f07-b45e-ad25d2581aa8","9aaf7827-d63c-4b61-89c3-182f06f82e5c","fc52cc4b-ed7d-472d-bbe7-b081c23ecc56") #"90927877-dcff-4af6-b346-2332c0b15bb7","da040e0a-b393-4bea-bb76-928b3fa1cf5a"
$TeamsSP = @("57ff2da0-773e-42df-b2af-ffb7a2317929")
$SPOSP = @("902b47e5-dcb2-4fdc-858b-c63a90a2bdb9","5dbe027f-2339-4123-9542-606e4d348a72","63038b2c-28d0-45f6-bc36-33062963b498","6b5b6a67-fc72-4a1f-a2b5-beecf05de761","a1f3d0a8-84c0-4ae0-bae4-685917b8ab48","c7699d2e-19aa-44de-8edf-1736da088ca1","0a4983bb-d3e5-4a09-95d8-b2d0127b3df5") #"4c9efd0c-8de7-4c71-8295-9f5fdb0dd048","a361d6e2-509e-4e25-a8ad-950060064ef4","afcafa6a-d966-4462-918c-ec0b4e0fe642"
$OfficeSP = @("094e7854-93fc-4d55-b2c0-3ab5369ebdc1","43de0ff5-c92c-492b-9116-175376d08c38","de9234ff-6483-44d9-b15e-dca72fdd27af") #Untattended/device ones: "8d77e2d9-9e28-4450-8431-0def64078fc5","3c994f28-87d5-4273-b07a-eb6190852599","18dfd9bd-5214-4184-8123-c9822d81a9bc")
$SPs = ($TeamsSP + $SPOSP + $ExOSP + $OfficeSP) -join "|"
#endregion Variables

#region Helper Functions
function CheckFullLicense {
    param([PSTypeName('Selected.Microsoft.Graph.PowerShell.Models.MicrosoftGraphUser')][PSCustomObject]$User)
    $FullLicense = $false

    #assignedPlans returns historical data, so plans that are NOT part of any currently assigned SKU can be visible, along with their (often incorrect) status
    #As we cannot rely on the ServicePlanId alone, and we don't want to use Get-MgUserLicenseDetail, we will check the SKU ID against the assigned licenses
    $userSPs = @()
    foreach ($userSKU in $user.AssignedLicenses.SkuId) {
        $userSPs += ($SKUs | ? {$_.SkuId -eq $userSKU}).ServicePlans.ServicePlanId | Select-Object -Unique
    }

    #Iterate over each assigned plan and make sure the status is Enabled.
    foreach ($plan in $user.AssignedPlans) {
        #Skip plans not in the user's currently assigned SKUs
        if ($plan.servicePlanId -notin $userSPs) { continue }

        #Skip plans not in our list
        if ($plan.ServicePlanId -notmatch $SPs) { continue }
        else {
            #If the plan is in our list, make sure it is enabled and continue to the next one
            if ($plan.CapabilityStatus -eq "Enabled") {
                $FullLicense = $true
            }
            #Otherwise, return false and exit the function
            else {
                $FullLicense = $false
                return $FullLicense
            }
        }
    }
    #If we got here, all plans were parsed, and are either in enabled state, or were filtered out.
    #If the $FullLicense variable is still false, you might want to review the user's assigned licenses manually.
    if (!$FullLicense) { Write-Warning "Unexpected status for user $($user.UserPrincipalName), please review manually." }
    return $FullLicense
}
#Endregion Helper Functions

#Connect to the Graph
Connect-MgGraph -Scopes User.Read.All,LicenseAssignment.Read.All -NoWelcome

#Fetch a list of all SKUs in the tenant, we need it later on as assignedPlans returns historical data and skews results
$SKUs = Get-MgSubscribedSku

#Get all users with licenses
$list = Get-MgUser -Filter 'assignedLicenses/$count ne 0' -ConsistencyLevel eventual -CountVariable licensedUserCount -PageSize 999 -All -Property UserPrincipalName,DisplayName,AssignedLicenses,AssignedPlans | Select-Object -Property UserPrincipalName,DisplayName,AssignedLicenses,AssignedPlans

#We don't care about Free licenses, exclude users with only such assigned
$listNonFree = $list | ? {$_.AssignedLicenses.SkuId | ? {$freeSKUs -notcontains $_}}
$listFree = $list | ? {$_ -notin $listNonFree}

#Determine whether the "full" license is assigned
$listReduced = @()
$listFull = @()
foreach ($user in $listNonFree) {
    if ((CheckFullLicense -User $user)) {
        $listFull += $user
    }
    else {
        $listReduced += $user
    }
}

#Export the results to CSV files for further analysis if needed
$listReduced | Select-Object UserPrincipalName,DisplayName,@{n="AssignedLicenses";e={$_.AssignedLicenses.SkuId -join ";"}},@{n="AssignedPlans";e={ ($_.AssignedPlans | % {"[$($_.CapabilityStatus)]$($_.ServicePlanId)"}) -join ";"} } | Export-Csv -Path "$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_UsersWithPartialLicenses.csv" -NoTypeInformation -Encoding UTF8 -UseCulture

#region HTML export

#Helper function to convert SKU GUIDs to friendly names
function SkuToName {
    param([Guid[]]$SkuIds)
    $names = @()
    foreach ($id in $SkuIds) {
        $sku = $SKUs | ? {$_.SkuId -eq $id}
        if ($sku) { $names += "$($sku.SkuPartNumber) ($id)" }
        else { $names += $id }
    }
    return $names
}

# Generate HTML report
$htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>License Report - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</title>
    <style>
        body { font-family: Lucida Console, sans-serif; margin: 20px; }
        h1 { font-family: Ariel, sans-serif; color: #333; }
        h2 { font-family: Ariel, sans-serif; color: #666; margin-top: 30px; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { font-family: Ariel, sans-serif; background-color: #f2f2f2; font-weight: bold; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        .summary { font-family: Ariel, sans-serif; background-color: #e6f3ff; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        .red { color: #c00000; }
    </style>
</head>
<body>
    <h1>License Report</h1>
    <div class="summary">
        <p><strong>Total licensed users:</strong> $($list.Count)</p>
        <p><strong>Users with non-free licenses:</strong> $($listNonFree.count)</p>
        <p><strong>Users with free licenses:</strong> $($listFree.count)</p>
        <p><strong>Users with full licenses:</strong> $($listFull.count)</p>
        <p><strong>Users with partial licenses:</strong> $($listReduced.count)</p>
    </div>

    <h2>Users with Free Licenses ($($listFree.count))</h2></summary>
    <details>
    <summary></summary>
    <table>
        <tr><th>User Principal Name</th><th>Display Name</th><th>Assigned Licenses</th></tr>
"@

foreach ($user in $listFree) {
    $htmlContent += "<tr><td>$($user.UserPrincipalName)</td><td>$($user.DisplayName)</td><td>$((SkuToName $user.AssignedLicenses.SkuId) -join "<br>")</td></tr>"
}

$htmlContent += @"
    </table>
    </details>

    <h2>Users with Full Licenses ($($listFull.count))</h2>
    <details>
    <summary></summary>
    <table>
        <tr><th>User Principal Name</th><th>Display Name</th><th>Assigned Licenses</th></tr>
"@

foreach ($user in $listFull) {
    $htmlContent += "<tr><td>$($user.UserPrincipalName)</td><td>$($user.DisplayName)</td><td>$((SkuToName $user.AssignedLicenses.SkuId) -join "<br>")</td></tr>"
}

$htmlContent += @"
    </table>
    </details>

    <h2>Users with Partial Licenses ($($listReduced.count))</h2>
    <details open>
    <summary></summary>
    <table>
        <tr><th>User Principal Name</th><th>Display Name</th><th>Assigned Licenses</th><th>Assigned Plans</th></tr>
"@

foreach ($user in $listReduced) {
    $assignedPlans = ($user.AssignedPlans | % {"[$($_.CapabilityStatus)]$($_.ServicePlanId)"})

    $assignedSKUs = @()
    foreach ($sku in $user.AssignedLicenses) {
        $disabledPlans = @()
        if ($sku.DisabledPlans) {
            foreach ($plan in ($sku.DisabledPlans | Sort-Object)) {
                if ($plan -match $SPs) { $disabledPlans += "<span class=red>$plan</span>" }
                else { $disabledPlans += $plan }
            }
            $assignedSKUs += "<br><strong>$(SkuToName $sku.SkuId) with disabled plans:</strong><br> $($disabledPlans -join "<br>")<br>"
        }
        else { $assignedSKUs += "<br><strong>$(SkuToName $sku.SkuId)</strong><br>" }
    }
    $htmlContent += "<tr><td>$($user.UserPrincipalName)</td><td>$($user.DisplayName)</td><td>$assignedSKUs</td><td>$(($assignedPlans | Sort-Object) -join "<br>")</td></tr>"
}

$htmlContent += @"
    </table>
    </details>
</body>
</html>
"@

#Export to HTML file
$htmlFileName = "$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_LicenseReport.html"
$htmlContent | Out-File -FilePath $htmlFileName -Encoding UTF8
#endregion HTML export
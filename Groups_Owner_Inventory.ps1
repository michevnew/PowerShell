param([switch]$IncludeExchangeManagedBy)

#Make sure we try the AzureADPreview module first, as it surfaces more details...
Remove-Module AzureAD -ErrorAction SilentlyContinue

#Check whether the Azure AD module exists and load it
try { Import-Module AzureADPreview -ErrorAction Stop }
catch [System.IO.FileNotFoundException] {

    Write-Host "Azure AD Preview module not found... Checking for the Azure AD module." -ForegroundColor Cyan

    try { Import-Module AzureAD -ErrorAction Stop }
    catch [System.IO.FileNotFoundException] { Write-Host "This script requires the Azure AD PowerShell module. Download it here: https://www.powershellgallery.com/packages/AzureAD/" -ForegroundColor Red; return }
}

#Check for connectivity to Azure AD and authenticate if needed
try { Get-AzureADTenantDetail | Out-Null }
catch { Connect-AzureAD | Out-Null }

#Helper function for fetching Owner data from Azure AD
function Get-AzureADGroupOwnersInventory {

    [CmdletBinding()]
    Param()

    #Get the list of Groups and their Owners
    if (Get-Command Get-AzureADMSGroup -ErrorAction SilentlyContinue) {
        Write-Host "Using the Azure AD Preview module." -ForegroundColor Cyan
        $output = Get-AzureADMSGroup -All:$true | % { $_ | Add-Member "Owners" ((Get-AzureADGroupOwner -ObjectId $_.id).UserPrincipalName -join ";") -PassThru }
        $output | Sort-Object DisplayName | select DisplayName,MailEnabled,SecurityEnabled,GroupTypes,Owners,@{n="ObjectId";e={$_.Id}}
    }
    else {
        Write-Host "Using the Azure AD module." -ForegroundColor Cyan
        $output = Get-AzureADGroup -All:$true | % { $_ | Add-Member "Owners" ((Get-AzureADGroupOwner -ObjectId $_.ObjectId).UserPrincipalName -join ";") -PassThru }
        $output | Sort-Object DisplayName | select DisplayName,MailEnabled,SecurityEnabled,Owners,ObjectId
    }
}

#Helper function for fetching ManagedBy data from Exchange Online
function Get-ExchangeObjectsOwnersInventory {

    [CmdletBinding()]
    Param()

    #Confirm connectivity to Exchange Online
    try { $session = Get-PSSession -InstanceId (Get-OrganizationConfig).RunspaceId.Guid -ErrorAction Stop }
    catch {
        try {
            $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential (Get-Credential) -Authentication Basic -AllowRedirection -ErrorAction Stop
            Import-PSSession $session -ErrorAction Stop | Out-Null
            }
        catch { Write-Error "No active Exchange Online session detected, please connect to ExO first: https://technet.microsoft.com/en-us/library/jj984289(v=exchg.160).aspx" -ErrorAction Stop; return }
    }

    #Get a list of all recipients that support ManagedBy/Owner attribute
    $outputExchange = Invoke-Command -Session $session -ScriptBlock { Get-Recipient -RecipientTypeDetails MailUniversalSecurityGroup,MailUniversalDistributionGroup,DynamicDistributionGroup,RoomList,GroupMailbox | Select-Object -Property Displayname,ManagedBy,PrimarySMTPAddress,RecipientTypeDetails,ExternalDirectoryObjectId }

    #If no objects are returned from the above cmdlet, stop the script and inform the user
    if (!$outputExchange) { Write-Error "No recipients found" -ErrorAction Stop }

    #Add the Owner data for each object
    foreach ($o in $outputExchange) {
        $o | Add-Member "MailEnabled" "True"
        $o | Add-Member "Owners" (&{If ($o.ExternalDirectoryObjectId) {(Get-AzureADGroupOwner -ObjectId $o.ExternalDirectoryObjectId).UserPrincipalName -join ";"}})
        $o | Add-Member "SecurityEnabled" (&{If($o.RecipientTypeDetails.Value -eq "MailUniversalSecurityGroup") {"True"} else {"False"}})
    }

    #Return the output
    $outputExchange | Sort-Object DisplayName | select DisplayName,MailEnabled,SecurityEnabled,ManagedBy,Owners,RecipientTypeDetails,@{n="ObjectId";e={$_.ExternalDirectoryObjectId}}
}

#Get the Azure AD Owner report
Get-AzureADGroupOwnersInventory -OutVariable global:varOwners # | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_GroupOwners.csv" -NoTypeInformation
Write-Host "Azure AD Owner report data is stored in the `$varOwners global variable" -ForegroundColor Cyan

#Get the Exchange Online ManagedBy/Owner report
if ($IncludeExchangeManagedBy) {
    Write-Host "Fetching data from Exchange Online." -ForegroundColor Cyan
    Get-ExchangeObjectsOwnersInventory -OutVariable global:varOwnersExchange # | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_GroupOwnersExchange.csv" -NoTypeInformation
    Write-Host "Exchange Online ManagedBy/Owner report data is stored in the `$varOwnersExchange global variable" -ForegroundColor Cyan
}
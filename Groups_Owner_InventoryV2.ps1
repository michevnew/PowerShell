#Requires -Version 3.0
#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.0.0" }
#Requires -Modules @{ ModuleName="Microsoft.Graph.Groups"; ModuleVersion="1.19.0" }
#Requires -Modules @{ ModuleName="Microsoft.Graph.Users"; ModuleVersion="1.19.0" }

param([switch]$IncludeExchangeManagedBy)

#Connect to Graph PowerShell and make sure we run with Directory.Read.All permissions
if (!(Get-MgContext) -or !((Get-MgContext).Scopes.Contains("Directory.Read.All"))) {
    Write-Verbose "Not connected to the Microsoft Graph or the required permissions are missing!"
    Connect-MgGraph -Scopes Directory.Read.All -ErrorAction Stop | Out-Null
}

#Double-check required permissions
if (!((Get-MgContext).Scopes.Contains("Directory.Read.All"))) { Write-Error "The required permissions are missing, please re-consent!"; return }

#Helper function for fetching Owner data from Azure AD
function Get-AzureADGroupOwnersInventory {

    [CmdletBinding()]
    Param() #needed for -OutVariable

    $output = Get-MgGroup -All -Property id,displayName,groupTypes,securityEnabled,mailEnabled,membershipRule,isAssignableToRole,mail,assignedLicenses,owners -Expand 'owners($select=id,userPrincipalName)' -ErrorAction Stop

    #If no objects are returned from the above cmdlet, stop the script and inform the user
    if (!$output) { Write-Error "No group objects found" -ErrorAction Stop }

    $output | Sort-Object DisplayName | select DisplayName,Id,Mail,MailEnabled,SecurityEnabled,GroupTypes,@{n="Owners";e={$_.Owners.AdditionalProperties.userPrincipalName -join ","}}
}

#Helper function for fetching ManagedBy data from Exchange Online
function Get-ExchangeObjectsOwnersInventory {

    [CmdletBinding()]
    Param() #needed for -OutVariable

    #Confirm connectivity to Exchange Online.
    try { Get-EXORecipient -ResultSize 1 -ErrorAction Stop | Out-Null }
    catch {
        try { Connect-ExchangeOnline -CommandName Get-Recipient -SkipLoadingFormatData } #needs to be non-REST cmdlet
        catch { Write-Error "No active Exchange Online session detected. To connect to ExO: https://docs.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps";return }
    }

    #Get a list of all recipients that support ManagedBy/Owner attribute
    $outputExchange = Get-ExORecipient -ResultSize Unlimited -RecipientTypeDetails MailUniversalSecurityGroup,MailUniversalDistributionGroup,DynamicDistributionGroup,RoomList -Properties Displayname,ManagedBy,PrimarySMTPAddress,RecipientTypeDetails,ExternalDirectoryObjectId -ErrorAction Stop
    #For M365 Groups, ManagedBy should match Owners on AAD side...unless affected by the bug I reported to Graph TAP...

    #If no objects are returned from the above cmdlet, stop the script and inform the user
    if (!$outputExchange) { Write-Error "No recipients found" -ErrorAction Stop }

    #Add the Owner data for each object
    foreach ($o in $outputExchange) {
        $o | Add-Member "MailEnabled" "True"
        #Apparently DDGs can have ExternalDirectoryObjectId values now?!
        $o | Add-Member "Owners" (&{If ($o.ExternalDirectoryObjectId -and $o.RecipientTypeDetails -ne "DynamicDistributionGroup") {(Get-MgGroup -GroupId $o.ExternalDirectoryObjectId -Expand 'owners($select=userPrincipalName)' -Property id,owners).Owners.AdditionalProperties.userPrincipalName -join ","}})
        $o | Add-Member "SecurityEnabled" (&{If($o.RecipientTypeDetails -eq "MailUniversalSecurityGroup") {"True"} else {"False"}})
    }

    #Return the output
    $outputExchange | Sort-Object DisplayName | select DisplayName,@{n="Id";e={$_.ExternalDirectoryObjectId}},PrimarySMTPAddress,MailEnabled,SecurityEnabled,ManagedBy,Owners,RecipientTypeDetails
}

#Get the Azure AD Owner report
Get-AzureADGroupOwnersInventory -OutVariable global:varOwners # | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_GroupOwners.csv" -NoTypeInformation
Write-Host "Azure AD Owner report data is stored in the `$varOwners global variable" -ForegroundColor Cyan

#Get the Exchange Online ManagedBy/Owner report
if ($IncludeExchangeManagedBy) {
    Get-ExchangeObjectsOwnersInventory -OutVariable global:varOwnersExchange -ErrorAction Stop # | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_GroupOwnersExchange.csv" -NoTypeInformation
    Write-Host "Exchange Online ManagedBy/Owner report data is stored in the `$varOwnersExchange global variable" -ForegroundColor Cyan
}
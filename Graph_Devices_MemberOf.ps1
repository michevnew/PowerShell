#Set the authentication details
$tenantID = "tenant.onmicrosoft.com" #your tenantID or tenant root domain
$appID = "12345678-1234-1234-1234-1234567890AB" #the GUID of your app. For best result, use app with Directory.Read.All scope granted
$client_secret = "XXXXXXXXXXXXXXXXXXX" #client secret for the app

$body = @{
    client_id     = $AppId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $client_secret
    grant_type    = "client_credentials"
}

#Get a token
$authenticationResult = Invoke-WebRequest -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -ContentType "application/x-www-form-urlencoded" -Body $body -ErrorAction Stop
$token = ($authenticationResult.Content | ConvertFrom-Json).access_token
$authHeader = @{'Authorization'="Bearer $token"}

#Get a list of all devices, make sure to handle multiple pages
$GraphDevices = @()
$uri = "https://graph.microsoft.com/v1.0/devices?`$select=id,deviceId,displayName"
do {
    $result = Invoke-WebRequest -Headers $authHeader -Uri $uri -ErrorAction Stop
    $uri = $result.'@odata.nextLink'
    #If we are getting multiple pages, best add some delay to avoid throttling
    Start-Sleep -Milliseconds 500
    $GraphDevices += ($result.Content | ConvertFrom-Json).Value
} while ($uri)

#Loop over each device and prepare the output
$Output = @()
$count = 1; $PercentComplete = 0;
foreach ($device in $GraphDevices) {
    #Simple progress indicator
    $ActivityMessage = "Retrieving data for user $($device.displayName). Please wait..."
    $StatusMessage = ("Processing user {0} of {1}: {2}" -f $count, @($GraphDevices).count, $device.id)
    $PercentComplete = ($count / @($GraphDevices).count * 100)
    Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
    $count++

    #Prepare the query, make sure to enter each additional group attribute you want to reference later on
    $uri = "https://graph.microsoft.com/v1.0/devices/$($device.id)/transitivememberof?`$select=id,displayName,groupTypes,mailEnabled,securityEnabled,visibility"
    $DeviceGroups = Invoke-WebRequest -Headers $authHeader -Uri $uri -ErrorAction Stop
    
    #Prepare the device object information we will export, add additional properties as needed here
    $deviceinfo = New-Object psobject
    $deviceinfo | Add-Member -MemberType NoteProperty -Name "DeviceId" -Value $device.deviceId
    $deviceinfo | Add-Member -MemberType NoteProperty -Name "ObjectId" -Value $device.id
    $deviceinfo | Add-Member -MemberType NoteProperty -Name "DeviceDisplayName" -Value $device.displayName
    #We're using the group Id here, as it's an unique identifier. If you prefer, replace .Id with .displayName. You can also add additional details as needed, dont forget to "select" the corresponding properties above though.
    $deviceinfo | Add-Member -MemberType NoteProperty -Name "MemberOf" -Value (($DeviceGroups.Content | ConvertFrom-Json).Value.Id -join ";")

    $Output += $deviceinfo
}

#Export to CSV file in the current directory
$Output | Export-Csv -Path "$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_DevicesMemberOf.csv" -NoTypeInformation -Encoding UTF8 -UseCulture
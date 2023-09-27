Param([ValidateNotNullOrEmpty()][String]$Mailbox)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/5773/configure-an-auto-reply-rule-for-microsoft-365-mailboxes-via-ews

#Load the EWS Managed API DLL
$ewsAPIPath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
try {
    Test-Path $ewsAPIPath -ErrorAction Stop | Out-Null
    Add-Type -Path $ewsAPIPath -ErrorAction Stop | Out-Null
}
catch {
    Write-Error "Unable to load the Exchange Web Services Managed API binaries, check the path above..."
}

#region Authentication
#We use the client credentials flow as an example. For production use, REPLACE the code below with your preferred auth method. NEVER STORE CREDENTIALS IN PLAIN TEXT!!!

#Variables to configure
$tenantID = "tenant.onmicrosoft.com" #your tenantID or tenant root domain
$appID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #the GUID of your app. Needs Exchange.ManageAsApp scope
$client_secret = "verylongsecurestring" #client secret for the app

#Prepare token request
$url = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'

$body = @{
    grant_type = "client_credentials"
    client_id = $appID
    client_secret = $client_secret
    scope = "https://outlook.office365.com/.default"
}

#Obtain the token
Write-Verbose "Authenticating..."
try {
    $tokenRequest = Invoke-WebRequest -Method Post -Uri $url -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing -ErrorAction Stop
    $token = ($tokenRequest.Content | ConvertFrom-Json).access_token
}
catch { Write-Error "Unable to obtain access token, aborting..." -ErrorAction Stop; return }
#endregion Authentication

#Configure the EWS service object
$exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1) #Create a new EWS service object
$exchangeService.Url = "https://outlook.office365.com/EWS/Exchange.asmx" #Set the EWS endpoint
$exchangeService.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $token #Set the OAuth token as credentials

$exchangeService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $Mailbox) #Impersonate the mailbox

#Create the message template
$tmTemplateEmail = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $exchangeService #Create a new message object
$tmTemplateEmail.ItemClass = "IPM.Note.Rules.ReplyTemplate.Microsoft" #This is the message class for auto-reply templates
$tmTemplateEmail.IsAssociated = $true #Designate the message as FAI
$tmTemplateEmail.Subject = "This is an auto-reply message" #Update as needed
$htmlBodyString = "Dear sender, I am currently out of office." #Update as needed
$tmTemplateEmail.Body = New-Object Microsoft.Exchange.WebServices.Data.MessageBody($htmlBodyString)

$PidTagReplyTemplateId = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x65C2, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary) #PR_REPLY_TEMPLATE_ID
$tmTemplateEmail.SetExtendedProperty($PidTagReplyTemplateId, [System.Guid]::NewGuid().ToByteArray()) #Set the PR_REPLY_TEMPLATE_ID property to a new GUID
$tmTemplateEmail.Save([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox) #Save the message in the Inbox folder

#Create the rule
$nrNewInboxRule = New-Object Microsoft.Exchange.WebServices.Data.Rule #Create a new rule object
$nrNewInboxRule.DisplayName = "Auto Reply Rule" #Update as needed
$nrNewInboxRule.Actions.ServerReplyWithMessage = $tmTemplateEmail.Id #Set the reply template to the message we created above
$cnCreateNewRule = New-Object Microsoft.Exchange.WebServices.Data.createRuleOperation[] 1 #Create a new rule operation object
$cnCreateNewRule[0] = $nrNewInboxRule #Add the rule to the operation object
$exchangeService.UpdateInboxRules($cnCreateNewRule,$true) #Update the inbox rules
Write-Verbose "Created auto-reply rule for mailbox ""$Mailbox"""
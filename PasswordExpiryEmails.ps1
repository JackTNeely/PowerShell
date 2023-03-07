# Get the client ID of the managed identity for the Function app
$clientId = (Get-AzWebApp -ResourceGroupName $env:ResourceGroupName -Name $env:FunctionAppName).Identity.PrincipalId

# Get an access token for the Azure AD Graph API using the managed identity
$graphUrl = "https://graph.microsoft.com/"
$graphScopes = "https://graph.microsoft.com/.default"
$graphTokenEndpoint = "https://login.microsoftonline.com/$env:TenantId/oauth2/v2.0/token"
$graphBody = @{
    "client_id" = $clientId
    "scope" = $graphScopes
    "client_secret" = $env:ClientSecret
    "grant_type" = "client_credentials"
}
$graphResponse = Invoke-RestMethod -Method Post -Uri $graphTokenEndpoint -Body $graphBody
$graphToken = $graphResponse.access_token

# Construct the authorization header for the Azure AD Graph API
$graphAuthHeader = @{
    "Authorization" = "Bearer $graphToken"
}

# Retrieve the SMTP server, port, and email addresses from environment variables
$smtpServer = "outlook.office365.com"
$smtpPort = 587
$senderEmail = $env:SenderEmail
$recipientEmail = $env:RecipientEmail

# Load the email template and replace the placeholders with the actual email addresses
$template = Get-Content -Path $env:EmailTemplatePath
$body = $template -replace "@@USERNAME@@", $recipientEmail

# Define the email message
$emailMessage = @{
    To = $recipientEmail
    From = $senderEmail
    Subject = "Password Expiring Soon"
    Body = $body
    BodyAsHtml = $true
}

# Get the SMTP credentials using OAuth2 authentication
$smtpUrl = "https://outlook.office365.com/ews/exchange.asmx"
$smtpScopes = "https://outlook.office365.com/.default"
$smtpTokenEndpoint = "https://login.microsoftonline.com/$env:TenantId/oauth2/v2.0/token"
$smtpBody = @{
    "client_id" = $clientId
    "scope" = $smtpScopes
    "client_secret" = $env:ClientSecret
    "grant_type" = "client_credentials"
}
$smtpResponse = Invoke-RestMethod -Method Post -Uri $smtpTokenEndpoint -Body $smtpBody
$smtpToken = $smtpResponse.access_token

# Create the SMTP client and send the email
$smtpClient = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)
$smtpClient.EnableSsl = $true
$smtpClient.UseDefaultCredentials = $false
$smtpClient.Credentials = New-Object System.Net.NetworkCredential("", $smtpToken)
$smtpClient.Send($emailMessage.From, $emailMessage.To, $emailMessage.Subject, $emailMessage.Body)

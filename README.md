# PowerShell
PowerShell, Remote Administration, Microsoft, Office, 365, Azure, Azure Active Directory, Bulk Scripting, Automation

<br />

<div align="center">
  <h2>Install Azure CLI and All az cli Extensions</h2>
</div>

<pre>
# Copyright © Jack T. Neely, 06/29/2023
# This script downloads the Azure CLI, installs all az cli extensions, and updates all az cli extensions.
# Note: You may see many warnings during the installation process. Most of these warnings are harmless but be sure you understand the consequences when running this script.

# Begin script #
# Install az cli via MSI over PowerShell ( https://learn.microsoft.com/en-us/cli/azure/install-azure-cli-windows?tabs=powershell#install-or-update )
$ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -Uri https://aka.ms/installazurecliwindows -OutFile .\AzureCLI.msi; Start-Process msiexec.exe -Wait -ArgumentList '/I AzureCLI.msi /quiet'; Remove-Item .\AzureCLI.msi

# Set az config to automatically install az cli extensions if an extension command is run
az config set extension.use_dynamic_install=yes_without_prompt

# Set default network credentials to allow proxy services. (This may be necessary for some machines.)
(New-Object System.Net.WebClient).Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

# Obtain the list of all available extensions
$extensions = az extension list-available --output tsv --query '[].name'

# Install each extension based on name
foreach ($extension in $extensions) {
    try {
        Write-Output "Installing $extension"
        az extension add --name $extension
    } catch {
        Write-Error "An error occurred. Unable to install $extension."
    }
}

# Update each extension based on name
foreach ($extension in $extensions) {
    try {
        Write-Output "Installing $extension"
        az extension update --name $extension
    } catch {
        Write-Error "An error occurred. Unable to update $extension."
    }
}

Write-Host "Script completed." -ForegroundColor Green
# End script #
</pre>

<br />

<div align="center">
  <h2>Get <em>Discovered Apps</em> Matching 'powerbi' on Intune-Enrolled Windows Devices</h2>
</div>

<pre>
# Define Azure AD and App Details
$ClientID     = "<Your-Azure-AD-App-Registration-Client-Id>"
$ClientSecret = "<Your-Azure-AD-App-Registration-Client-Secret>"
$TenantID     = "<Your-Tenant-ID>"
$YourId       = "<Your-Azure-AD-User-ID-For-the-Graph-SendMail-Endpoint>"

$RequestTimeout = 120 # Request timeout in seconds

# Start local log transcript
$LogFilePath = Join-Path $PSScriptRoot ".log"
Start-Transcript -Path $LogFilePath

# Get Access Token
$TokenUrl = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
$TokenBody = @{
    Client_Id     = $ClientID
    Scope         = "https://graph.microsoft.com/.default"
    Client_Secret = $ClientSecret
    Grant_Type    = "client_credentials"
}
$TokenResponse = Invoke-RestMethod -Uri $TokenUrl -Method Post -Body $TokenBody -TimeoutSec $RequestTimeout
$AccessToken = $TokenResponse.Access_Token

# Define Headers for Graph API
$Headers = @{
    Authorization = "Bearer $AccessToken"
}

# Query Microsoft Graph API for Managed Devices
$BaseUrl = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
$Filter = "`$filter=operatingSystem+eq+'Windows'"
$ApiUrl = $BaseUrl + '?' + $Filter

Function Invoke-WithRetry {
    Param (
        [ScriptBlock]$ScriptBlock,
        [Int]$MaxRetries = 5,
        [Int]$TimeoutSec = 60
    )

    $RetryCount = 0
    Do {
        Try {
            Return & $ScriptBlock
        }
        Catch {
            $IsThrottled = $_.Exception.Response.StatusCode -eq 'TooManyRequests' -or $_.Exception.Response.StatusCode -eq '429'
            $IsConnectionClosed = $_.Exception.Message -like "*unexpected EOF or 0 bytes from the transport stream*"
            If ($IsThrottled -or $IsConnectionClosed) {
                $RetryCount++
                $BaseWaitTime = If ($_.Exception.Response.Headers['Retry-After']) { 
                    $_.Exception.Response.Headers['Retry-After']
                }
                Else { [Math]::Pow(2, $RetryCount) }
                # Adding randomness to the wait time
                $WaitTime = $BaseWaitTime + (Get-Random -Minimum 1 -Maximum 5)

                Write-Host "Request throttled or connection closed. Retrying in $WaitTime seconds..."
                Start-Sleep -Seconds $WaitTime
            }
            Else {
                Throw $_
            }
        }
    } While ($RetryCount -lt $MaxRetries)

    Throw "Maximum number of retries reached."
}

# Fetch the devices (Filter for Windows Devices)
$ManagedDevicesResponse = @()
$Top = "`$top=100"
$NextPageUrl = $ApiUrl + '&' + $Top

Do {
    Write-Host "Querying URL: $NextPageUrl"

    $Response = Invoke-WithRetry -ScriptBlock {
        Invoke-RestMethod -Uri $NextPageUrl -Headers $Headers -TimeoutSec $RequestTimeout
    } -TimeoutSec $RequestTimeout

    $ManagedDevicesResponse += $Response.Value

    $NextPageUrl = $Response."@odata.nextLink"
} While ($NextPageUrl)

# Loop Through Devices and Retrieve Information
$ReportData = @()
$TotalDevices = $ManagedDevicesResponse.Count
$Counter = 0

ForEach ($Device in $ManagedDevicesResponse) {
    try {
        $Counter++
        Write-Progress -Activity "Processing Devices" -Status "$Counter of $TotalDevices" -PercentComplete (($Counter / $TotalDevices) * 100)
        $DeviceId = $Device.Id

        # Define the URL for fetching detected apps for this device
        $DetectedAppsUrl = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$DeviceId/detectedApps"

        Write-Host "Fetching apps for DeviceId $DeviceId."

        # Fetch all detected apps for this device
        $InstalledApps = @()
        Do {
            $Response = Invoke-WithRetry -ScriptBlock {
                Invoke-RestMethod -Uri $DetectedAppsUrl -Headers $Headers -TimeoutSec 60
            }
            If ($Response.Value) {
                $InstalledApps += $Response.Value
            }
            $DetectedAppsUrl = $Response.'@odata.nextLink'
        } While ($DetectedAppsUrl)

        Write-Host "Apps detected on device $DeviceId : $($InstalledApps.Count)"

        # Filter Apps Containing "powerbi" and Add to Report Data
        $MatchingApps = $InstalledApps | Where-Object { $_.DisplayName -like "*powerbi*" }
        Write-Host "Apps containing 'powerbi' on device $DeviceId : $($MatchingApps.Count)"
        ForEach ($App in $MatchingApps) {
            $ReportData += [PSCustomObject]@{
                UserPrincipalName = $Device.UserPrincipalName
                DeviceName        = $Device.DeviceName
                Manufacturer      = $Device.Manufacturer
                Model             = $Device.Model
                OperatingSystem   = $Device.OperatingSystem
                AppName           = $App.DisplayName
                AppVersion        = $App.Version
            }
        }
    }
    catch {
        Write-Host "An error occurred processing device $DeviceId. Error: $($_.Exception.Message)"
        continue
    }
}

# Sort the Report Data
$SortedReportData = $ReportData | Sort-Object DeviceName, UserPrincipalName, AppName, AppVersion, OperatingSystem

# Export Sorted Report Data to CSV
$CurrentDate = Get-Date -Format "MM-dd-yyyy"
$CsvPath = "PowerBIReportedInstallations-$CurrentDate.csv"
$SortedReportData | Export-Csv -Path $CsvPath -NoTypeInformation

# Setup for sending the email
$SendMailUrl = "https://graph.microsoft.com/v1.0/users/$YourId/sendMail"
$EmailHeaders = @{
    Authorization  = "Bearer $AccessToken"
    "Content-Type" = "application/json"
}

$EmailBody = @{
    Message         = @{
        Subject      = "Power BI Reported Installations - $CurrentDate"
        ToRecipients = @(
            @{
                EmailAddress = @{
                    Address = "<Your-Sender-Email-Address>"
                }
            }
            @{
                EmailAddress = @{
                    Address = "<Enter-a-Valid-Email-Address-For-TO-Recipient>"
                }
            }
        )
        CcRecipients = @(
            @{
                EmailAddress = @{
                    Address = "<Enter-a-Valid-Email-Address-For-CC-Recipient>"
                }
            }
        )
        From         = @{
            EmailAddress = @{
                Address = "<Enter-a-Valid-From-Email-Address-if-Sending-As-Another-Mailbox>"
            }
        }
        Attachments  = @(
            @{
                '@odata.type' = "#microsoft.graph.fileAttachment"
                ContentBytes  = [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($CsvPath))
                ContentType   = "text/csv"
                Name          = "PowerBIReportedInstallations-$CurrentDate.csv"
            }
        )
    }
    SaveToSentItems = $true
} | ConvertTo-Json -Depth 10

# Send the email
Invoke-WithRetry -ScriptBlock {
    Invoke-RestMethod -Uri $SendMailUrl -Headers $EmailHeaders -Method Post -Body $EmailBody -TimeoutSec $RequestTimeout
} -TimeoutSec $RequestTimeout

Write-Host "Email sent."
Write-Host "Script completed."

# Stop local log transcript
Stop-Transcript
</pre>

<br />

<div align="center">
  <h2>Sort HubSpot Leads List (CSV file) by Users@Domains using Microsoft 365 and Export to a New CSV</h2>
</div>

<pre>
Write-Host "`r`n`r`n`r`n`r`nWhat is the full path for your file? `r`nE.g. C:\leads.csv" -ForegroundColor Magenta
$path = Read-Host -Prompt "Path"

$leads = Import-Csv -Path $path -Header "Record ID","First Name","Last Name","Email","Office Number","Primary Associated Company ID","Associated Company","Record ID - Company","Company name","Cell Number","City","Industry"
$count = $leads.count
$results = @()

Write-Host "`r`n`r`nProcessing . . . Do not exit." -ForegroundColor Magenta
Write-Host "This could take some time. Have a coffee. Maybe a few." -ForegroundColor Cyan
Start-Sleep -Seconds 3

$i = 0

foreach ($lead in $leads) {

    $domain = ($lead.Email).Split("@")[1]
    $autodiscover = Resolve-DnsName -Name autodiscover.$domain -Type CNAME -DnsOnly -QuickTimeout -ErrorAction SilentlyContinue
    $i++

    if ($autodiscover.NameHost -eq "autodiscover.outlook.com") {
        $results += $lead

        #Write-Host $domain -ForegroundColor Cyan
        #Write-Host $lead.Email -ForegroundColor Cyan
        #Write-Host $autodiscover.NameHost "`r`n"
        #Write-Host $results | ft

    } else {}

    Write-Host "$i out of $count"
}

$results | Export-Csv -Path C:\M365_Leads.csv -NoTypeInformation
Write-Host "`r`n`r`n`r`n`r`nScript complete!" -ForegroundColor Magenta
Write-Host "Results output to C:\M365_Leads.csv" -ForegroundColor Cyan
Write-Host "`r`nPress enter to exit." -ForegroundColor Magenta
</pre>

<br />

<div align="center">
<a href="https://lh3.googleusercontent.com/un87V2kkHHTlXhk6KwKgKygUCrtTzr4L-ikNGoCc5YZUIuBIKzQJ95o-70sTVKmcIQPCG6mBYBkHfLkrkeGp_Brli001wgi2wB-iNWhQa8yFHCb1e97a9eG-S8IKyWBm8Q_pszMfKPY=w2400?source=screenshot.guru" target="_blank"><img src="https://lh3.googleusercontent.com/un87V2kkHHTlXhk6KwKgKygUCrtTzr4L-ikNGoCc5YZUIuBIKzQJ95o-70sTVKmcIQPCG6mBYBkHfLkrkeGp_Brli001wgi2wB-iNWhQa8yFHCb1e97a9eG-S8IKyWBm8Q_pszMfKPY=w2400?source=screenshot.guru" />
</div>

<br />
<br />
<br />
    
  
<div align="center">
  <h2>Break/Fix Local Outlook Autodiscover Malfunction by Validating DNS and Modifying Machine Registry</h2>
</div>
<br />
- <b>Validate CNAME autodisocver record, modify client registry, and repair Outlook autodiscover based on type of Exchange services (using email address as input).</b>
<br />Note: This PowerShell script gets passed into cmd, allowing a seamless click-to-run user experience. 


<pre>
@echo off
PowerShell -Command "Get-Content '%~dpnx0' | Select-Object -Skip 3 | Out-String | Invoke-Expression"
goto :eof
Write-Host "`r`n`r`nWhat is your email address?" -ForegroundColor Cyan
$domain = (Read-Host).Split("@")[1]

Write-Host "`r`n`r`nFetching CNAME record for autodiscover.$domain . . ." -ForegroundColor Cyan
$d = Resolve-DnsName -Name autodiscover.$domain -Type CNAME -DnsOnly -QuickTimeout -ErrorAction SilentlyContinue
$d | ft Name,NameHost

$dS = ($d.NameHost -eq "autodiscover.outlook.com") -or ($d.NameHost -eq "adr.exghost.com")
if ($d.NameHost -eq $null ) { Write-Host "`r`n`r`nAutodiscover record not published. Please contact your administrator or try again." -ForegroundColor Cyan }
elseif (!$dS) { Write-Host "`r`n`r`nAutodiscover record not valid. Please contact your administrator or try again." -ForegroundColor Cyan }
else { Write-Host "`r`n`r`nAutodiscover record valid"

Write-Host "`r`n`r`nTesting connection for autodiscover.$domain . . ." -ForegroundColor Cyan
$p = Test-Connection autodiscover.$domain -Count 1 -Quiet

If (!$p) { Do { Write-Host "`r`nping autodiscover.$domain was NOT successful.`r`nTesting connection again . . ."
                                $p = Test-Connection autodiscover.$domain -Count 1 -Quiet
                                Write-Host "`r`n`r`n( Press Ctrl C to cancel )"
                                } until ($p) }
Write-Host "`r`nping autodiscover.$domain successful`r`n`r`n"

if ($dS) {
            If ($dS -and $p -and $d.NameHost -eq "autodiscover.outlook.com") { Write-Host "No autodiscover issues detected.`r`nMicrosoft 365 Autodiscover Registry Edit is recommended." } 
            elseif ($dS -and $p -and $d.NameHost -eq "adr.exghost.com") { Write-Host "No autodiscover issues detected.`r`nHosted Exchange Autodiscover Registry Edit is recommended." }
            else { Write-Host "Something went wrong." } }

            Write-Host "`r`n`r`nBacking up registry . . ." -ForegroundColor Cyan

            Set-Location -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\AutoDiscover

            $reg = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\AutoDiscover'
            $out = "$temp\backup.reg"

            reg export HKCU\SOFTWARE\Microsoft\Office\16.0\Outlook\AutoDiscover C:\Temp\backup.reg /y

            Write-Host "`r`n`r`n$reg backed up to $out" -ForegroundColor Cyan 
            Write-Host "`r`nPrevious registry can be restored by double clicking this file to run.`r`nThis will reverse the following registry changes:`r`n`r`n" 

            Set-ItemProperty -Path $reg -Name 'PreferLocalXML' -Value 0
            Set-ItemProperty -Path $reg -Name 'ExcludeHTTPRedirect' -Value 0
            Set-ItemProperty -Path $reg -Name 'PreferLocalXML' -Value 1
            Set-ItemProperty -Path $reg -Name 'ExcludeHttpsAutoDiscoverDomain' -Value 1
            Set-ItemProperty -Path $reg -Name 'ExcludeScpLookup' -Value 1
            Set-ItemProperty -Path $reg -Name 'ExcludeSrvLookup' -Value 1
            Set-ItemProperty -Path $reg -Name 'ExcludeSrvRecord' -Value 1


            if ($d.NameHost -eq "autodiscover.outlook.com") { 
                Set-ItemProperty -Path $reg -Name 'ExcludeExplicitO365Endpoint' -Value 0
                Write-Host "`r`n`r`nM365 Autodiscover Registry Edit successfully executed.`r`n`r`nProgram complete." -ForegroundColor Cyan }

            elseif ($d.NameHost -eq "adr.exghost.com") { 
                Set-ItemProperty -Path $reg -Name 'ExcludeExplicitO365Endpoint' -Value 1
                Write-Host "`r`n`r`nHosted Exchange Autodiscover Registry Edit successfully executed.`r`n`r`nProgram complete." -ForegroundColor Cyan }

            else { Write-Host "`r`n`r`nSomething went wrong. Please terminate this program." -ForegroundColor Cyan }
            }
Write-Host "`r`n`r`n`r`n`r`nPress enter to exit.`r`n`r`n`r`n" -ForegroundColor Magenta
Read-Host
</pre>
  
  
<br />
<br />
<br />

<div align="center">
<h2>Bulk Remove Automapping for All Users</h2>
</div>
  
<pre>  
$objUsers = Get-Mailbox -Resultsize Unlimited
Foreach ($objUser in $objUsers){
        $RemoveAutoMapping = Get-MailboxPermission -Identity $($objUser.Alias) | Where 
            {$.AccessRights -eq "FullAccess"} 
        $RemoveAutoMapping | Remove-MailboxPermission
        $RemoveAutoMapping | Foreach 
            {Add-MailboxPermission -Identity $.Identity -User $.User -AccessRights:FullAccess -AutoMapping $False}
    }
</pre>

<br/>
<br/>

<div align="center">
  <h2>Much More to Come!</h2>
</div>

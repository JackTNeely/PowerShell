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

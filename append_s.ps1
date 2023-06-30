# Copyright © Jack T. Neely, 06/30/2023

# This script appends '_s' to the name of each column in a flat JSON file
# This script is intended only for scenarios where you are ingesting custom log data via the Data Collector API, and where '_s' is appended to all custom column names.
# The purpose of the script is to support future ingestions of custom columns to Azure Log Analytics tables after migrating from the Data Collector API to the Log Ingestion API

# Before using this script, a POST request should first be issued against the specified table to convert/migrate the table to support usage of the Log Ingestion API
# POST https://management.azure.com/subscriptions/{subscriptionId}/resourcegroups/{resourceGroupName}/providers/Microsoft.OperationalInsights/workspaces/{workspaceName}/tables/{tableName}/migrate?api-version=2021-12-01-preview
# Reference: https://learn.microsoft.com/en-us/azure/azure-monitor/logs/custom-logs-migrate

# Don't forget to follow the prerequisite steps outlined in Microsoft's documentation
# This includes the creation an Azure AD app registration and API scopes, a Data Collection Endpoint for the Log Ingestion API, a new Data Collection Rule (I recommend ARM template deployments), and IAM permissions for the DCR
# Ref: https://learn.microsoft.com/en-us/azure/azure-monitor/logs/custom-logs-migrate
# Ref: https://learn.microsoft.com/en-us/azure/azure-monitor/logs/tutorial-logs-ingestion-api
# Try not to get too bogged down in the docs - the documentation for Log Analytics can be confusing. If you need help, try to open a support ticket with Microsoft if possible.

# Feel free to modify this script to suit your needs. You may have other custom data typed fields, like boolean ('_b') as an example. Or you may want to change to logic for providing the JSON.
# E.g., This script prompts for a JSON file at the host. You may want to feed it raw JSON, or add automation logic.

# Load the necessary assembly for working with JSON
Add-Type -AssemblyName System.Web.Extensions

# Prompt the user for the file path
$OriginalFilePath = Read-Host -Prompt 'Enter the path of the JSON file'

# Read the JSON file
$JsonString = Get-Content -Path $OriginalFilePath -Raw

# Convert the JSON to a PowerShell object
$JsonObject = ConvertFrom-Json -InputObject $JsonString

# Create a new empty object
$NewJsonObject = @()

# Check if $JsonObject is an array
if ($JsonObject -is [System.Array]) {
    foreach ($Object in $JsonObject) {
        $TempObject = New-Object -TypeName PSObject
        foreach ($Property in $Object.PSObject.Properties) {
            Add-Member -InputObject $TempObject -NotePropertyName ($Property.Name + "_s") -NotePropertyValue $Property.Value
        }
        $NewJsonObject += $TempObject
    }
}
else {
    $TempObject = New-Object -TypeName PSObject
    foreach ($Property in $JsonObject.PSObject.Properties) {
        Add-Member -InputObject $TempObject -NotePropertyName ($Property.Name + "_s") -NotePropertyValue $Property.Value
    }
    $NewJsonObject += $TempObject
}

# Convert the new object back to JSON
$NewJsonString = ConvertTo-Json -InputObject $NewJsonObject -Depth 100

# Create the new file path
$Directory = Split-Path -Path $OriginalFilePath -Parent
$OriginalFileName = Split-Path -Path $OriginalFilePath -Leaf
$NewFileName = $OriginalFileName -replace '\.json$', '_s.json'
$NewFilePath = Join-Path -Path $Directory -ChildPath $NewFileName

# Write the new JSON to a file
$NewJsonString | Out-File -FilePath $NewFilePath

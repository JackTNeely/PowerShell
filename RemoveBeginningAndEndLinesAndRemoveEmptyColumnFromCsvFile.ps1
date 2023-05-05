# Specify source CSV file path
$SourcePath = "C:\CsvLogs\Ascend\Ascend User Update for Splunk March 2023.csv"

# Specify destination path
$DestinationPath = "C:\CsvLogs\Ascend\Ascend_update.csv"

$ImportFile = Get-ChildItem $SourcePath

# Remove the first 2 lines and the last 9 lines from CSV file
(Get-Content $ImportFile | Select-Object -Skip 2 | Select-Object -SkipLast 9) | Set-Content $DestinationPath

# Import CSV data
$CsvData = Import-Csv -Path $DestinationPath

# Find header names and remove specified header (fifth column - E)
$Headers = $CsvData[0].PSObject.Properties.Name
$DesiredHeaders = $Headers[0..3] + $Headers[5..($Headers.Count - 1)]

# Remove specified column and shift remaining columns to the left
$ModifiedCsvData = $CsvData | Select-Object -Property $DesiredHeaders

# Export to a new CSV file
$ModifiedCsvData | Export-Csv -Path $DestinationPath -NoTypeInformation

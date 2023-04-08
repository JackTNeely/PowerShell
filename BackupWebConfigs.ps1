# Source folder to search for 'web.config' files
$SourceFolder = "E:\Program Files\Microsoft\Exchange Server\V15"

# Specify destination folder to copy files for backup
$DestinationFolder = "$env:userprofile\Desktop\Backups"

# Specify file name to backup
$FileName = "web.config"

# Gather list of all files named 'web.config' (recursively)
$BackupFiles = Get-ChildItem -Path $SourceFolder -Recurse -Filter $FileName -File

# Loop through each file and create a backup in the destination using same folder hierarchy
foreach ($File in $BackupFiles) {
    $DestinationPath = Join-Path $DestinationFolder ($File.FullName -replace [regex]::Escape($SourceFolder), "")
    $DestinationDirectory = Split-Path $DestinationPath
    New-Item -ItemType Directory -Path $DestinationDirectory -Force | Out-Null
    Copy-Item $File.FullName $DestinationPath -Force
}

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

<#
.SYNOPSIS
    Save a secret (e.g., an Azure service principal secret) in the PowerShell Secret Store for secure retrieval in other scripts.
.DESCRIPTION
    This script demonstrates how to save a secret (e.g., an Azure service principal secret) in the PowerShell Secret Store using the Microsoft.PowerShell.SecretManagement and Microsoft.PowerShell.SecretStore modules. 
    The secret can then be securely retrieved in other scripts (like 1_send-from-geico.ps1) without hardcoding sensitive information.
    This is not meant to be run multiple times, but rather as a one-time setup to store the secret securely. After running this script, you can retrieve the secret in your other scripts using Get-Secret cmdlet.
.NOTES
    - These modules need to be pre-installed:
      Install-Module 'Microsoft.PowerShell.SecretManagement', 'Microsoft.PowerShell.SecretStore'
    - IMPORTANT: Remember to delete or change the hard-coded secret in this script after running it to avoid security risks.
#>

# Just a weird behavior I didn't expect
Write-Host "Expect this script to ask for a password twice, but it will be removed after you enter it the second time. This is normal behavior for setting up the Secret Store." -ForegroundColor Green

# Run once to save your secret in the PS Secret Store.
$secret = "some-service-principal-secret"  # Replace with your actual secret
$secretStore = 'store-name`'
Register-SecretVault -Name 'SecretStore' -ModuleName 'Microsoft.PowerShell.SecretStore' -DefaultVault
Set-SecretStoreConfiguration -Authentication 'None'  # For unattended scripts
Set-Secret -Name $secretStore -Secret $secret -Vault 'SecretStore'

# Password: 080808

# Remind yourself to delete/change the hard-coded secret after running this script
Write-Host "Secret has been saved in the PS Secret Store. Please remember to delete or change the hard-coded secret in this script to avoid security risks." -ForegroundColor Yellow

<#
To remove: Remove-Secret -Name 'store-name' -Vault 'SecretStore'
#>
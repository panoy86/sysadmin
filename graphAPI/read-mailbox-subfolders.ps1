<#
.SYNOPSIS
    This script reads all subfolders under the Inbox (or a specified folder) of a mailbox and reads the contents
    of those subfolders using Microsoft Graph API. Limits itself to the last 5 days of messages.
.NOTES
    Requires a service principal and Microsoft Graph API permissions to read mailbox folders and messages.
#>

# Script-wide variables
$script:Token = $null
$script:ListOfFolders = @() # Initialize an array to hold the list of folders retrieved from the mailbox

#------------------------------------------------------------------------------
# Authenticate to Graph API using an Azure app/secret combination
#------------------------------------------------------------------------------
function Connect-ToGraph
{
    param (
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret
    )
    $body = @{
        grant_type    = "client_credentials"
        scope         = "https://graph.microsoft.com/.default"
        client_id     = $ClientId
        client_secret = $ClientSecret
    }
    try {
        $response = Invoke-WebRequest -Method Post `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
            -Body $body -ContentType "application/x-www-form-urlencoded" -UseBasicParsing -ErrorAction Stop
        $script:Token = ($response.Content | ConvertFrom-Json).access_token
    }
    catch {
        Write-Host "Failed to acquire token for Graph API. Please check your tenant ID, client ID and client secret."
        throw $_
    }
}

#------------------------------------------------------------------------------
# Gets the sub-folders of a specified folder (default is Inbox)
#------------------------------------------------------------------------------
function Get-MailboxSubFolders
{
    param (
        [string]$RecipientEmail,
        [string]$ParentFolderId = "Inbox"
    )
    Write-Progress -Activity "Getting subfolders" -Status $ParentFolderId
    $listFolders = @()  # Initialize an array to hold the folders retrieved from the mailbox
    $uri = "https://graph.microsoft.com/v1.0/users/$RecipientEmail/mailFolders/$ParentFolderId/childFolders"
    try {
        do {
            $response = Invoke-WebRequest -Method Get -Uri $uri `
                -Headers @{ Authorization = "Bearer $script:Token" } `
                -UseBasicParsing -ErrorAction Stop
            [array]$listFolders += ($response.Content | ConvertFrom-Json).value
            $uri = ($response.Content | ConvertFrom-Json)."@odata.nextLink"
        } while ($uri)
    }
    catch {
        Write-Host "Failed to retrieve subfolders for $RecipientEmail. Please check if the recipient email is correct and if the app has the necessary permissions."
        throw $_
    }
    # Do a recursive call for each folder to get its subfolders as well, until there are no more child folders
    foreach ($folder in $listFolders) {
        $script:ListOfFolders += $folder
        Get-MailboxSubFolders -RecipientEmail $RecipientEmail -ParentFolderId $folder.id
    }
    Write-Progress -Activity "Getting subfolders" -Completed
}

#------------------------------------------------------------------------------
# Gets the content of a specified folder (defaults to Inbox) from the last 5 days only
#------------------------------------------------------------------------------
function Get-MailboxFolderContent
{
    param (
        [string]$RecipientEmail,
        [string]$FolderId = "Inbox",
        [string]$DaysBack = 5
    )
    #Write-Host "Retrieving content (past $DaysBack days only) for folder $FolderId of" $RecipientEmail -ForegroundColor Cyan
    $listMessages = @()  # Initialize an array to hold the messages retrieved from the mailbox
    $uri = "https://graph.microsoft.com/v1.0/users/$RecipientEmail/mailFolders/$FolderId/messages?`$filter=receivedDateTime ge " `
        + (Get-Date).AddDays(-$DaysBack).ToString("o")
    try {
        do {
            $response = Invoke-WebRequest -Method Get -Uri $uri `
                -Headers @{ Authorization = "Bearer $script:Token" } `
                -UseBasicParsing -ErrorAction Stop
            [array]$listMessages += ($response.Content | ConvertFrom-Json).value
            $uri = ($response.Content | ConvertFrom-Json)."@odata.nextLink"
        } while ($uri)
    }
    catch {
        Write-Host "Failed to retrieve folder content for $RecipientEmail. Please check if the recipient email is correct and if the app has the necessary permissions."
        throw $_
    }
    return $listMessages
}

#------------------------------------------------------------------------------
# Main script execution starts here
#------------------------------------------------------------------------------

# Set up progress bar and compatibility settings for different PowerShell versions
if ([int]$PSVersionTable.PSVersion.Major -ge 7) {
    Import-Module PSReadLine -Force     # Fixes progress bar issues in PowerShell 7+
    $PSStyle.Progress.View = "Minimal"  # Other value: "Minimal", only works in PowerShell 7.2+
}
$ProgressPreference = "Continue"

# Connect to Graph API using the specified app registration and secret.
<# PPGEICO
$clientId = "e6e0c5f4-be2c-418c-86a2-250bf44038f4"
$tenantId = "25798ea0-b97a-44b0-b1d8-3747bb6a5f3e"
$clientSecret = Get-Secret -Name 'busmes_ppgeico' -Vault 'SecretStore' -AsPlainText
Connect-ToGraph -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret
#>
$clientId = "55d8f305-44f4-4f18-bfff-3be2119a0247" # email-exchange-readyonly-appreg
$tenantId = "7389d8c0-3607-465c-a69f-7d4426502912" # GEICO
$clientSecret = Get-Secret -Name ("geicosecret_" + $clientId) -Vault "SecretStore" -AsPlainText
Connect-ToGraph -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret

# Get the sub-folders under the Inbox of the target mailbox
$emailAddress = "bethanymack@geico.com"
$parentFolderId = "Inbox"
$script:ListOfFolders = @() # Clear the list of folders before populating it
Get-MailboxSubFolders -RecipientEmail $emailAddress -ParentFolderId $parentFolderId

# Normalize the folder path for each
foreach ($folder in $script:ListOfFolders) {
    $folderPath = $folder.parentFolderId
    while ($folderPath -ne $parentFolderId) {
        $parentFolder = $script:ListOfFolders | Where-Object { $_.id -eq $folderPath }
        if ($parentFolder) {
            $folderPath = $parentFolder.parentFolderId
            $folder.displayName = "$($parentFolder.displayName)/$($folder.displayName)"
        }
        else {
            break
        }
    }
}

# For each sub-folder, get the content of that folder from the last 5 days only
$listFoldersMessages = @() # Initialize an array to hold all messages from all folders
$listFoldersMsgCount = @() # Initialize an array to hold the count of messages in each folder
$nCtr = 0
foreach ($folder in $script:ListOfFolders) {
    # Show progress
    $nCtr++
    Write-Progress -Activity "Getting messages from" -Status $folder.displayName `
        -PercentComplete (($nCtr / $script:ListOfFolders.Count) * 100)
    # Get the content of the folder and add it to our list of messages, along with the folder name for reference    
    $folderContent = @()
    $folderContent = Get-MailboxFolderContent -RecipientEmail $emailAddress -FolderId $folder.id -DaysBack 3
    foreach ($message in $folderContent) {
        $listFoldersMessages += New-Object PSObject -Property @{
            folder = $folder.displayName
            subject = $message.subject
            receivedDateTime = $message.createdDateTime
            from = $message.from.emailAddress.address
        }
    }
    # Also get the count of messages in the folder for reference
    $listFoldersMsgCount += New-Object PSObject -Property @{
        folder = $folder.displayName
        messageCount = $folderContent.Count
    }
}
Write-Progress -Activity "Getting messages from" -Status "Completed" -PercentComplete 100
#$listFoldersMessages
# $listFoldersMsgCount | Format-Table -AutoSize
$listFoldersMsgCount | Where-Object {$_.messageCount -gt 0} | Format-Table -AutoSize
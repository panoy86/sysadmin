<#
.SYNOPSIS
    Check the received emails in PPGEICO shared mailbox and update the tracker CSV file with the status of each email.
.DESCRIPTION
    This script is the second part of a three-script set that checks the received emails from a shared mailbox in PPGEICO. It references a tracker file that contains unique subject identifiers for the sent emails, and it updates the tracker file with the received date and status of each email. This script is meant to be run after 1_send-from-geico.ps1 has been executed to send the test emails.
.NOTES
    - CsvFileTracker has the Status property, with only 2 possible values: "Success" or "Fail".
    - Change the $script:MaxCheckAttempts variable in the script to adjust how many times the script will check for an email before marking it as "Fail" in the tracker.  If for example you set it to 3, the script will check for the email up to 3 times (each time you run the script) and if after the 3rd check the email is still not found, it will update the tracker with "Fail" status for that email entry.
#>

# Script-wide variables
$script:Token = $null
$script:CsvFileTrackerPath = "C:\Scripts\mailmonitor\EmailTracker.csv"  # Use full path (relative path will fail)
$script:LockFilePath = "C:\Scripts\mailmonitor\EmailTracker.lock"       # Path for lock file to manage concurrent access to CSV tracker file
$script:ListMessages = @()    # Initialize an array to store messages retrieved from the mailbox
$script:MaxCheckAttempts = 3  # Maximum number of check attempts for each email before marking as "Fail"

#------------------------------------------------------------------------------
#-- Authenticate to Graph API using an Azure app/secret combination
#------------------------------------------------------------------------------
function Connect-Graph
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
        $response = Invoke-WebRequest -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $body -ContentType "application/x-www-form-urlencoded"
        $script:Token = ($response.Content | ConvertFrom-Json).access_token
    }
    catch {
        Write-Host "Failed to acquire token for Graph API. Please check your tenant ID, client ID and client secret."
        throw $_
    }
}

#------------------------------------------------------------------------------
#-- Gets the Inbox content of a mailbox from the last 3 days only
#------------------------------------------------------------------------------
function Get-MailboxContent
{
    param (
        [string]$RecipientEmail
    )
    # Set the initial URI for retrieving messages from the recipient's mailbox, filtering for messages received in the last 3 days to limit results
    $uri = "https://graph.microsoft.com/v1.0/users/$RecipientEmail/mailFolders/Inbox/messages?`$filter=receivedDateTime ge " + (Get-Date).AddDays(-3).ToString("o")
    try {
        # Graph API may paginate results, so we need to loop through all pages to get the complete list of messages
        do {
            $response = Invoke-WebRequest -Method Get -Uri $uri -Headers @{ Authorization = "Bearer $script:Token" }
            [array]$script:ListMessages += ($response.Content | ConvertFrom-Json).value
            $uri = ($response.Content | ConvertFrom-Json)."@odata.nextLink"
        } while ($uri)
    }
    catch {
        Write-Host "Failed to retrieve mailbox content for $RecipientEmail. Please check if the recipient email is correct and if the app has the necessary permissions."
        throw $_
    }
}

#------------------------------------------------------------------------------
#-- Updates the CSV tracker file with the received date and status for each email entry
#------------------------------------------------------------------------------
function Update-CsvFileTracker
{
    # If the tracker file doesn't exist, throw an error
    if (-not (Test-Path -Path $script:CsvFileTrackerPath)) {
        throw "Tracker file does not exist: $script:CsvFileTrackerPath"
    }
    # Open the CSV tracker file with exclusive lock to prevent concurrent access issues
    $maxWait = 60 # seconds
    Write-Host "Attempting to acquire lock on CSV tracker file for update (60 seconds max wait)..."
    do {
        try {
            $lock = [System.IO.File]::Open($script:LockFilePath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            break
        }
        catch {
            Start-Sleep -Seconds 5
            $maxWait -= 5
            if ($maxWait -le 0) {
                Write-Host "Failed to acquire lock on CSV tracker file after multiple attempts. Exiting without updating tracker."
                return $false
            }
        }
    } while ($true)
    Write-Host "Lock acquired on CSV tracker file. Updating tracker..." -ForegroundColor Green
    # Create a hashtable to store the email unique identifiers and their received dates for quick lookup
    $hashEmails = @{}
    foreach ($message in $script:ListMessages) {
        $hashEmails[$message.subject] = ($message.receivedDateTime).ToUniversalTime().ToString("o")
    }
    # Read the existing CSV data
    $csvData = Import-Csv -Path $script:CsvFileTrackerPath
    # Update each entry in the CSV data with received date and status
    foreach ($entry in $csvData) {
        if ($entry.Status -notin ("Success", "Fail")) {
            # See if we need to set the Status to "Fail" due to max check attempts
            if ([int]$entry.CheckCount -ge $script:MaxCheckAttempts) {
                $entry.Status = "Fail"
                continue
            }
            # Else, increase our check count and update the last check date
            $entry.CheckCount = [int]$entry.CheckCount + 1
            $entry.LastCheckDate = ((Get-Date).ToUniversalTime()).ToString("o")
            # Then see if the unique identifier (which is the email subject) exists in the retrieved messages
            $key = $entry.UniqueIdentifier
            if ($hashEmails.ContainsKey($key)) {
                $entry.EmailReceivedDate = $hashEmails[$key]
                $entry.Status = "Success"
            }
        }
    }
    # Write the updated CSV data back to the file
    $csvData | Export-Csv -Path $script:CsvFileTrackerPath -Force
    # Release the lock
    $lock.Close()
    Write-Host "CSV tracker file updated and lock released." -ForegroundColor Green
}

#------------------------------------------------------------------------------
#-- Log start and end of the script execution with timestamps
#------------------------------------------------------------------------------
function Write-ScriptExecution
{
    param (
        [string]$Action,
        [string]$Logfile
    )
    $timeStamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    if ($Action -eq "Start") {
        Out-File -FilePath $Logfile -InputObject "[$timeStamp] Starting script: $($PSCommandPath)" -Append
    }
    elseif ($Action -eq "End") {
        Out-File -FilePath $Logfile -InputObject "[$timeStamp]   Ending script: $($PSCommandPath)" -Append
    }
    else {
        Out-File -FilePath $Logfile -InputObject "[$timeStamp] $Action $($PSCommandPath)" -Append
    }
}

#------------------------------------------------------------------------------
#-- Main
#------------------------------------------------------------------------------
Write-ScriptExecution -Action "Start" -Logfile "C:\Scripts\mailmonitor\EmailMonitor.log"
$clientId = "service-principal-id"
$tenantId = "tenant-id"
$clientSecret = Get-Secret -Name 'store-name' -Vault 'SecretStore' -AsPlainText
Connect-Graph -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret

# Get emails from our target shared mailbox, only the last 5 days to limit the results
Get-MailboxContent -RecipientEmail "smb2@somedomain2.com"
Write-Host "Retrieved $($script:ListMessages.Count) messages from the mailbox." -ForegroundColor Green
if ($script:ListMessages.Count -eq 0) {
    Write-Host "No messages found in the mailbox for the last 5 days."
    exit
}

# Update the CSV tracker file with the received date and status for each email entry
Update-CsvFileTracker
Write-ScriptExecution -Action "End" -Logfile "C:\Scripts\mailmonitor\EmailMonitor.log"
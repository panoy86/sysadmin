<#
.SYNOPSIS
    Send test emails from a shared mailbox in GEICO to PPGEICO and track the sent emails in a CSV file.
.DESCRIPTION
    This script is the first part of a three-script set that is responsible for sending "monitor-type" emails from a shared mailbox in GEICO to PPGEICO. It generates a unique identifier for each email, sends the email using the Graph API, and updates a tracker CSV file with the details of the sent email.
    It authenticates to the Graph API using client credentials and then sends an email on behalf of a specified sender.
.NOTES
    - Needs a valid shared mailbox for the sender email address, and the app registration must have the appropriate permissions (e.g., Mail.Send) granted and consented.
    - CsvFileTracker has the Status property, with only 2 possible values: "Success" or "Fail".
#>

# Script-wide variables
$script:Token = $null
$script:CsvFileTrackerPath = "C:\Scripts\mailmonitor\EmailTracker.csv"  # Use full path (relative path will fail)
$script:LockFilePath = "C:\Scripts\mailmonitor\EmailTracker.lock"       # Path for lock file to manage concurrent access to CSV tracker file

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
#-- Send mail using Graph API
#------------------------------------------------------------------------------
function Send-MailAs
{
    param (
        [string]$SenderEmail,
        [string]$RecipientEmail,
        [string]$Subject,
        [string]$Body
    )
    Write-Host "Sending as" $SenderEmail "to" $RecipientEmail
    Write-Host "Subject:" $Subject
    Write-Host "Body:" $Body
    # Construct the email payload for Graph API
    $uri = "https://graph.microsoft.com/v1.0/users/$SenderEmail/sendMail"
    $email = @{
        message = @{
            subject = $Subject
            body = @{
                contentType = "Text"
                content = $Body
            }
            toRecipients = @(
                @{
                    emailAddress = @{
                        address = $RecipientEmail
                    }
                }
            )
        }
        saveToSentItems = $true
    }
    $jsonBody = $email | ConvertTo-Json -Depth 10
    # Send the email using Graph API
    Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/json" -Body $jsonBody -Headers @{Authorization = "Bearer $script:Token"}
}

#------------------------------------------------------------------------------
#-- Update CSV tracker file with email details and status
#------------------------------------------------------------------------------
function Update-CsvFileTracker
{
    param (
        [string]$UniqueIdentifier,
        [string]$SenderEmail
    )
    # Create a new entry for our tracker CSV file with the email details and initial status
    $emailSentDate = ((Get-Date).ToUniversalTime()).ToString("o") # ISO 8601 format
    $newEntry = @{
        UniqueIdentifier = $UniqueIdentifier
        SenderEmail = $SenderEmail
        EmailSentDate = $emailSentDate
        EmailReceivedDate = ""
        LastCheckDate = ""
        CheckCount = 0  # maybe we attempt 3x to check if recipient received the email
        Status = ""     # possible values: "Success", "Fail"
        AlertCount = 0  # count of how many times we've alerted about this email (e.g., if Status -eq "Fail")
    }
    # If the tracker file doesn't exist, create a new one
    if (-not (Test-Path -Path $script:CsvFileTrackerPath)) {
        $newEntry | Export-Csv -Path $script:CsvFileTrackerPath -NoTypeInformation
        return $true
    }
    else {
        # Attempt to pseudo-lock the CSV file before updating to prevent concurrent access issues
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
    }
    # Add the new entry to the CSV file
    Write-Host "Acquired lock on CSV tracker file. Updating with new email entry."
    [array]$csvData = Import-Csv -Path $script:CsvFileTrackerPath
    $csvData += $newEntry
    $csvData | Export-Csv -Path $script:CsvFileTrackerPath -Force
    # Release the lock
    #Write-Host "test, sleeping for 90 seconds to simulate work with the locked file..."
    #Start-Sleep -Seconds 90
    $lock.Close()
    return $true
}

#------------------------------------------------------------------------------
#-- Create a random 15-character string to use as a unique identifier
#------------------------------------------------------------------------------
function New-RandomString
{
    $characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'.ToCharArray()
    $randomString = -join ($characters | Get-Random -Count 15)
    return $randomString
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

# Send from geico.com to ppgeico.com
$subject = New-RandomString
$fromAddress = "smb1@somedomain1.com"
$toAddress = "smb2@somedomain2.com"
if (Update-CsvFileTracker -UniqueIdentifier $subject -SenderEmail $fromAddress) {
    Send-MailAs -SenderEmail $fromAddress -RecipientEmail $toAddress -Subject $subject -Body "For monitoring purpose only. Please ignore. This email is sent from an automated script to test email sending and tracking functionality."
}
Write-ScriptExecution -Action "End" -Logfile "C:\Scripts\mailmonitor\EmailMonitor.log"
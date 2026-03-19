<#
.SYNOPSIS
    Send alert emails when failures are detected in the email tracking process, and optionally send a daily summary email
        with the status of all tracked emails.
.DESCRIPTION
    This script is the third part of a three-script set that is responsible for sending alert emails when failures are
        detected in the email tracking process. It reads from a CSV tracker file to identify any emails that have a
        "Fail" status and sends an alert email with the details of those failures. The script also includes functionality
        to send a daily summary email with the status of all tracked emails.
.NOTES
    Modules needed: none (using built-in PowerShell cmdlets and Invoke-WebRequest for Graph API calls)
    Requires a service princial with Microsoft Graph app-permissions of Mail.Send for the source shared mailbox in GEICO.
    The script uses a lock file mechanism to manage concurrent access to the CSV tracker file, ensuring that updates to the
        tracker are done safely without conflicts.
    The alert email includes a table with details of each failure, such as the unique identifier, sender email, sent date,
        received date, status, and alert count.
    The daily summary email provides an overview of all tracked emails and their statuses.
    CsvFileTracker has the Status property, with only 2 possible values: "Success" or "Fail".
    Has an optional daily summary email that can be sent at a specific hour (e.g., 8 AM) with the status of all tracked emails.
    Requires PowerShell 7+ to run due to the use of some newer cmdlets and features.
#>

# Script-wide variables
$script:Token = $null
$script:CsvFileTrackerPath = "C:\Scripts\mailmonitor\EmailTracker.csv"  # Use full path (relative path will fail)
$script:LockFilePath = "C:\Scripts\mailmonitor\EmailTracker.lock"       # Path for lock file to manage concurrent access to CSV tracker file
$script:MaxAlerts = 2  # Maximum number of alerts to send for each "Status -eq 'Fail'" email

#------------------------------------------------------------------------------
#-- Authenticate to Graph API using an Azure app/secret combination
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
    #Write-Host "Body:" $Body
    # Construct the email payload for Graph API
    $uri = "https://graph.microsoft.com/v1.0/users/$SenderEmail/sendMail"
    $email = @{
        message = @{
            subject = $Subject
            body = @{
                contentType = "HTML"
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
    Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/json" `
        -Body $jsonBody -Headers @{Authorization = "Bearer $script:Token"} `
        -UseBasicParsing -ErrorAction SilentlyContinue
}

#------------------------------------------------------------------------------
#-- Alert function to send an alert email when a failure is detected
#------------------------------------------------------------------------------
function Send-Alert
{
    param (
        [string]$SenderEmail,
        [string]$RecipientEmail
    )
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
    # Read the CSV tracker file and find the entry for the email that failed
    $listFailures = @()
    $csvData = Import-Csv -Path $script:CsvFileTrackerPath
    foreach ($entry in $csvData) {
        #  Init the alert count if needed
        if ($entry.AlertCount.Length -eq 0) {
            $entry.AlertCount = 0
        }
        # Check for failed entries and haven't exceeded max alert count
        # And, sent date is from the last 3 days only to avoid alerting on old failures
        if ($entry.Status -eq "Fail" -and [int]$entry.AlertCount -lt $script:MaxAlerts -and [datetime]$entry.EmailSentDate -ge (Get-Date).AddDays(-3)) {
            # Increment the alert count for this entry and update the CSV file
            $entry.AlertCount = [int]$entry.AlertCount + 1
            $listFailures += $entry
        }
    }
    # Save the updated CSV data back to the file
    $csvData | Export-Csv -Path $script:CsvFileTrackerPath -Force
    $lock.Close()
    Write-Host "CSV tracker file updated and lock released." -ForegroundColor Green
    # Send an alert email about the failures found, if any
    if ($listFailures.Count -gt 0) {
        # Construct the alert email body with details of the failures        
        $alertBody = "<html><body><table border='1' style='border-collapse: collapse; font-size:13px;'><tr><th>Unique Identifier</th><th>Sender Email</th><th>Email Sent Date</th><th>Email Received Date</th><th>Status</th><th>Alert Count</th></tr>"
        foreach ($failure in $listFailures) {
            $dateSent = Get-Date($failure.EmailSentDate) -Format "yyyy-MM-dd HH:mm:ss zzz"
            $dateReceived = if ($failure.EmailReceivedDate.Length -gt 0) { Get-Date($failure.EmailReceivedDate) -Format "yyyy-MM-dd HH:mm:ss zzz" } else { "N/A" }
            $alertBody += "<tr><td>$($failure.UniqueIdentifier)</td><td>$($failure.SenderEmail)</td><td>$dateSent</td><td>$dateReceived</td><td>$($failure.Status)</td><td>$($failure.AlertCount)</td></tr>"
        }
        $alertBody += "</table></body></html>"
        $alertBody | Out-File -FilePath ".\zz.html" -Encoding utf8
        Send-MailAs -SenderEmail $SenderEmail -RecipientEmail $RecipientEmail -Subject "Alert: Email tracking failures detected" -Body $alertBody
    }
}

#------------------------------------------------------------------------------
#-- Optional, send a daily summary email with the status of all tracked emails
#------------------------------------------------------------------------------
function Send-DailySummary
{
    param (
        [string]$SenderEmail,
        [string]$RecipientEmail,
        [int]$HourToSend
    )
    # Assume this script is scheduled to run hourly, check the current hour and only send the summary at a specific time (e.g., 8 AM)
    $currentHour = (Get-Date).Hour
    if ($currentHour -eq $HourToSend) {
        # If the tracker file doesn't exist, throw an error
        if (-not (Test-Path -Path $script:CsvFileTrackerPath)) {
            throw "Tracker file does not exist: $script:CsvFileTrackerPath"
        }
        # Read the CSV tracker file and construct the summary email body
        # Only include emails sent in the last 24 hours for the summary
        $csvData = Import-Csv -Path $script:CsvFileTrackerPath |
            Where-Object {[datetime]$_.EmailSentDate -ge (Get-Date).AddDays(-1)}
        # Add column for transit time (time between sent and received)
        $csvData | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -Name "TransitTime" -Value "N/A"}
        foreach ($data in $csvData) {
            if ($data.EmailReceivedDate.Length -gt 0 -and $data.EmailSentDate.Length -gt 0) {
                $sent = Get-Date $data.EmailSentDate
                $received = Get-Date $data.EmailReceivedDate
                $data.TransitTime = ($received - $sent).TotalMinutes.ToString("F2")
            }
        }
        # Create the table for the email body
        $foundErrors = $false
        $foundWarnings = $false
        $summaryBody = "<html><body><placeholder><table border='1' style='border-collapse: collapse; font-size:13px;'>`
            <tr><th>Unique Identifier</th>`
            <th>Sender Email</th>`
            <th>Email Sent Date</th>`
            <th>Email Received Date</th>`
            <th>Status</th>`
            <th>Transit Time (minutes)</th></tr>"
        foreach ($entry in $csvData) {
            $dateSent = Get-Date($entry.EmailSentDate) -Format "yyyy-MM-dd HH:mm:ss zzz"
            $dateReceived = if ($entry.EmailReceivedDate.Length -gt 0) { Get-Date($entry.EmailReceivedDate) `
                -Format "yyyy-MM-dd HH:mm:ss zzz" } else { "N/A" }
            # Change the color of the Status column based on Success/Fail
            $colorStatus = "Green"
            if ($entry.Status -eq "Fail") {
                $colorStatus = "Red"
                $foundErrors = $true
            }
            # Change the color of the Transit Time column if transit time is greater than 5 minutes
            $colorTransit = "Black"
            if ($entry.TransitTime -ne "N/A" -and [double]$entry.TransitTime -gt 5) {
                $colorTransit = "OrangeRed"
                $foundWarnings = $true
            }
            $summaryBody += "<tr><td>$($entry.UniqueIdentifier)</td>`
                <td>$($entry.SenderEmail)</td>`
                <td>$dateSent</td>`
                <td>$dateReceived</td>`
                <td style='color:$colorStatus'>$($entry.Status)</td>`
                <td style='color:$colorTransit' align='right'>$($entry.TransitTime)</td></tr>"
        }
        $summaryBody += "</table></body></html>"
        # Replace the placeholder in the summary body with an appropriate header based on whether errors or warnings were found
        if ($foundErrors) {
            $summaryBody = $summaryBody.Replace("<placeholder>", "<h2 style='color:Red;'>Errors found in email tracking</h2>")
        }
        elseif ($foundWarnings) {
            $summaryBody = $summaryBody.Replace("<placeholder>", "<h2 style='color:Orange;'>Warnings found in email tracking</h2>")
        }
        else {
            $summaryBody = $summaryBody.Replace("<placeholder>", "<h2 style='color:Green;'>No issues found in email tracking</h2>")
        }
        # Create a temp copy of the summary email body for debugging purposes, and send the email
        $summaryBody | Out-File -FilePath ".\zz_summary.html" -Encoding utf8
        Send-MailAs -SenderEmail $SenderEmail `
            -RecipientEmail $RecipientEmail `
            -Subject "Daily Summary: Email tracking status" `
            -Body $summaryBody
    }
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
#-- Main script
#------------------------------------------------------------------------------
Write-ScriptExecution -Action "Start" -Logfile "C:\Scripts\mailmonitor\EmailMonitor.log"
$clientId = "something" # email-exchange-readyonly-appreg
$tenantId = "something" # GEICO
$clientSecret = Get-Secret -Name ("geicosecret_" + $clientId) -Vault "SecretStore" -AsPlainText
Connect-ToGraph -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret
Send-Alert -SenderEmail "111@geico.com" -RecipientEmail "222@geico.com"
Send-DailySummary -SenderEmail "111m@geico.com" -RecipientEmail "222@geico.com" -HourToSend 14
Write-ScriptExecution -Action "End" -Logfile "C:\Scripts\mailmonitor\EmailMonitor.log"

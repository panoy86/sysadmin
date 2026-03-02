<#
.DESCRIPTION
    This script retrieves message trace information for emails sent from a specified sender to a specified recipient within a given number of days using Microsoft Graph API.
.PARAMETER SenderAdress
    The email address of the sender for whom you want to retrieve message trace information.
.PARAMETER RecipientAdress
    The email address of the recipient for whom you want to retrieve message trace information.
.PARAMETER NumDays
    The number of days to look back for message trace information, default is 2 days.
.NOTES
    Requires ExchangeMessageTrace.Read.All permission.
#>
param (
    [Parameter(Mandatory = $false, HelpMessage = "Enter the sender's email address.")]
    [string]$SenderAdress,
    [Parameter(Mandatory = $false, HelpMessage = "Enter the recipient's email address.")]
    [string]$RecipientAdress,
    [string]$NumDays = 2
)

#-- Internal variables
$Script:Token = $null

#------------------------------------------------------------------------------
#-- Get the existing session token
#------------------------------------------------------------------------------
function Get-SessionToken
{
    $request = @{
      Method = "GET"
      URI = "/v1.0/users"
      OutputType = "HttpResponseMessage"
    }
    $response = Invoke-GraphRequest @request -ErrorAction Stop
    $headers = $response.RequestMessage.Headers
    $Script:Token = $headers.Authorization.Parameter
}

#------------------------------------------------------------------------------
# This script retrieves message trace information for emails sent from a
# specified sender to a specified recipient within a given number of days.
#-------------------------------------------------------------------------------
Get-SessionToken

$baseuri = "https://graph.microsoft.com/beta/admin/exchange/tracing/messageTraces"
if ($SenderAdress -and -not $RecipientAdress) {
    $filter = "(senderAddress eq '$SenderAdress') and (receivedDateTime ge " + (Get-Date).AddDays(-$NumDays).ToString("o") + ")"
}
if (-not $SenderAdress -and $RecipientAdress) {
    $filter = "(recipientAddress eq '$RecipientAdress') and (receivedDateTime ge " + (Get-Date).AddDays(-$NumDays).ToString("o") + ")"
}
if ($SenderAdress -and $RecipientAdress) {
    $filter = "(senderAddress eq '$SenderAdress') and (recipientAddress eq '$RecipientAdress') and (receivedDateTime ge " + (Get-Date).AddDays(-$NumDays).ToString("o") + ")"
}
$uri = $baseuri + "?`$filter=" + $filter
$uri
try {
    $result = Invoke-WebRequest -Method GET -Uri $uri -ContentType "application/json" -Headers @{Authorization = $script:Token} -ErrorAction SilentlyContinue -UseBasicParsing
    $resultContent = ConvertFrom-Json $result.Content
    if ($result.StatusCode -eq 200) {
        $messageTraces = $resultContent.value
        if ($messageTraces.Count -gt 0) {
            Write-Output "Message trace information for emails sent from '$SenderAdress' to '$RecipientAdress' in the last $NumDays days:"
            foreach ($trace in $messageTraces) {
                Write-Output "Subject: $($trace.subject), Status: $($trace.status), Received: $($trace.receivedDateTime)"
            }
        }
        else {
            Write-Output "No message trace information found for the specified criteria."
        }
    }
    else {
        Write-Error "Failed to retrieve message trace information. Status code: $($result.StatusCode)"
    }
}
catch {
    Write-Error "An error occurred while retrieving message trace information: $_"
}
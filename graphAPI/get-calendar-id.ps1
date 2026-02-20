#-- First, get the existing token (assumes you have already authenticated to MS Graph)
$oRequest = @{
  Method = "GET"
  URI = "/v1.0/users"
  OutputType = "HttpResponseMessage"

}
$oResponse = Invoke-GraphRequest @oRequest
$oHeaders = $oResponse.RequestMessage.Headers
$oToken = $oHeaders.Authorization.Parameter

#-- Next, get the calendar ids for a specific unified group
#-- Run this first from EOL: Get-UnifiedGroup esteetest365 | fl *id, then use the value from "ExternalDirectoryObjectId"
#$sUri = "https://graph.microsoft.com/v1.0/groups?$filter=groupTypes/any(c:c eq 'Unified')&$select=id,displayName"
#$sUri = "https://graph.microsoft.com/v1.0/groups?$filter=displayName eq 'EsteeTest365'"
$sUri = "https://graph.microsoft.com/v1.0/groups/0875532d-3502-4139-aa65-ad0d0336bebd/calendar"
$oResult = Invoke-WebRequest -Method GET -Uri $sUri -ContentType "application/json" -Headers @{Authorization = "Bearer $oToken"} -ErrorAction Stop
Write-Host "Status code:" $oResult.StatusCode
$oContent = ConvertFrom-Json $oResult.Content
$oContent | Format-List
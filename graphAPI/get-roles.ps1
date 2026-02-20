
#-- Internal variables
$script:oToken = $null
$script:aAzureADDirectoryRoles = @()

#-- Get the existing session token
function Get-SessionToken
{
    $oRequest = @{
      Method = "GET"
      URI = "/v1.0/users"
      OutputType = "HttpResponseMessage"
    }
    $oResponse = Invoke-GraphRequest @oRequest
    $oHeaders = $oResponse.RequestMessage.Headers
    $script:oToken = $oHeaders.Authorization.Parameter
}

#-- Get all Azure directory roles
function Get-AzureADDirectoryRoles
{
    $sUri = "https://graph.microsoft.com/v1.0/directoryRoles"
    $oResult = Invoke-WebRequest -Method GET -Uri $sUri -ContentType "application/json" -Headers @{Authorization = "Bearer $script:oToken"} -ErrorAction Stop
    if ($oResult.StatusCode -eq 200)
    {
        $oContent = ConvertFrom-Json $oResult.Content
        return $oContent.value
    }
}

#-- Main
Get-SessionToken

$u = Get-MgUser -UserId a335407cs@geico.com; $u.Id
#-> "7e7a1cbe-2964-4cf1-a049-3bba7fc9e76f"
<#
2f28220d-a35d-4eb2-a671-b753dce423fe
4e026aff-0735-4856-9442-1b9816bd9a67
#>
$oHeaders = @{
    Authorization = "Bearer $script:oToken"
}
$oResult = Get-AzureADDirectoryRoles
#$sUri = "https://graph.microsoft.com/v1.0/users/" + $u.Id + "/memberOf"
## $sUri = "https://graph.microsoft.com/v1.0/users/me/appRoleAssignments"
#$oResult = Invoke-RestMethod -Uri $sUri -Headers $oHeaders -Method Get
#$oResult.value | Format-Table principalDisplayName
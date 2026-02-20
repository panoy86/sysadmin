<#
.SYNOPSIS
    Get Microsoft 365 license for a specified user using Microsoft Graph API.   
.DESCRIPTION
    This script retrieves the Microsoft 365 license for a specified user using Microsoft Graph API.
.PARAMETER UserId
    The UserId parameter accepts either the alias or the User Principal Name (UPN) of the user for whom you want to retrieve the license information.
.EXAMPLE
    .\get-user-license.ps1 -UserId "jdoe"
    This command retrieves the Microsoft 365 license for the user with the alias "jdoe".
.NOTES
    This requires that an existing/authenticated Microsoft Graph PowerShell session.
    URI used when UserId is an alias:
    https://graph.microsoft.com/v1.0/users?$filter=mailNickname eq 'franaurtan'&$select=userPrincipalName

    URI used when UserId is a UPN:
    https://graph.microsoft.com/v1.0/users/franaurtan@geico.com
#>
param (
    [string]$UserId #-- Either alias or UPN
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
    $response = Invoke-GraphRequest @request
    $headers = $response.RequestMessage.Headers
    $Script:Token = $headers.Authorization.Parameter
}

#------------------------------------------------------------------------------
#-- Get the licenses applied to the user
#------------------------------------------------------------------------------
function Get-LicensesPerUser
{
    param (
        [string]$Upn #-- UserPrincipalName of the user
    )
    $uri = "https://graph.microsoft.com/v1.0/users/" + $upn + "/licenseDetails"
    $result = Invoke-WebRequest -Method GET -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $Script:Token"} -ErrorAction Stop
    $returnValue = "notfound"
    if ($result.StatusCode -eq 200)
    {
        $content = ConvertFrom-Json $result.Content
        $skuLicenses = @()
        foreach ($license in $content.value)
        {
            $skuLicenses += $license.skuPartNumber
        }
        if ($skuLicenses -contains "ENTERPRISEPACK")
        {
            $returnValue = "E3 (Enterprise)"
        }
        elseif (($skuLicenses -contains "STANDARDPACK") -or ($skuLicenses -contains "EXCHANGEENTERPRISE"))
        {
            $returnValue = "F3 (Standard)"
        }
    }
    return $returnValue
}

#------------------------------------------------------------------------------
#-- Validate the UserId input
#------------------------------------------------------------------------------
function Test-UserId
{
    param (
        [string]$UserId
    )
    if ($null -eq $UserId -or $UserId.Trim().Length -eq 0)
    {
        Write-Host "Usage: get-user-manager-hierarchy-peers.ps1 -UserId <upn|alias>"
        exit
    }
    # Check if the input is an UPN or alias and construct the appropriate Graph API query
    $isUpn = $null
    if ($UserId -like "*@*")
    {
        $uri = "https://graph.microsoft.com/v1.0/users('" + $UserId + "')"
        $isUpn = $true
    }
    else
    {
        $uri = "https://graph.microsoft.com/v1.0/users?`$filter=mailNickname eq '$UserId'&`$select=userPrincipalName"
        $isUpn = $false
    }
    # Query Graph API to get the UPN, if the input is an alias, or validate the UPN exists
    $isUserFound = $false
    $result = $null
    $result = Invoke-WebRequest -Method GET -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $Script:Token"} -ErrorAction SilentlyContinue
        if ($null -eq $result)
    {
        #Write-Error "Failed to query Graph API for user information."
        exit
    }
    $content = ConvertFrom-Json $result.Content
    if ($result.StatusCode -eq 200)
    {
        $content = ConvertFrom-Json $result.Content
        if ($isUpn)
        {
            $returnUpn = $content.userPrincipalName
            $isUserFound = $true
        }
        else
        {
            if ($content.value.Count -eq 1)
            {
                $returnUpn = $content.Value[0].userPrincipalName
                $isUserFound = $true
            }
        }
    }
    if (-not $isUserFound)
    {
        Write-Error "Alias '$UserId' not found or is not unique."
        exit
    }
    return $returnUpn
}

#------------------------------------------------------------------------------
#-- Main script
#------------------------------------------------------------------------------
Get-SessionToken
$upn = Test-UserId -UserId $UserId

$license = Get-LicensesPerUser -Upn $upn
$returnValue = [PSCustomObject]@{
    User = $upn
    License = $license
}
return $returnValue

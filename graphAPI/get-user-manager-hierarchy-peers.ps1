<#
.SYNOPSIS
    Get the manager hierarchy and peers for a specified user in Microsoft 365.
.DESCRIPTION
    This script retrieves the manager hierarchy and peers for a specified user in Microsoft 365 using Microsoft Graph API.
.PARAMETER UserId
    The UserId parameter accepts either the alias or the User Principal Name (UPN) of the user for whom you want to retrieve the manager hierarchy and peers.
.EXAMPLE
    .\get-user-manager-hierarchy-peers.ps1 -UserId "jdoe"
    This command retrieves the manager hierarchy and peers for the user with the alias "jdoe".
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
$Script:ManagerHierarchy = @()
$Script:Peers = @()

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
#-- Get user manager hierarchy
#------------------------------------------------------------------------------
function Get-UserManagerHierarchy
{
    param (
        [string]$Upn
    )
    $uri = "https://graph.microsoft.com/v1.0/users/" + $Upn + "/manager"
    $result = Invoke-WebRequest -Method GET -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $Script:Token"} -ErrorAction Stop -UseBasicParsing
    if ($result.StatusCode -eq 200)
    {
        $content = ConvertFrom-Json $result.Content
        $newManager = [PSCustomObject]@{
            Id = $content.id
            DisplayName = $content.displayName
            UserPrincipalName = $content.userPrincipalName
            Title = $content.jobTitle
        }    
        # Recursively get the manager's manager, reached the top when UPN is the same as the current user
        if ($content.userPrincipalName -ne $Upn)
        {
            $Script:ManagerHierarchy += $newManager
            Get-UserManagerHierarchy -Upn $content.userPrincipalName
        }
    }
}

#------------------------------------------------------------------------------
#-- Get the peers for this user
#------------------------------------------------------------------------------
function Get-UserPeers
{
    param (
        [string]$ManagerUpn
    )
    $uri = "https://graph.microsoft.com/v1.0/users/" + $ManagerUpn + "/directReports"
    $result = Invoke-WebRequest -Method GET -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $Script:Token"} -ErrorAction Stop
    if ($result.StatusCode -eq 200)
    {
        $content = ConvertFrom-Json $result.Content
        foreach ($user in $content.value)
        {
            $newPeer = [PSCustomObject]@{
                Id = $user.id
                DisplayName = $user.displayName
                UserPrincipalName = $user.userPrincipalName
                Title = $user.jobTitle
            }    
            $Script:Peers += $newPeer
        }
    }
}

#------------------------------------------------------------------------------
#-- Get the licenses applied to the peers
#------------------------------------------------------------------------------
function Get-LicensesPerUser
{
    foreach ($peer in $Script:Peers)
    {
        $upn = $peer.UserPrincipalName
        $uri = "https://graph.microsoft.com/v1.0/users/" + $upn + "/licenseDetails"
        $result = Invoke-WebRequest -Method GET -Uri $uri -ContentType "application/json" -Headers @{Authorization = "Bearer $Script:Token"} -ErrorAction Stop
        if ($result.StatusCode -eq 200)
        {
            $content = ConvertFrom-Json $result.Content
            $skuLicenses = @()
            foreach ($license in $content.value)
            {
                $skuLicenses += $license.skuPartNumber
            }
            $peer | Add-Member -MemberType NoteProperty -Name Licenses -Value ($skuLicenses -join ", ")
            $peer | Add-Member -MemberType NoteProperty -Name BasicLicense -Value ''
            if ($skuLicenses -contains "ENTERPRISEPACK")
            {
                $peer.BasicLicense = "E3"
            }
            elseif (($skuLicenses -contains "STANDARDPACK") -or ($skuLicenses -contains "EXCHANGEENTERPRISE"))
            {
                $peer.BasicLicense = "F3"
            }
        }
    }
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
    Write-Host $uri -ForegroundColor Yellow
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

Write-Host "Manager hierarchy for user:" $upn -ForegroundColor Cyan
Get-UserManagerHierarchy -Upn $upn
$Script:ManagerHierarchy | Select-Object DisplayName,UserPrincipalName,Title | Format-Table

Write-Host "Peer information for user:" $upn -ForegroundColor Cyan
Get-UserPeers -ManagerUpn $Script:ManagerHierarchy[0].UserPrincipalName
Get-LicensesPerUser
$Script:Peers | Sort-Object DisplayName | Select-Object DisplayName,UserPrincipalName,Title,BasicLicense | Format-Table

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
    URI used when UserId is an alias:
    https://graph.microsoft.com/v1.0/users?$filter=mailNickname eq 'franaurtan'&$select=userPrincipalName

    URI used when UserId is a UPN:
    https://graph.microsoft.com/v1.0/users/franaurtan@geico.com
#>
param (
    [string]$UserId #-- Either alias or UPN

)

#-- Internal variables
$script:oToken = $null
$script:aManagerHierarchy = @()
$script:aPeers = @()

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

#-- Get user manager hierarchy
function Get-UserManagerHierarchy
{
    param (
        [string]$sUpn
    )
    $sUri = "https://graph.microsoft.com/v1.0/users/" + $sUpn + "/manager"
    $oResult = Invoke-WebRequest -Method GET -Uri $sUri -ContentType "application/json" -Headers @{Authorization = "Bearer $script:oToken"} -ErrorAction Stop
    if ($oResult.StatusCode -eq 200)
    {
        $oContent = ConvertFrom-Json $oResult.Content
        $oNew = [PSCustomObject]@{
            Id = $oContent.id
            DisplayName = $oContent.displayName
            UserPrincipalName = $oContent.userPrincipalName
            Title = $oContent.jobTitle
        }    
        # Recursively get the manager's manager, reached the top when UPN is the same as the current user
        if ($oContent.userPrincipalName -ne $sUpn)
        {
            $script:aManagerHierarchy += $oNew
            Get-UserManagerHierarchy -sUpn $oContent.userPrincipalName
        }
    }
}

#-- Get the peers for this user
function Get-UserPeers
{
    param (
        [string]$sManagerUpn
    )
    $sUri = "https://graph.microsoft.com/v1.0/users/" + $sManagerUpn + "/directReports"
    $oResult = Invoke-WebRequest -Method GET -Uri $sUri -ContentType "application/json" -Headers @{Authorization = "Bearer $script:oToken"} -ErrorAction Stop
    if ($oResult.StatusCode -eq 200)
    {
        $oContent = ConvertFrom-Json $oResult.Content
        foreach ($oUser in $oContent.value)
        {
            $oNew = [PSCustomObject]@{
                Id = $oUser.id
                DisplayName = $oUser.displayName
                UserPrincipalName = $oUser.userPrincipalName
                Title = $oUser.jobTitle
            }    
            $script:aPeers += $oNew
        }
    }
}

#-- Get the licenses applied to the peers
function Get-LicensesPerUser
{
    foreach ($oPeer in $script:aPeers)
    {
        $sUpn = $oPeer.UserPrincipalName
        $sUri = "https://graph.microsoft.com/v1.0/users/" + $sUpn + "/licenseDetails"
        $oResult = Invoke-WebRequest -Method GET -Uri $sUri -ContentType "application/json" -Headers @{Authorization = "Bearer $script:oToken"} -ErrorAction Stop
        if ($oResult.StatusCode -eq 200)
        {
            $oContent = ConvertFrom-Json $oResult.Content
            $aLicenses = @()
            foreach ($oLicense in $oContent.value)
            {
                #$oLicense | Format-List
                $aLicenses += $oLicense.skuPartNumber
            }
            $oPeer | Add-Member -MemberType NoteProperty -Name Licenses -Value ($aLicenses -join ", ")
            $oPeer | Add-Member -MemberType NoteProperty -Name BasicLicense -Value ''
            if ($aLicenses -contains "ENTERPRISEPACK")
            {
                $oPeer.BasicLicense = "E3"
            }
            elseif (($aLicenses -contains "STANDARDPACK") -or ($aLicenses -contains "EXCHANGEENTERPRISE"))
            {
                $oPeer.BasicLicense = "F3"
            }
        }
    }
}

#-- Validate the UserId input
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
    $bIsUpn = $null
    if ($UserId -like "*@*")
    {
        #$sUri = "https://graph.microsoft.com/v1.0/users/" + $UserId + "&`$select=userPrincipalName"
        $sUri = "https://graph.microsoft.com/v1.0/users('" + $UserId + "')"
        $bIsUpn = $true
    }
    else
    {
        $sUri = "https://graph.microsoft.com/v1.0/users?`$filter=mailNickname eq '$UserId'&`$select=userPrincipalName"
        $bIsUpn = $false
    }
    Write-Host $sUri -ForegroundColor Yellow
    # Query Graph API to get the UPN, if the input is an alias, or validate the UPN exists
    $bUserFound = $false
    $oResult = $null
    #$ProgressPreference = 'SilentlyContinue'
    $oResult = Invoke-WebRequest -Method GET -Uri $sUri -ContentType "application/json" -Headers @{Authorization = "Bearer $script:oToken"} -ErrorAction SilentlyContinue
    #$oResult = Invoke-WebRequest -Method GET -Uri $sUri -ContentType "application/json" -Headers @{Authorization = "Bearer $script:oToken"} -ProgressAction SilentlyContinue
    if ($null -eq $oResult)
    {
        #Write-Error "Failed to query Graph API for user information."
        exit
    }
    $oContent = ConvertFrom-Json $oResult.Content
    if ($oResult.StatusCode -eq 200)
    {
        $oContent = ConvertFrom-Json $oResult.Content
        if ($bIsUpn)
        {
            $sUpn = $oContent.userPrincipalName
            $bUserFound = $true
        }
        else
        {
            if ($oContent.value.Count -eq 1)
            {
                $sUpn = $oContent.Value[0].userPrincipalName
                $bUserFound = $true
            }
        }
    }
    if (-not $bUserFound)
    {
        Write-Error "Alias '$UserId' not found or is not unique."
        exit
    }
    return $sUpn
}

#-- Main code
Get-SessionToken
$sUpn = Test-UserId -UserId $UserId

Write-Host "Manager hierarchy for user:" $sUpn -ForegroundColor Cyan
Get-UserManagerHierarchy -sUpn $sUpn
$script:aManagerHierarchy | Select-Object DisplayName,UserPrincipalName,Title | Format-Table

Write-Host "Peer information for user:" $sUpn -ForegroundColor Cyan
Get-UserPeers -sManagerUpn $script:aManagerHierarchy[0].UserPrincipalName
Get-LicensesPerUser
$script:aPeers | Select-Object DisplayName,UserPrincipalName,Title,BasicLicense | Format-Table

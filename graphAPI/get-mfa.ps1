Set-Location -Path C:\Scripts\GraphAPI\tio
$PSStyle.Progress.View = "Minimal"  #-- Other value: "Classic", only works in PowerShell 7+
$ProgressPreference = "Continue"

#-- Permissions required:
#   Connect-MgGraph -Scopes "User.Read.All", "Policy.ReadWrite.AuthenticationMethod"
#-- Commented this out as I typically pre-connect a session

#-- Get all the users
$rUsers = Get-MgUser -All
Write-Host "Total users: $($rUsers.Count)"

#-- Loop through each user and get their authentication methods
$nCtr = 0
$rFinal = @()
$hAuthMethods = @{}
ForEach ($oUser in $rUsers)
{
    #-- Show progress
    $nCtr++
    Write-Progress -Activity "Getting authentication methods" -Status $oUser.DisplayName -PercentComplete ($nCtr / $rUsers.Count * 100)
    
    #-- Get the authentication methods for the user
    $authMethods = Get-MgUserAuthenticationMethod -UserId $oUser.Id -ea SilentlyContinue
    
    #-- Save the authentication method info
    if ($null -ne $authMethods)
    {
        #-- Put the authentication methods into a list
        $rTmp = @()
        $authMethods | ForEach-Object {
            $sAuthMethod = $_.AdditionalProperties["@odata.type"]
            $rTmp += ($sAuthMethod -split "microsoft.graph.")[1]  #-- Save the shortedned version
            if (-not $hAuthMethods.Contains($sAuthMethod)) {$hAuthMethods.Add($sAuthMethod, 1)} else {$hAuthMethods[$sAuthMethod]++}
        }
        #-- Save the user and their authentication methods (as a comma-separated string)
        $rFinal += [PSCustomObject]@{
            User = $oUser.UserPrincipalName
            AuthenticationMethods = $rTmp -join ", "
        }
    }
    else
    {
        $rFinal += [PSCustomObject]@{
            User = $oUser.UserPrincipalName
            AuthenticationMethods = "None"
        }
    }
}

#-- Show progress complete
Write-Progress -Activity "Getting authentication methods" -Completed

#-- Output the results
$rFinal | Sort-Object User | Format-Table -AutoSize
$hAuthMethods.GetEnumerator() | Sort-Object Name | ForEach-Object {
    Write-Host "$($_.Name): $($_.Value)"
}
$rFinal | Export-Csv -Path .\authentication-methods.csv -NoTypeInformation
Write-Host "Authentication methods saved to authentication-methods.csv"

$r = @(); (Get-RetentionCompliancePolicy | where {$_.Name -like "policy-sp-*"} | sort Name) | foreach {$_.Name; $r += Get-RetentionCompliancePolicy -Identity $_.Name -DistributionDetail}; $r.Count

$rFinal = @()
foreach ($oPolicy in $r)
{
    if ($oPolicy.SharePointLocation.Count -gt 0)
    {
        foreach ($oSharePointLocation in $oPolicy.SharePointLocation)
        {
            $oNew = New-Object PSObject -Property @{
                PolicyName = $oPolicy.Name
                SharePointName = $oSharePointLocation.DisplayName
                SharePointUrl = $oSharePointLocation.Name
            }
            $rFinal += $oNew
        }
    }
}
$rFinal.Count
$rFinal | Export-Csv .\sp-no-deletion_retention-policies.csv


#-- Get all locations under compliance
$r = Get-RetentionCompliancePolicy -DistributionDetail
$r.Count

#-- Save the OneDrive holds
$r1 = $r | where {$_.Enabled -and $_.OneDriveLocation.Count -gt 0 -and $_.Type -eq "Hold"}
$r1.Count
$r2 = @(); $r1 | foreach {$_.OneDriveLocation | foreach {$r2 += $_.Name}}
$r2.Count
$r3 = @()
foreach ($t in $r1)
{
    $sPolicy = $t.Name
    $t.OneDriveLocation | foreach {$r3 += New-Object psobject -Property @{Policy = $sPolicy;Url = $_.Name}}
}
$r3.Count
$r2 | Out-File holds-odfb.txt -Encoding ascii
$r3 | Export-Csv holds-odfb.csv -NoTypeInformation


#-- Save the SharePoint holds
$r1 = $r | where {$_.Enabled -and $_.SharePointLocation.Count -gt 0 -and $_.Type -eq "Hold"}
$r1.Count
$r2 = @(); $r1 | foreach {$_.SharePointLocation | foreach {$r2 += $_.Name}}
$r2.Count
$r3 = @()
foreach ($t in $r1)
{
    $sPolicy = $t.Name
    $t.SharePointLocation | foreach {$r3 += New-Object psobject -Property @{Policy = $sPolicy;Url = $_.Name}}
}
$r3.Count
$r2 | Out-File holds-sp.txt -Encoding ascii
$r3 | Export-Csv holds-sp.csv -NoTypeInformation


$r1 = $r | where {$_.Enabled -and $_.ExchangeLocation.Count -gt 0 -and $_.Type -eq "Hold"}
$r1.Count
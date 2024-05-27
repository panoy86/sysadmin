$sFileReport = "m365-report-spsite-usage.csv"

#-- Connect to SPO via this Entra-App
function udf_ConnectSPO
{
    $sTenantId = "57b5a800-5145-4b65-b2e8-e0e50bace0e5"
    $sAppId = "7c9c2973-02c9-4037-ba20-030a55182b05"
    $sCertThumbPrint = "af131e55dc6dffb21bacb7745165d385f4e87cb8"
    Connect-MgGraph -ClientID $sAppId -TenantId $sTenantId -CertificateThumbprint $sCertThumbPrint -NoWelcome
}

function udf_GenReport
{
    Write-Host "Generating m365 SP Site Usage Detail report..."
    Get-MgReportSharePointSiteUsageDetail -Period D180 -OutFile $sFileReport
}

function udf_AddSiteData
{
    Write-Host "Adding specific SP site data..."
    $rSites = Import-Csv $sFileReport
    $rSites | foreach {$_ | Add-Member -Type NoteProperty -Name CreatedDateTime -Value '' -Force}
    $rSites | foreach {$_ | Add-Member -Type NoteProperty -Name LastModifiedDateTime -Value '' -Force}

    #-- Look thru our list
    foreach ($oSite in $rSites)
    {
        $oTmp = $null
        $oTmp = Get-MgSite -SiteId $oSite."Site Id"
        if ($oTmp -ne $null)
        {
            $oSite."Site URL" = $oTmp.WebUrl
            $oSite.CreatedDateTime = $oTmp.CreatedDateTime
            $oSite.LastModifiedDateTime = $oTmp.LastModifiedDateTime
        }
    }
    $rSites | Export-Csv $sFileReport
}


#-- Main
udf_ConnectSPO
udf_GenReport
udf_AddSiteData


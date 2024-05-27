$sFileReport = "m365-report-spsite-usage.csv"

#-- Connect to SPO via this Entra-App
function udf_ConnectSPO
{
    #-- Add your Entra-app Id, Tenant, and cert key
    #-- Entra app would need 2 application permissions: reports.read.all & sites.read.all
    $sTenantId = "ttttt"
    $sAppId = "aaaaa"
    $sCertThumbPrint = "ppppp"
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


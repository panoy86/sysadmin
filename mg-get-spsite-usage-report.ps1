$sFileReport = "m365-report-spsite-usage.csv"

#-- Connect to SPO via this Entra-App
function udf_ConnectSPO
{
    #-- Add your Entra-app Id, Tenant, and cert key
    #-- Entra app would need 2 application permissions: reports.read.all & sites.read.all
    #-- Need a global administrator to set or confirm this setting:
    #-- In the admin center, go to the Settings > Org Settings > Services page.
    #-- Select Reports.
    #-- Uncheck the statement In all reports, display de-identified names for users, groups, and sites, and then save your changes.
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
    $nCtr = 0
    foreach ($oSite in $rSites)
    {
        $nCtr++
        Write-Progress -Activity "Processing site data..." -Status "Processing site $nCtr" -PercentComplete (($nCtr / $rSites.Count) * 100)
        $oTmp = $null
        $oTmp = Get-MgSite -SiteId $oSite."Site Id"
        if ($oTmp -ne $null)
        {
            $oSite."Site URL" = $oTmp.WebUrl
            $oSite.CreatedDateTime = $oTmp.CreatedDateTime
            $oSite.LastModifiedDateTime = $oTmp.LastModifiedDateTime
        }
    }
    Write-Progress -Activity "Processing site data..." -Completed
    $rSites | Export-Csv $sFileReport
}

#-- Main
udf_ConnectSPO
udf_GenReport
udf_AddSiteData


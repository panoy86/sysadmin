#-- Script to extract all members of a DL, and members of sub-DLS, etc...
$rDLs = @()
$rDls += "mgr-achriss-director-above"

#-- Main program, do not change
$global:hTotalMembers = @{}
$global:nTotalDLCount = 0
$global:hGroupsFound = @{}  #-- This is used to detect loops
$global:rSubDLs = @()


#-- Recursive function for DL search
function udfGet-DLMembers([String] $sDL)
{
    #-- Find the DL
    $d = $null
    $d = Get-DistributionGroup -Identity $sDL -ea SilentlyContinue
    if ($d -ne $null)
    {
        #-- Check for loop
        if ($global:hGroupsFound.ContainsKey($d.Guid.ToString()))
        {
            Write-Host ($sSpacePrefix + $d.Name.ToString() + " (loop found)") -ForegroundColor Green
        }
        else
        {
            #-- Get the members
            $global:hGroupsFound.Add($d.Guid.ToString(), 1)
            [array]$rMembers = Get-DistributionGroupMember $d.Guid.ToString() -ResultSize unlimited
            
            #-- Show DL info
            #if ($r.Count -ge 500) {Write-Host $d.DisplayName $r.Count -ForegroundColor Yellow}
            #else {Write-Host $d.DisplayName $r.Count}
            
            #-- Show progress
            Write-Host ($sSpacePrefix + $d.Name + ' ' + $rMembers.Count.ToString())

            #-- Loop thru members
            if ($rMembers.Count -gt 0)
            {
                #-- Track the sub-DLs
                $global:rSubDLs += ($rMembers.Count.ToString() + ", " + $d.Name)
                Write-Progress ($rMembers.Count.ToString() + ", " + $d.Name) -Id 2
                foreach($t in $rMembers)
                {
                    #-- Recursively call if member is another group
                    if ($t.RecipientType -like "*Group")
                    {
                        $global:nTotalDLCount++
                        udfGet-DLMembers($t.PrimarySmtpAddress.ToString())
                    }
                    #-- Add user-mailbox/mail-contact to list
                    else
                    {
                        $sEmail = $t.PrimarySmtpAddress.ToString().ToLower().Trim()
                        if (-not $hMembers.ContainsKey($sEmail)) {$hMembers.Add($sEmail, $d.Name)}
                        else {Write-Host "   duplicate:" $sEmail $hMembers.Get_Item($sEmail) -ForegroundColor Red} 
                        if (-not $global:hTotalMembers.ContainsKey($sEmail))
                        { $global:hTotalMembers.Add($sEmail, $t) }
                    }
                }
            }
        }
    }
}


#-- Loop thru list of DLs
for($i=0; $i -lt $rDLs.Count; $i++)
{
    $sDL = $rDLs[ $i ]
    $hMembers = @{}
    udfGet-DLMembers($sDL)
    Write-Host ($sDL  + " (" + $hMembers.Count + ")")
}
Write-Host "Total users:" $global:hTotalMembers.Count
Write-Host "Total DLs (duplicates allowed):" $global:nTotalDLCount
$global:hTotalMembers[0] | fl

<#-- Fix all-upper-case names
$r1 = @(); $hMembers.GetEnumerator()| foreach {$r1 += $_.Value}
$TextInfo = (Get-Culture).TextInfo
foreach($t in $r1)
{
    #$sTmp = $t.FirstName.Trim()
    $sTmp = $t.LastName.Trim()
    if ($sTmp.Length -gt 1 -and $sTmp -cmatch "^[A-Z]*$") {Write-Host $sTmp $TextInfo.ToTitleCase($sTmp.ToLower())}
    #if ($sTmp.Length -gt 1 -and $sTmp -cmatch "^[A-Z]*$") {$t.FirstName = $TextInfo.ToTitleCase($sTmp.ToLower())}
    #if ($sTmp.Length -gt 1 -and $sTmp -cmatch "^[A-Z]*$") {$t.LastName = $TextInfo.ToTitleCase($sTmp.ToLower())}
}
#>

#-- Get the SMTP list only
$rFinal = @()
$rTmp = $global:hTotalMembers.Get_Keys()
foreach($t in $rTmp)
{
    if ($global:hTotalMembers.ContainsKey($t)) { $rFinal += $global:hTotalMembers.Get_Item($t) }
}
$rTmp = @()
$rFinal | foreach {$rTmp += $_.PrimarySmtpAddress}
$rTmp | Out-File -Encoding ASCII .\emails-list.txt
Write-Host "File created: emails-list.txt"
Write-Host "Members:" $rFinal.Count
<#
$global:rSubDLs | Out-File .\Get-DLMembers_Sub-DLs.txt
Write-Host "SubDL count: Get-DLMembers_Sub-DLs.txt"
#>

<#
#-- Transfer to an array
$rFinal = @()
$rTmp = $global:hTotalMembers.Get_Keys()
foreach($t in $rTmp)
{
    if ($global:hTotalMembers.ContainsKey($t)) { $rFinal += $global:hTotalMembers.Get_Item($t) }
}
$rFinal | select SamAccountName,Name,PrimarySmtpAddress,Database,ServerName | Sort Database | Export-Csv zz.csv
#-- OR

$nCtr = 0
$rExport = @()
foreach($t in $rFinal)
{
    $nCtr++; Write-Progress -Activity '.' -Status $t.Name -PercentComplete ($nCtr/$rFinal.Count * 100)
    $oNew = New-Object System.Object
    $oNew | Add-Member –Type NoteProperty –Name SamAccountName –Value $t.SamAccountName
    $oNew | Add-Member –Type NoteProperty –Name Name –Value $t.Name
    $oNew | Add-Member –Type NoteProperty –Name PrimarySmtpAddress –Value $t.PrimarySmtpAddress.ToString()
    $oNew | Add-Member –Type NoteProperty –Name Database –Value $t.Database
    $oNew | Add-Member –Type NoteProperty –Name ServerName –Value $t.ServerName
    $sEmailAddresses = ""
    foreach($oEmail in $t.EmailAddresses)
    {
        if ($oEmail.PrefixString.ToLower() -eq "smtp") { $sEmailAddresses += ($oEmail.AddressString + "; ") }
    }
    $oNew | Add-Member –Type NoteProperty –Name EmailAddresses –Value $sEmailAddresses
    $rExport += $oNew
}
$rExport | Sort Database | Export-Csv zz.csv
#>
<#
class myDL
{
    #-- Properties
    [string] $sDLName
    [array] $rMembers
    $oDL
    $hTmp
    
    #-- Init
    myDL([string] $sDLIdentity)
    {
        $this.oDL = $null
        $this.oDL = Get-DistributionGroup $sDLIdentity -ea SilentlyContinue
        $this.hTmp = @{}
    }
    
    #-- Method, get members of DL recursively (to be called internally only)
    GetDLMembersRecursively([string] $sDLEmail)
    {
        Write-Host $sDLEmail
        $d = $null
        $d = Get-DistributionGroup -Identity $sDLEmail -ea SilentlyContinue
        if ($d -ne $null)
        {
            [array]$rTmp = Get-DistributionGroupMember $sDLEmail -Resultsize unlimited
            if ($rTmp.Count -gt 0)
            {
                foreach($oMember in $rTmp)
                {
                    if ($oMember.RecipientType -like "*Group")
                    {
                        $this.GetDLMembersRecursively($oMember.PrimarySmtpAddress.ToString())
                    }
                    else
                    {
                        $sKey = $oMember.PrimarySmtpAddress.ToString()
                        if (-not $this.hTmp.ContainsKey($sKey))
                        {
                            $this.hTmp.Add($sKey, 1)
                            $this.rMembers += $oMember
                        }
                    }
                }
            }
        }
    }
    
    #-- Method, get members of DL
    GetMembers()
    {
        $this.rMembers = @()
        $this.GetDLMembersRecursively($this.oDL.PrimarySmtpAddress.ToString())
    }
}

$cDL = [myDL]::new("dl-pp-gco")
$cDL.GetMembers()
$cDL.rMembers | select PrimarySmtpAddress | Out-File .\zz.txt
Write-Host "Found" $cDL.rMembers.Count "(zz.txt)"
Write-Host "   $cDL.GetMembers - has the members list"
#>

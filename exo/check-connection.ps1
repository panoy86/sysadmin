#------------------------------------------------------------------------------
# Checks if the current PowerShell session has an existing EXO connection
#------------------------------------------------------------------------------
function CheckEXOConnection
{
    try {
        $r = @()
        $r = Get-ConnectionInformation
    } catch {
        return $false
    }
    if ($null -eq $r -or $r.Count -eq 0) {return $false}
    $returnValue = $false
    foreach ($i in $r) {
        if ($i.Name -match "ExchangeOnline") {
            $returnValue = $true
        }
    }
    return $returnValue
}

CheckEXOConnection
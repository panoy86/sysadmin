#-- Get all mailboxes, assuming it's < 10K
Write-Host "Getting all mailboxes, please wait..." -NoNewLine
$rm = Get-Mailbox -ResultSize 10000; $rm.Count

#-- Normalize certain properties
$rm | ForEach-Object {$_ | Add-Member -Type NoteProperty -Name "EmailAddressesString" -Value ($_.EmailAddresses -join ';') -Force}

#-- Save the export
$rm | Select-Object Alias,DisplayName,Name,EmailAddressesString,Guid,LegacyExchangeDN,PrimarySmtpAddress,RecipientTypeDetails,OrganizationalUnit | Export-Csv exp-mbx.csv -NoTypeInformation

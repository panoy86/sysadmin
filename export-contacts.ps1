#-- This script will export mail-contact objects from Exchange Online to a CSV file

Write-Host "Exporting mail-contacts to CSV file..." -NoNewline
$rc = Get-MailContact -ResultSize Unlimited
Write-Host $rc.Count

#-- Normalize the email addresses
$rc | foreach {$_ | Add-Member -MemberType NoteProperty -Name EmailAddressesString -Value ($_.EmailAddresses -join ";") -Force}

#-- Export to a CSV file
$rc | select Alias,CustomAttribute1,CustomAttribute2,CustomAttribute3,CustomAttribute4,CustomAttribute5,CustomAttribute6,CustomAttribute7,CustomAttribute8,CustomAttribute9,CustomAttribute10,CustomAttribute11,CustomAttribute12,CustomAttribute13,CustomAttribute14,CustomAttribute15,DisplayName,EmailAddressesString,ExternalEmailAddress,Guid,HiddenFromAddressListsEnabled,LegacyExchangeDN,Name,PrimarySmtpAddress | Export-Csv -Path "exp-contacts.csv" -NoTypeInformation
# sysadmin
Quick and dirty PowerShell scripts to help you as a Systems Administrator

Some PowerShell scripts here may require creation of a local/self-signed certificate, example below to create one:

#-- Create cert in user certificate store with 2-year expiry
$oCert = New-SelfSignedCertificate -Subject spgraph -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(2)
#-- Export a .CER file from this
Export-Certificate -Cert $oCert -FilePath "C:\Scripts\certs\spgraph.cer"
#-- Next import this .CER onto the Entra-App, under "certificates and secrets"


function udf_SendEmailAlert($sBody)
{
    $oSecurePassword = ConvertTo-SecureString "none" -AsPlainText -Force
    $oCredential = New-Object System.Management.Automation.PSCredential ("anonymous", $oSecurePassword)
    Send-MailMessage -To "4082034622@vtext.com" -SmtpServer "atom.paypalcorp.com" -Credential $oCredential -UseSsl -Subject "Provision-Mail alert" -Port "25" -Body $sBody -From "ftan@paypal.com" -BodyAsHtml
}

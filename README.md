# sysadmin
Quick and dirty PowerShell scripts to help you as a Systems Administrator

Some PowerShell scripts here may require creation of a local/self-signed certificate, example below to create one:

#-- Create cert in user certificate store with 2-year expiry
$oCert = New-SelfSignedCertificate -Subject spgraph -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(2)
#-- Export a .CER file from this
Export-Certificate -Cert $oCert -FilePath "C:\Scripts\certs\spgraph.cer"
#-- Next import this .CER onto the Entra-App, under "certificates and secrets"

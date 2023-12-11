
# ensure script is being run in a secure directory where saved certs won't be leaked

# based off: 
# https://somerandominternetguy.com/archives/164
# https://techcommunity.microsoft.com/t5/windows-admin-center/how-do-i-create-a-new-certificate-for-windows-admin-center/m-p/1271973

## Cert Generation - CA
# Create a root certificate authority with the current computer name as the DNS Hostname
# The certificate is valid for 20 years

$myHostName = $env:COMPUTERNAME
$dateVariable = Get-Date -Format "yyyy-MM-dd"

$rootCert = New-SelfSignedCertificate `
-CertStoreLocation Cert:\CurrentUser\My `
-Subject "$dateVariable Root CA For Windows Admin Center - $myHostName CA" `
-TextExtension @("2.5.29.19={text}CA=true","2.5.29.17={text}DNS=$($myHostName)") `
-KeyUsage CertSign,CrlSign,DigitalSignature `
-NotAfter (Get-Date).AddYears(20)

# Password protect and export the root certificate authority to be imported on the target machine (client)
[String]$rootCertPath = Join-Path -Path 'cert:\CurrentUser\My\' -ChildPath "$($rootCert.Thumbprint)"
Export-Certificate -Cert $rootCertPath -FilePath "RootCA_$dateVariable_$($myHostName).crt"

## Cert Generation - Client
# Create a self signed client certificate and specify the IP Address and DNS Hostname
# Certificate is valid for 10 years
$testCert = New-SelfSignedCertificate `
-CertStoreLocation Cert:\LocalMachine\My `
-Subject "$dateVariable Windows Admin Center - $myHostName - (Self-Signed) client" `
-TextExtension @("2.5.29.17={text}DNS=$($myHostName)") `
-KeyExportPolicy Exportable `
-KeyLength 2048 `
-NotAfter (Get-Date).AddYears(10) `
-KeyUsage DigitalSignature,KeyEncipherment `
-Signer $rootCert

# Add the certificate to the certificate store and export it
[String]$testCertPath = Join-Path -Path 'cert:\LocalMachine\My\' -ChildPath "$($testCert.Thumbprint)"
Export-Certificate -Cert $testCertPath -FilePath "clientcert_$dateVariable_$($myHostName).crt"

# Optional - Export as PFX, remember to set a password if you do
# [System.Security.SecureString]$rootCertPassword = ConvertTo-SecureString -String "password" -Force -AsPlainText
# Export-PfxCertificate -Cert $testCertPath -FilePath testcert.pfx -Password $rootCertPassword


# Show Certs in Cert Stores
Get-ChildItem Cert:\CurrentUser\Root
Get-ChildItem Cert:\LocalMachine\Root
Get-ChildItem Cert:\LocalMachine\My

## Install the cert to Windows Admin Center
# Show current SSL certs
$netshCommandShow = "netsh http show sslcert"
& cmd.exe /c $netshCommandShow

# Stop WAC
Get-Service ServerManagementGateway* | Stop-Service

# Get Thumbprint of imported cert
$GetSSLThumbprint = (Get-ChildItem -Path Cert:\LocalMachine\My\ | Where-Object { $_.Subject -eq "CN=$dateVariable Windows Admin Center - $env:COMPUTERNAME - (Self-Signed) client" }).Thumbprint
Write-Host "New SSL Thumbprint: $GetSSLThumbprint"

# Get App ID of WAC
$netshOutput = & cmd.exe /c $netshCommandShow
$netshCommandGetAppID = ($netshOutput | Select-String -Pattern 'Application ID\s+:\s+(.+)').Matches.Groups[1].Value.Trim()
Write-Host "Application ID: $netshCommandGetAppID"

# Delete SSL Binding
$netshCommandDelete = "netsh http delete sslcert ipport=0.0.0.0:6516"
& cmd.exe /c $netshCommandDelete

# Recreate SSL Binding with same App ID and new SSL Thumbprint
$netshCommandAdd = "netsh http add sslcert ipport=0.0.0.0:6516 certhash=$GetSSLThumbprint appid=$netshCommandGetAppID"
& cmd.exe /c $netshCommandAdd

# Start WAC
Get-Service ServerManagementGateway* | Start-Service

# Show current SSL certs bingin
$netshCommandShow = "netsh http show sslcert"
& cmd.exe /c $netshCommandShow

# Show Certs in Cert Stores
Get-ChildItem Cert:\CurrentUser\Root
Get-ChildItem Cert:\LocalMachine\Root
Get-ChildItem Cert:\LocalMachine\My

# Import generated CA into cert store to prevent browser error
Import-Certificate -FilePath "RootCA_$env:COMPUTERNAME.crt" -CertStoreLocation cert:\LocalMachine\Root

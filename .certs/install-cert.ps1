$certPath = "C:\Users\Alex\Desktop\Projects\excel-copilot\.certs\localhost.crt"
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath)
$store = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "CurrentUser")
$store.Open("ReadWrite")
$store.Add($cert)
$store.Close()
Write-Output "Done: cert installed to Trusted Root store"

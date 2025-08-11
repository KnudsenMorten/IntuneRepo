# === Define your certificate info ===
$CertDownloadUrl = "https://xxxx.blob.core.windows.net/<blob name>/xxxxxCodeSigning_public.cer"  # Use full public or SAS URL
$LocalCertPath = "$env:TEMP\xxxxxCodeSigning_public.cer"
$ExpectedThumbprint = "xxxxxx"

function Get-InstalledCert {
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("TrustedPublisher", "CurrentUser")
    $store.Open("ReadOnly")

    $normalizedExpected = ($ExpectedThumbprint -replace '[^\da-fA-F]', '').ToLower()

    $cert = $store.Certificates | Where-Object {
        ($_.Thumbprint -replace '[^\da-fA-F]', '').ToLower() -eq $normalizedExpected
    }

    $store.Close()
    return $cert
}


# Check if already installed
if (-not (Get-InstalledCert)) {
    Write-Output "Certificate NOT present."
    Exit 1
}
else {
    Write-Output "Certificate already present."
    exit 0
}

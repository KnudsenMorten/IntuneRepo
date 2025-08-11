# === Define your certificate info ===
$CertDownloadUrl = "https://xxxx.blob.core.windows.net/intunerepo/xxxxxCodeSigning_public.cer"  # Use full public or SAS URL
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
    try {
        Invoke-WebRequest -Uri $CertDownloadUrl -OutFile $LocalCertPath -UseBasicParsing

        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.Import($LocalCertPath)

        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store("TrustedPublisher", "CurrentUser")
        $store.Open("ReadWrite")
        $store.Add($cert)
        $store.Close()

        Write-Output "Certificate installed successfully from Azure."
        exit 0
    }
    catch {
        Write-Error "Failed to download or install certificate: $_"
        exit 1
    }
}
else {
    Write-Output "Certificate already present."
    exit 0
}

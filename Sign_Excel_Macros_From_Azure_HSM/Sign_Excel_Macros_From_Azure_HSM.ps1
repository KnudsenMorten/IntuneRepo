<#
Signing of Excel Macro files using Azure Key Vault
Runs AzureSignTool via dotnet tool run (x86) to avoid ASR block on user-profile EXEs
Created by Morten Knudsen (aka.ms/morten) — Updated
#>

################################################################################
# VARIABLES
################################################################################

# Microsoft Office Subject Interface Packages (SIPs)
$tmp  = "$env:TEMP\officesips.exe"
$url  = "https://download.microsoft.com/download/f/b/4/fb46f8ca-6a6f-4cb0-b8f4-06bf3d44da48/officesips_16.0.16507.43425.exe"
$dest = "C:\Program Files\Microsoft Office SIPs"

# Signing details
$VaultUri = "https://<keyvault name>.vault.azure.net"
$CertName = "<Keyvault certificate name>"     # KV certificate object name
$TenantId = "<tenant id>"
$ClientId = "<clientid>"
$ClientSecret = "<App Secret>"
$TimeStampUrl = "http://timestamp.globalsign.com/tsa/r6advanced1"  # or your TSA of choice
$FileToSign = "<XLSM file>"

# Use x86 dotnet host explicitly (important for Office SIP compatibility)
$DotNetX86 = "C:\Program Files (x86)\dotnet\dotnet.exe"

# Local tool workspace for dotnet tool manifest (so we can use `dotnet tool run`)
$ToolWorkDir = Join-Path $env:TEMP "signing-tool-workdir"

################################################################################
# PRE-CHECKS
################################################################################

# Admin is required for regsvr32 and most installs
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) { throw "Please run PowerShell as Administrator." }

# Unblock the macro file if it came from the internet
Unblock-File -Path $FileToSign -ErrorAction SilentlyContinue

################################################################################
# STEP 1: INSTALLATION (ONE-TIME)
################################################################################

# Ensure SIPs are present and registered
New-Item -ItemType Directory -Force -Path $dest | Out-Null
Invoke-WebRequest -Uri $url -OutFile $tmp
Start-Process -FilePath $tmp -ArgumentList "/extract:`"$dest`" /quiet" -Wait
regsvr32 /s "$dest\msosip.dll"
regsvr32 /s "$dest\msosipx.dll"

# Ensure .NET SDKs (x64 + x86) are installed
winget install --id Microsoft.DotNet.SDK.8 --source winget --architecture x64 --accept-package-agreements --accept-source-agreements
winget install --id Microsoft.DotNet.SDK.8 --source winget --architecture x86 --accept-package-agreements --accept-source-agreements

# NuGet source (idempotent) – no -ErrorAction (dotnet nuget doesn't support it)
try {
  dotnet nuget add source https://api.nuget.org/v3/index.json -n nuget.org 2>$null
} catch {
  # ignore if it already exists
}

# Prepare a local tool manifest so we can run "dotnet tool run AzureSignTool"
New-Item -ItemType Directory -Force -Path $ToolWorkDir | Out-Null
Push-Location $ToolWorkDir
if (-not (Test-Path ".config\dotnet-tools.json")) {
  & $DotNetX86 new tool-manifest | Out-Null
}

# Install (or update) AzureSignTool as a *local* tool into this workdir
# Using x86 dotnet host ensures the tool runs in a 32-bit context
try {
  & $DotNetX86 tool install AzureSignTool | Out-Null
} catch {
  & $DotNetX86 tool update AzureSignTool | Out-Null
}

################################################################################
# STEP 2: SIGN FILE (3 PASSES)
################################################################################

for ($i = 1; $i -le 3; $i++) {
  Write-Host "Signing pass $i..." -ForegroundColor Cyan
  & $DotNetX86 tool run AzureSignTool -- sign `
    "$FileToSign" `
    -kvu "$VaultUri" -kvc "$CertName" `
    -kvt "$TenantId" -kvi "$ClientId" -kvs "$ClientSecret" `
    -tr "$TimeStampUrl" -td sha256
  if ($LASTEXITCODE -ne 0) {
    throw "AzureSignTool failed on pass $i with exit code $LASTEXITCODE."
  }
}

Pop-Location
Write-Host "Signing completed successfully." -ForegroundColor Green

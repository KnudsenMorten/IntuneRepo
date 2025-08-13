# Signing of Excel Macro files
# Created by Morten Knudsen (aka.ms/morten)

################################################################################
# VARIABLES
################################################################################

# Details for Microsoft Office Subject Interface Packages for Digitally Signing VBA Projects
$tmp = "$env:TEMP\officesips.exe"
$url = "https://download.microsoft.com/download/f/b/4/fb46f8ca-6a6f-4cb0-b8f4-06bf3d44da48/officesips_16.0.16507.43425.exe"
$dest = "C:\Program Files\Microsoft Office SIPs"

# Details for Signing
$VaultUri = "https://<keyvault name>.vault.azure.net"
$CertName = "<Keyvault certificate name>"     # KV certificate object name
$TenantId = "<tenant id>"
$ClientId = "<clientid>"
$ClientSecret = "<App Secret>"
$TimeStampUrl = "http://timestamp.globalsign.com/tsa/r6advanced1"  # or your TSA of choice
$FileToSign = "<XLSM file>"

################################################################################
# STEP 1: INSTALLATION (ONE-TIME TASK)
################################################################################

# Create folder C:\Program Files\Microsoft Office SIPs
MD $dest -Force

# Install Microsoft Office Subject Interface Packages for Digitally Signing VBA Projects
# https://www.microsoft.com/en-us/download/details.aspx?id=56617
# This package (version 16.0.16507.43425, published July 15, 2024) is specifically for digitally signing and verifying VBA projects in Office files, and it doesnâ€™t require Office to be installed
# Extract files to "C:\Program Files\Microsoft Office SIPs"

Invoke-WebRequest -Uri $url -OutFile $tmp

# Download Microsoft Office Subject Interface Packages for Digitally Signing VBA Projects - and 
# extact to "C:\Program Files\Microsoft Office SIPs"

New-Item -ItemType Directory -Force -Path $dest | Out-Null
Start-Process -FilePath $tmp -ArgumentList "/extract:`"$dest`" /quiet" -Wait

# Register DLLs for Microsoft Office Subject Interface Packages for Digitally Signing VBA Projects
regsvr32 /s "$dest\msosip.dll"
regsvr32 /s "$dest\msosipx.dll"

# Install x64 SDK
winget install --id Microsoft.DotNet.SDK.8 --source winget --architecture x64 --force

# Install x86 SDK
winget install --id Microsoft.DotNet.SDK.8 --source winget --architecture x86 --force

# Install NuGet
dotnet nuget add source https://api.nuget.org/v3/index.json -n nuget.org

# Install AzureSignTool into the x86 tool cache (important for macro signing)
& "C:\Program Files (x86)\dotnet\dotnet.exe" tool install --global AzureSignTool

################################################################################
# STEP 2: SIGN FILE
# Why multiple passes? The Office SIP readme/workflows expect successive signatures
# to cover older and newer VBA signature schemes so Office can validate and, 
# where applicable, show the signature as V3
################################################################################

# --- SIGN PASS 1 (legacy) ---
dotnet run -r win-x86 -- sign `
  "$FileToSign" `
  -kvu "$VaultUri" -kvc "$CertName" `
  -kvt "$TenantId" -kvi "$ClientId" -kvs "$ClientSecret" `
  -tr "$TimeStampUrl" -td sha256

# --- SIGN PASS 2 (agile) ---
dotnet run -r win-x86 -- sign `
  "$FileToSign" `
  -kvu "$VaultUri" -kvc "$CertName" `
  -kvt "$TenantId" -kvi "$ClientId" -kvs "$ClientSecret" `
  -tr "$TimeStampUrl" -td sha256

# --- SIGN PASS 3 (V3) ---
dotnet run -r win-x86 -- sign `
  "$FileToSign" `
  -kvu "$VaultUri" -kvc "$CertName" `
  -kvt "$TenantId" -kvi "$ClientId" -kvs "$ClientSecret" `
  -tr "$TimeStampUrl" -td sha256

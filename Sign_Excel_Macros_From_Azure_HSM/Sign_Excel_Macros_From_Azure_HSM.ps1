<#
.SYNOPSIS
    Sign_Excel_Macros_From_Azure_HSM - engine script in the IntuneRepo solution.

Signing of Excel Macro files using Azure Key Vault
Runs AzureSignTool via dotnet tool run (x86), because the Office SIPs are 32-bit
Created by Morten Knudsen (aka.ms/morten) - Updated

.DESCRIPTION
    Signs the VBA project inside an Office macro-enabled file with a code-signing
    certificate held in Azure Key Vault (HSM-backed).

    IMPORTANT - the two prerequisites behind almost every failure.
    The Office SIP package readme lists two requirements that are easy to miss, and
    skipping either produces the same unhelpful error:

        Signing failed with error 800403F4   (TRUST_E_SUBJECT_FORM_UNKNOWN)

    1) vbe7.dll (shipped in the SIP package) needs the Visual C++ x86 runtime.
       Without it the DLL cannot load, msosipx.dll silently declines every file,
       and WinTrust reports "subject form unknown" - even though the SIPs are
       registered correctly.

    2) vbe7.dll must be discoverable. Placing it beside the SIP DLLs is not enough,
       because that resolves against the CALLING process's search path, and
       signtool / AzureSignTool live elsewhere. Register its full path at:
           HKLM\SOFTWARE\Microsoft\VBA  ->  REG_SZ  Vbe71DllPath

    This script performs both automatically and verifies them before signing.

.NOTES
    Solution       : IntuneRepo
    File           : Sign_Excel_Macros_From_Azure_HSM.ps1
    Developed by   : Morten Knudsen, Microsoft MVP (Security, Azure, Security Copilot)
    Blog           : https://mortenknudsen.net  (alias https://aka.ms/morten)
    GitHub         : https://github.com/KnudsenMorten
    Support        : For public repos, open a GitHub Issue on that solution's repo.

    Requires       : Windows PowerShell 5.1, run elevated (regsvr32 + HKLM writes)

#>

################################################################################
# CONFIGURATION - edit these values, then run the script (elevated).
################################################################################

# --- Signing details ----------------------------------------------------------
$VaultUri     = "https://<keyvault name>.vault.azure.net"
$CertName     = "<Keyvault certificate name>"   # KV certificate object name
$TenantId     = "<tenant id>"
$ClientId     = "<clientid>"
$ClientSecret = "<App Secret>"
$TimeStampUrl = "http://timestamp.globalsign.com/tsa/r6advanced1"   # or your TSA of choice
$FileToSign   = "<XLSM file>"

# --- Signature passes ---------------------------------------------------------
# Office VBA uses three signature formats: Legacy, Agile and V3. The Windows
# crypto stack can only create one per pass, so the file is signed three times.
$SigningPasses = 3

# --- Tool versions ------------------------------------------------------------
# Keep these consistent - a tool cannot run on a runtime that is not installed.
# AzureSignTool 7.x targets .NET 10; 6.0.0 and 5.0.0 target .NET 8.
#
#   $DotNetChannel = '10.0'  ->  $ToolVersion = '7.0.1'   (current)
#   $DotNetChannel = '8.0'   ->  $ToolVersion = '6.0.0'
#
# .NET 8 goes out of support 10 Nov 2026; .NET 10 is LTS until Nov 2028.
$ToolVersion   = '7.0.1'
$DotNetChannel = '10.0'

# --- Behaviour ----------------------------------------------------------------
# $true = skip the SIP / runtime / SDK install steps. They are all idempotent,
# so leaving this $false is safe and makes the script self-healing.
$SkipInstall = $false

################################################################################
# VARIABLES
################################################################################

$ErrorActionPreference = 'Stop'

# Let a tool built for an older major run on a newer runtime if you mix versions
$env:DOTNET_ROLL_FORWARD = 'Major'

# Microsoft Office Subject Interface Packages (SIPs)
$SipTmp  = "$env:TEMP\officesips.exe"
$SipUrl  = "https://download.microsoft.com/download/f/b/4/fb46f8ca-6a6f-4cb0-b8f4-06bf3d44da48/officesips_16.0.16507.43425.exe"
$SipDest = "C:\Program Files\Microsoft Office SIPs"

# Visual C++ x86 runtime - required by vbe7.dll (see header)
$VcRedistUrl = "https://download.microsoft.com/download/C/6/D/C6D0FD4E-9E53-4897-9B91-836EBA2AACD3/vcredist_x86.exe"

$X86Root     = "C:\Program Files (x86)\dotnet"
$ToolWorkDir = Join-Path $env:TEMP "signing-tool-workdir"
$StageDir    = "C:\SignStage"

################################################################################
# HELPERS
################################################################################

function Invoke-Native {
    param([Parameter(Mandatory)][string]$Exe, [string[]]$Arguments, [switch]$IgnoreExitCode)

    # Native executables do not throw - they set $LASTEXITCODE. try/catch around
    # them does nothing, which is why failures used to pass silently.
    & $Exe @Arguments
    if (-not $IgnoreExitCode -and $LASTEXITCODE -ne 0) {
        throw "'$Exe $($Arguments -join ' ')' failed with exit code $LASTEXITCODE."
    }
    return $LASTEXITCODE
}

function Install-DotNetSdkX86 {
    param([string]$Channel)

    # NOTE: do NOT use winget here. winget matches on package ID only, so once the
    # x64 SDK is installed it reports "No available upgrade found" and never lays
    # down the x86 payload - leaving C:\Program Files (x86)\dotnet\ missing.
    Write-Host "Installing .NET $Channel SDK (x86) into $X86Root ..." -ForegroundColor Cyan

    $installer = Join-Path $env:TEMP "dotnet-install.ps1"
    Invoke-WebRequest -Uri "https://dot.net/v1/dotnet-install.ps1" -OutFile $installer -UseBasicParsing
    & $installer -Channel $Channel -Architecture x86 -InstallDir $X86Root

    if (-not (Test-Path (Join-Path $X86Root 'dotnet.exe'))) {
        throw "The .NET $Channel x86 SDK did not install. If x86 builds are no longer published for that channel, set `$DotNetChannel = '8.0' at the top of this script."
    }
}

function Set-Vbe7Path {
    param([Parameter(Mandatory)][string]$Vbe7Path)

    if (-not (Test-Path -LiteralPath $Vbe7Path)) { throw "vbe7.dll not found at $Vbe7Path" }

    # msosipx.dll is 32-bit and reads the WOW6432Node view - write both so the
    # value is found regardless of the calling process bitness.
    foreach ($key in 'HKLM:\SOFTWARE\Microsoft\VBA', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\VBA') {
        if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
        New-ItemProperty -Path $key -Name 'Vbe71DllPath' -Value $Vbe7Path -PropertyType String -Force | Out-Null
    }
    Write-Host "Registered vbe7.dll path: $Vbe7Path" -ForegroundColor DarkGray
}

function Test-Vbe7Loads {
    param([Parameter(Mandatory)][string]$Vbe7Path)

    # Must be probed in a 32-bit process. LoadLibrary returning 126 ("module not
    # found") on a file that plainly exists means a missing DEPENDENCY - the
    # Visual C++ x86 runtime.
    $src = @"
Add-Type -Name N -Namespace W -MemberDefinition @'
[DllImport("kernel32", SetLastError = true, CharSet = CharSet.Unicode)]
public static extern IntPtr LoadLibrary(string path);
'@
if ([W.N]::LoadLibrary("$Vbe7Path") -eq [IntPtr]::Zero) { "FAIL" } else { "OK" }
"@
    $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($src))
    $out = & "$env:WinDir\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -EncodedCommand $enc
    return ($out -join '') -match 'OK'
}

function Install-VcRedistX86 {
    Write-Host "Installing the Visual C++ x86 runtime (Office SIP prerequisite)..." -ForegroundColor Cyan
    $exe = Join-Path $env:TEMP "vcredist_x86.exe"
    Invoke-WebRequest -Uri $VcRedistUrl -OutFile $exe -UseBasicParsing
    Start-Process -FilePath $exe -ArgumentList "/q", "/norestart" -Wait
}

function Register-OfficeSip {
    param([Parameter(Mandatory)][string]$DllPath)

    if (-not (Test-Path -LiteralPath $DllPath)) {
        throw "SIP DLL not found: $DllPath - the officesips.exe extraction probably failed."
    }

    # Both SIPs are 32-bit, so they must be registered by the SysWOW64 regsvr32.
    # ('x' in msosipx means XML/OPC formats, not x64.) Capture the exit code -
    # regsvr32 /s hides failures completely.
    $p = Start-Process -FilePath "$env:WinDir\SysWOW64\regsvr32.exe" `
                       -ArgumentList "/s", "`"$DllPath`"" -Wait -PassThru -NoNewWindow
    if ($p.ExitCode -ne 0) {
        throw "regsvr32 failed on '$DllPath' with exit code $($p.ExitCode)."
    }
    Write-Host "Registered SIP: $DllPath" -ForegroundColor DarkGray
}

function Test-OfficeSipRegistered {
    $root = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Cryptography\OID\EncodingType 0'

    foreach ($fn in 'CryptSIPDllIsMyFileType2', 'CryptSIPDllCreateIndirectData') {
        $hit = Get-ChildItem "$root\$fn" -ErrorAction SilentlyContinue | ForEach-Object {
            (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).Dll
        } | Where-Object { $_ -match 'msosip' }
        if (-not $hit) { return $false }
    }
    return $true
}

################################################################################
# PRE-CHECKS
################################################################################

$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) { throw "Please run PowerShell as Administrator." }

if (-not (Test-Path -LiteralPath $FileToSign)) { throw "File to sign not found: $FileToSign" }
Unblock-File -Path $FileToSign -ErrorAction SilentlyContinue

# The Office SIP signs the VBA project, not the workbook. No VBA project means
# there is genuinely nothing to sign - fail here rather than at 0x800403F4.
if ([IO.Path]::GetExtension($FileToSign) -in '.xlsm', '.xlsb', '.xlam', '.xltm', '.docm', '.dotm', '.pptm', '.potm', '.ppam') {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($FileToSign)
    try   { $hasVba = $zip.Entries.FullName -match 'vbaProject\.bin' }
    finally { $zip.Dispose() }

    if (-not $hasVba) {
        throw "'$FileToSign' contains no VBA project (xl\vbaProject.bin) - there is nothing to sign."
    }
}

################################################################################
# STEP 1: INSTALLATION (IDEMPOTENT)
################################################################################

if (-not $SkipInstall) {

    # --- Office SIPs -----------------------------------------------------------
    if (-not (Test-Path "$SipDest\msosip.dll")) {
        New-Item -ItemType Directory -Force -Path $SipDest | Out-Null
        Invoke-WebRequest -Uri $SipUrl -OutFile $SipTmp -UseBasicParsing
        Start-Process -FilePath $SipTmp -ArgumentList "/extract:`"$SipDest`" /quiet" -Wait
    }

    # --- SIP prerequisites: THIS is what fixes 0x800403F4 -----------------------
    $vbe7 = "$SipDest\vbe7.dll"
    Set-Vbe7Path -Vbe7Path $vbe7

    if (-not (Test-Vbe7Loads -Vbe7Path $vbe7)) {
        Install-VcRedistX86
        if (-not (Test-Vbe7Loads -Vbe7Path $vbe7)) {
            throw "vbe7.dll still will not load in a 32-bit process. Install the Visual C++ 2015-2022 x86 redistributable and retry."
        }
    }
    Write-Host "vbe7.dll loads in a 32-bit process." -ForegroundColor DarkGray

    Register-OfficeSip -DllPath "$SipDest\msosip.dll"
    Register-OfficeSip -DllPath "$SipDest\msosipx.dll"

    # --- x86 .NET SDK ----------------------------------------------------------
    if (-not (Test-Path (Join-Path $X86Root "dotnet.exe"))) {
        Install-DotNetSdkX86 -Channel $DotNetChannel
    }
}

################################################################################
# STEP 2: VERIFY THE ENVIRONMENT
################################################################################

$DotNet = Join-Path $X86Root "dotnet.exe"
if (-not (Test-Path -LiteralPath $DotNet)) {
    throw "x86 dotnet host not found at '$DotNet'. Set `$SkipInstall = `$false so it is installed."
}
Write-Host "Using dotnet host: $DotNet" -ForegroundColor DarkGray

# Confirm the host really is x86 - a stray DOTNET_ROOT can redirect the muxer,
# and a 64-bit process will never see the 32-bit SIP registration.
$info = (& $DotNet --info 2>&1) -join "`n"
if ($info -match 'Architecture:\s*(\S+)') {
    $arch = $Matches[1]
    Write-Host "Host architecture: $arch" -ForegroundColor DarkGray
    if ($arch -ne 'x86') {
        throw "Expected an x86 dotnet host but '$DotNet' reports '$arch'. Check `$env:DOTNET_ROOT."
    }
}

if (-not (Test-OfficeSipRegistered)) {
    throw "The Office SIP is not registered. Set `$SkipInstall = `$false so the script registers it."
}
Write-Host "Office SIP is registered." -ForegroundColor DarkGray

# NOTE: `dotnet nuget list source` returns a string[]. With an array on the left,
# -notmatch returns the non-matching ELEMENTS rather than a boolean - which is
# always truthy. Join to a single string before testing.
$sources = (& $DotNet nuget list source 2>&1) -join "`n"
if ($sources -notmatch 'api\.nuget\.org/v3/index\.json') {
    Invoke-Native -Exe $DotNet -Arguments @('nuget','add','source','https://api.nuget.org/v3/index.json','-n','nuget.org') -IgnoreExitCode | Out-Null
}

################################################################################
# STEP 3: INSTALL AZURESIGNTOOL AND SIGN
################################################################################

New-Item -ItemType Directory -Force -Path $ToolWorkDir | Out-Null
Push-Location $ToolWorkDir
try {
    if (-not (Test-Path ".config\dotnet-tools.json")) {
        Invoke-Native -Exe $DotNet -Arguments @('new','tool-manifest')
    }

    # Pin the version so a future release cannot change behaviour silently
    & $DotNet tool uninstall AzureSignTool 2>&1 | Out-Null
    Invoke-Native -Exe $DotNet -Arguments @('tool','install','AzureSignTool','--version',$ToolVersion) | Out-Null

    Write-Host "AzureSignTool version in use:" -ForegroundColor DarkGray
    & $DotNet tool list | Where-Object { $_ -match 'azuresigntool' } | Write-Host

    # Sign a staging copy, then write it back. This keeps a pristine unsigned
    # original, which matters because re-signing a file that already has the
    # legacy and agile signatures only overwrites the third one - you would never
    # regenerate the full set.
    New-Item -ItemType Directory -Force -Path $StageDir | Out-Null
    $stageFile = Join-Path $StageDir ("sign" + [IO.Path]::GetExtension($FileToSign))

    $backup = "$FileToSign.unsigned"
    $source = if (Test-Path -LiteralPath $backup) { $backup } else { $FileToSign }

    Copy-Item -LiteralPath $source -Destination $stageFile -Force
    Write-Host "Staged from: $source" -ForegroundColor DarkGray

    try {
        $signArgs = @(
            'tool','run','AzureSignTool','--','sign',
            $stageFile,
            '-kvu', $VaultUri,
            '-kvc', $CertName,
            '-kvt', $TenantId,
            '-kvi', $ClientId,
            '-kvs', $ClientSecret,
            '-fd',  'sha256',
            '-tr',  $TimeStampUrl,
            '-td',  'sha256'
        )

        for ($i = 1; $i -le $SigningPasses; $i++) {
            Write-Host "Signing pass $i of $SigningPasses ..." -ForegroundColor Cyan
            & $DotNet @signArgs
            if ($LASTEXITCODE -ne 0) {
                throw "AzureSignTool failed on pass $i with exit code $LASTEXITCODE."
            }
        }

        if (-not (Test-Path -LiteralPath $backup)) {
            Copy-Item -LiteralPath $FileToSign -Destination $backup -Force
        }
        Copy-Item -LiteralPath $stageFile -Destination $FileToSign -Force

        Write-Host "Signed file written back to: $FileToSign" -ForegroundColor Green
        Write-Host "Unsigned original kept at:   $backup" -ForegroundColor DarkGray
    }
    finally {
        Remove-Item -LiteralPath $stageFile -Force -ErrorAction SilentlyContinue
    }
}
finally {
    Pop-Location
}

Write-Host "Signing completed successfully." -ForegroundColor Green

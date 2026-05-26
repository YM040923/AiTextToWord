#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$PackagePath,
    [string]$CertificatePath,
    [switch]$NoLaunch
)

$ErrorActionPreference = 'Stop'

function Write-Step {
    param([string]$Message)
    Write-Host "[AiTextToWord] $Message"
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Resolve-FirstFile {
    param([string[]]$Patterns)

    foreach ($pattern in $Patterns) {
        $file = Get-ChildItem -Path $pattern -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1

        if ($null -ne $file) {
            return $file.FullName
        }
    }

    return $null
}

$scriptDirectory = Split-Path -Parent $PSCommandPath

if ([string]::IsNullOrWhiteSpace($PackagePath)) {
    $PackagePath = Resolve-FirstFile @(
        (Join-Path $scriptDirectory 'AiTextToWord.App_*_x64.msix'),
        (Join-Path $scriptDirectory '*AiTextToWord*.msix'),
        (Join-Path $scriptDirectory '..\src\AiTextToWord.App\AppPackages\AiTextToWord.App_*_x64_Test\AiTextToWord.App_*_x64.msix')
    )
}

if ([string]::IsNullOrWhiteSpace($CertificatePath)) {
    $CertificatePath = Resolve-FirstFile @(
        (Join-Path $scriptDirectory 'AiTextToWord.App_*_x64.cer'),
        (Join-Path $scriptDirectory '*AiTextToWord*.cer'),
        (Join-Path $scriptDirectory '..\src\AiTextToWord.App\AppPackages\AiTextToWord.App_*_x64_Test\AiTextToWord.App_*_x64.cer')
    )
}

if ([string]::IsNullOrWhiteSpace($PackagePath) -or -not (Test-Path -LiteralPath $PackagePath)) {
    throw "MSIX package was not found. Put this script and AiTextToWord.App_*_x64.msix in the same folder, or pass -PackagePath."
}

if ([string]::IsNullOrWhiteSpace($CertificatePath) -or -not (Test-Path -LiteralPath $CertificatePath)) {
    throw "Certificate file was not found. Put this script and AiTextToWord.App_*_x64.cer in the same folder, or pass -CertificatePath."
}

$PackagePath = (Resolve-Path -LiteralPath $PackagePath).Path
$CertificatePath = (Resolve-Path -LiteralPath $CertificatePath).Path

if (-not (Test-IsAdministrator)) {
    Write-Step 'Administrator permission is required to trust the local test certificate. Requesting UAC...'

    $arguments = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', ('"{0}"' -f $PSCommandPath),
        '-PackagePath', ('"{0}"' -f $PackagePath),
        '-CertificatePath', ('"{0}"' -f $CertificatePath)
    )

    if ($NoLaunch) {
        $arguments += '-NoLaunch'
    }

    Start-Process -FilePath powershell.exe -Verb RunAs -ArgumentList ($arguments -join ' ') -Wait
    exit $LASTEXITCODE
}

Write-Step "Package: $PackagePath"
Write-Step "Certificate: $CertificatePath"

$certificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath)
$thumbprint = $certificate.Thumbprint

foreach ($storeName in @('Root', 'TrustedPeople')) {
    $storePath = "Cert:\LocalMachine\$storeName"
    $trusted = Get-ChildItem -Path $storePath | Where-Object { $_.Thumbprint -eq $thumbprint }

    if ($null -eq $trusted) {
        Write-Step "Importing certificate into LocalMachine\$storeName ..."
        Import-Certificate -FilePath $CertificatePath -CertStoreLocation $storePath | Out-Null
    }
    else {
        Write-Step "Certificate already exists in LocalMachine\$storeName."
    }
}

Write-Step 'Installing app...'
Add-AppxPackage -Path $PackagePath -ForceUpdateFromAnyVersion

$installed = Get-AppxPackage | Where-Object { $_.PackageFamilyName -eq 'YM040923.AiTextToWord_b7bt7fh9488q8' }
if ($null -ne $installed) {
    Write-Step "Installed version: $($installed.Version)"
}
else {
    Write-Step 'Install command returned, but this PowerShell context cannot read the package record. Search for "AI Text to Word" in Start.'
}

if (-not $NoLaunch) {
    Write-Step 'Launching app...'
    Start-Process explorer.exe 'shell:AppsFolder\YM040923.AiTextToWord_b7bt7fh9488q8!App'
}

Write-Step 'Done.'

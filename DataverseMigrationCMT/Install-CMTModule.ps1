# Instalacja modułu CMT (Configuration Migration Tool)
# Uruchom raz: .\Install-CMTModule.ps1
# Wymaga: PowerShell 5.1+, uprawnienia do Install-Module

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$libPath = Join-Path $scriptDir 'Lib'
if (Test-Path $libPath) {
    . (Join-Path $libPath 'Connect-CMT.ps1')
} else {
    # fallback gdy uruchamiane z innego katalogu
    $libPath = Join-Path (Get-Location) 'Lib'
    if (Test-Path (Join-Path $libPath 'Connect-CMT.ps1')) {
        . (Join-Path $libPath 'Connect-CMT.ps1')
    }
}

Write-Host "Sprawdzanie modulu Microsoft.Xrm.Tooling.ConfigurationMigration..."
$mod = Install-CMTModuleIfNeeded -MinimumVersion '1.0.0.88'
Write-Host "Zainstalowano: $($mod.Name) $($mod.Version)"
Write-Host ""
Write-Host "Dostepne cmdlety:"
Get-Command -Module 'Microsoft.Xrm.Tooling.ConfigurationMigration' | ForEach-Object { Write-Host "  $($_.Name)" }
Write-Host ""
Write-Host "Pomoc Export-CrmDataFile:"
Get-Help Export-CrmDataFile -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Synopsis
Get-Help Import-CrmDataFile -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Synopsis

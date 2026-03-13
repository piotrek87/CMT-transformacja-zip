# Skrypt buduje MigracjaDataverse.exe przy uzyciu modulu ps2exe.
# Wymaga: Install-Module ps2exe -Scope CurrentUser
# Wynik: MigracjaDataverse.exe w bieznym folderze. Caly folder (z Lib, Config, *.ps1) trzeba dystrybuowac razem z exe.

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
if (-not $root) { $root = (Get-Location).Path }

$launcherScript = Join-Path $root 'Launcher-Exe.ps1'
$outExe = Join-Path $root 'MigracjaDataverse.exe'

if (-not (Test-Path $launcherScript)) {
    Write-Error "Brak pliku Launcher-Exe.ps1 w $root"
}

if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "Modul ps2exe nie jest zainstalowany." -ForegroundColor Yellow
    Write-Host "Zainstaluj: Install-Module ps2exe -Scope CurrentUser" -ForegroundColor Cyan
    Write-Host "Nastepnie uruchom ponownie: .\Build-Exe.ps1" -ForegroundColor Cyan
    exit 1
}

Import-Module ps2exe -Force
Write-Host "Budowanie MigracjaDataverse.exe..." -ForegroundColor Cyan
Invoke-ps2exe -inputFile $launcherScript -outputFile $outExe -title "Migracja Dataverse" -noConsole
if (Test-Path $outExe) {
    Write-Host "Gotowe: $outExe" -ForegroundColor Green
    Write-Host "Uruchom exe z tego folderu (wymaga Lib, Config i pozostalych plikow .ps1 obok exe)." -ForegroundColor Gray
} else {
    Write-Host "Budowanie nie powiodlo sie." -ForegroundColor Red
    exit 1
}

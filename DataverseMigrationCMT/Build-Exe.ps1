# Buduje MigracjaCMT.exe przy uzyciu modulu ps2exe.
# Wymaga: Install-Module ps2exe -Scope CurrentUser
# Wynik: MigracjaCMT.exe w bieznym folderze. Caly folder (Lib, Config, Start-CMTMigration.ps1, Install-CMTModule.ps1) trzymaj obok exe.

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
if (-not $root) { $root = (Get-Location).Path }

$launcherScript = Join-Path $root 'Launcher-Exe.ps1'
$outExe = Join-Path $root 'MigracjaCMT.exe'

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
Write-Host "Budowanie MigracjaCMT.exe (okno konsoli)..." -ForegroundColor Cyan
# -noConsole $false zeby uzytkownik widzial menu i logi
Invoke-ps2exe -inputFile $launcherScript -outputFile $outExe -title "Migracja CMT Dataverse" -noConsole:$false
if (Test-Path $outExe) {
    Write-Host "Gotowe: $outExe" -ForegroundColor Green
    Write-Host "Uruchom exe z tego folderu. Obok exe musza byc: Config, Lib, Start-CMTMigration.ps1, Start-CMTMigrationMenu.ps1, Install-CMTModule.ps1." -ForegroundColor Gray
} else {
    Write-Host "Budowanie nie powiodlo sie." -ForegroundColor Red
    exit 1
}

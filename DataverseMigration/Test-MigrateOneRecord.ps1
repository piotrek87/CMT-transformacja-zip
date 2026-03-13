<#
.SYNOPSIS
    Test migracji jednego rekordu (contact) – do diagnozy problemu "nic sie nie przemigrowalo".
.DESCRIPTION
    Uruchamia Start-DataverseMigration z -EntityFilter contact -MaxRecordsPerEntity 1.
    Polaczenia: zmienne srodowiskowe DATAVERSE_SOURCE_CONNECTION / DATAVERSE_TARGET_CONNECTION
    (Config\MigrationConfig.ps1) lub DATAVERSE_MIGRATION_SRC / DATAVERSE_MIGRATION_TGT.
    Uruchom z katalogu DataverseMigration: .\Test-MigrateOneRecord.ps1
#>
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
Set-Location $root

$src = $env:DATAVERSE_SOURCE_CONNECTION
$tgt = $env:DATAVERSE_TARGET_CONNECTION
if (-not $src) { $src = $env:DATAVERSE_MIGRATION_SRC }
if (-not $tgt) { $tgt = $env:DATAVERSE_MIGRATION_TGT }
if (-not $src -or -not $tgt) {
    Write-Host "Brak connection stringow." -ForegroundColor Yellow
    Write-Host "Ustaw zmienne srodowiskowe przed uruchomieniem:" -ForegroundColor Gray
    Write-Host '  $env:DATAVERSE_SOURCE_CONNECTION = "AuthType=OAuth;Url=https://...;..."'
    Write-Host '  $env:DATAVERSE_TARGET_CONNECTION = "AuthType=OAuth;Url=https://...;..."'
    Write-Host "Albo uruchom z GUI (polaczenia sa w pamieci) i uzyj parametru -MaxRecordsPerEntity 1 w kodzie wywolujacym Start-DataverseMigration." -ForegroundColor Gray
    exit 1
}

Write-Host "Test migracji 1 rekordu (contact)..." -ForegroundColor Cyan
& (Join-Path $root 'Start-DataverseMigration.ps1') `
    -SourceConnectionString $src `
    -TargetConnectionString $tgt `
    -EntityFilter 'contact' `
    -MaxRecordsPerEntity 1 `
    -MigrationMode 'Upsert' `
    -MatchBy 'IdThenName'
Write-Host "Koniec testu." -ForegroundColor Cyan

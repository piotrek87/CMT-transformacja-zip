<#
.SYNOPSIS
    Migracja Dataverse z logowaniem interaktywnym (przegladarka).
.DESCRIPTION
    Uruchom z KONSOLI PowerShell (nie z exe). Dwa razy otworzy sie przegladarka do logowania:
    najpierw do srodowiska ZRODLA, potem do CELU. Dziala gdy logowanie login+haslo nie dziala
    (Seamless SSO, MFA, brak dostepu do AD/intranet).
.EXAMPLE
    .\Run-InteractiveMigration.ps1
    .\Run-InteractiveMigration.ps1 -WhatIf
#>
param(
    [switch] $WhatIf
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
if (-not $root) { $root = (Get-Location).Path }

Write-Host "Migracja Dataverse - logowanie interaktywne (przegladarka)" -ForegroundColor Cyan
Write-Host ""

Write-Host "Krok 1/2: Polacz ze ZRODLEM. Otworzy sie okno przegladarki - zaloguj sie do srodowiska zrodlowego." -ForegroundColor Yellow
$libPath = Join-Path $root 'Lib'
. (Join-Path $libPath 'Connect-Dataverse.ps1')
$SourceConn = Connect-DataverseEnvironment -Interactive
if (-not (Test-DataverseConnection -Connection $SourceConn)) {
    throw "Test polaczenia ze zrodlem nie powiodl sie."
}
Write-Host "Zrodlo: OK." -ForegroundColor Green
Write-Host ""

Write-Host "Krok 2/2: Polacz z CELEM. Otworzy sie okno przegladarki - zaloguj sie do srodowiska docelowego." -ForegroundColor Yellow
$TargetConn = Connect-DataverseEnvironment -Interactive
if (-not (Test-DataverseConnection -Connection $TargetConn)) {
    throw "Test polaczenia z celem nie powiodl sie."
}
Write-Host "Cel: OK." -ForegroundColor Green
Write-Host ""

Write-Host "Uruchamiam migracje..." -ForegroundColor Cyan
& (Join-Path $root 'Start-DataverseMigration.ps1') -SourceConn $SourceConn -TargetConn $TargetConn -WhatIf:$WhatIf
Write-Host "Zakonczono." -ForegroundColor Green

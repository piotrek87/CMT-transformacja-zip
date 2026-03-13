# Uruchamiany w osobnym procesie powershell.exe (standardowy host), zeby modul
# Microsoft.Xrm.Data.PowerShell zaladowal sie (wymaga minimum host version 1.0).
# Parametr: -ConfigPath (sciezka do CMTConfig.ps1)

param(
    [Parameter(Mandatory = $true)]
    [string] $ConfigPath
)
$ErrorActionPreference = 'Stop'
$scriptRoot = [System.IO.Path]::GetDirectoryName((Resolve-Path $PSCommandPath).Path)
$dmRoot = Join-Path (Split-Path -Parent $scriptRoot) 'DataverseMigration'
$dmScript = Join-Path $dmRoot 'Start-DataverseMigration.ps1'
if (-not (Test-Path $dmScript)) {
    Write-Host "Brak: $dmScript" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $ConfigPath)) {
    Write-Host "Brak configu: $ConfigPath" -ForegroundColor Red
    exit 1
}
$config = & $ConfigPath
$srcConn = $config.SourceConnectionString
$tgtConn = $config.TargetConnectionString
if ([string]::IsNullOrWhiteSpace($srcConn) -or [string]::IsNullOrWhiteSpace($tgtConn)) {
    Write-Host "Brak SourceConnectionString lub TargetConnectionString w configu." -ForegroundColor Red
    exit 1
}
$entityList = @($config.SchemaEntityIncludeOnly)
if (-not $entityList -or $entityList.Count -eq 0) {
    Write-Host "Brak SchemaEntityIncludeOnly w configu." -ForegroundColor Red
    exit 1
}
Write-Host "Test: 1 rekord na encje ($($entityList.Count) encji). Uruchamiam DataverseMigration..." -ForegroundColor Cyan
try {
    & $dmScript -SourceConnectionString $srcConn -TargetConnectionString $tgtConn -EntityIncludeOnly $entityList -MaxRecordsPerEntity 1
    Write-Host "Zakonczono." -ForegroundColor Green
    exit 0
} catch {
    Write-Host "Blad: $_" -ForegroundColor Red
    exit 1
}

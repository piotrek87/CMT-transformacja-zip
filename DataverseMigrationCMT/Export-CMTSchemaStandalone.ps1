# Uruchamiany w osobnym procesie powershell.exe (standardowy host), zeby modul
# Microsoft.Xrm.Data.PowerShell zaladowal sie bez bledu "minimum host version".
# Parametr: -ConfigPath (sciezka do CMTConfig.ps1)

param(
    [Parameter(Mandatory = $true)]
    [string] $ConfigPath
)
$ErrorActionPreference = 'Stop'
$scriptRoot = [System.IO.Path]::GetDirectoryName((Resolve-Path $PSCommandPath).Path)
$libPath = Join-Path $scriptRoot 'Lib'
. (Join-Path $libPath 'Get-CMTSchemaFromSource.ps1')
try {
    $outPath = Get-CMTSchemaFromSource -ConfigPath $ConfigPath
    $configDir = [System.IO.Path]::GetDirectoryName($ConfigPath)
    $lastSchemaFile = Join-Path $configDir 'LastGeneratedSchema.txt'
    [System.IO.File]::WriteAllText($lastSchemaFile, $outPath, [System.Text.UTF8Encoding]::new($false))
    Write-Host ('Schemat zapisany: ' + $outPath) -ForegroundColor Green
    Write-Host 'Sciezka zapisana w Config\LastGeneratedSchema.txt – przy opcjach 1 i 2 zostanie uzyty ten schemat.' -ForegroundColor Gray
    $logDir = Join-Path $scriptRoot 'Logs'
    $latestLog = Get-ChildItem -Path $logDir -Filter 'CMT_Schema_*.log' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latestLog) { Write-Host ('Log z pobierania: ' + $latestLog.FullName) -ForegroundColor Gray }
} catch {
    Write-Host ('Blad: ' + $_.Exception.Message) -ForegroundColor Red
    $logDir = Join-Path $scriptRoot 'Logs'
    $latestLog = Get-ChildItem -Path $logDir -Filter 'CMT_Schema_*.log' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latestLog) { Write-Host ('Szczegoly w logu: ' + $latestLog.FullName) -ForegroundColor Yellow }
    exit 1
}

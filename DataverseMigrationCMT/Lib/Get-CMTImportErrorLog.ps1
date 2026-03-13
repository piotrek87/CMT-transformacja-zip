# Wyszukuje ostatnie logi CMT (Configuration Migration Tool) i wypisuje wpisy z bledami.
# Uruchom po nieudanym imporcie (Stage Failed), zeby zobaczyc szczegoly w DataMigrationUtility.log / ImportDataDetail.log.
# Uzycie: .\Get-CMTImportErrorLog.ps1 [-Tail 80] [-All]

[CmdletBinding()]
param(
    [int] $Tail = 100,
    [switch] $All
)

$appData = $env:APPDATA
$localAppData = $env:LOCALAPPDATA
$searchRoots = @(
    (Join-Path $appData 'Microsoft'),
    (Join-Path $localAppData 'Microsoft'),
    (Join-Path $localAppData 'Microsoft Dynamics 365')
)
$logNames = @('DataMigrationUtility.log', 'ImportDataDetail.log', 'Login_ErrorLog.log')
$errorKeywords = @('Error', 'Failed', 'Fail', 'Exception', 'Stage', 'Invalid', 'Missing', 'Validation', 'Unable', 'blad', 'nie powiodl')

$found = @()
foreach ($root in $searchRoots) {
    if (-not (Test-Path $root)) { continue }
    Get-ChildItem -Path $root -Filter *.log -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { ((Get-Date) - $_.LastWriteTime).TotalHours -lt 48 } |
        ForEach-Object { $found += $_ }
}

# Tez katalog narzedzia CMT (np. Tools\ConfigurationMigration) – jesli skrypt w tym samym drive
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$possibleDirs = @(
    $scriptRoot,
    (Join-Path $scriptRoot '..'),
    (Join-Path $scriptRoot '..\..')
)
foreach ($d in $possibleDirs) {
    if (Test-Path $d) {
        Get-ChildItem -Path $d -Filter *.log -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match 'Migration|Import|DataMigration' -or $logNames -contains $_.Name } |
            Where-Object { ((Get-Date) - $_.LastWriteTime).TotalHours -lt 48 } |
            ForEach-Object { $found += $_ }
    }
}

# Priorytet: logi CMT (Configuration Migration, ImportDataDetail, DataMigration)
$cmtRelated = $found | Where-Object {
    $p = $_.FullName; $n = $_.Name
    $p -match 'Configuration Migration|DataMigration|ConfigurationMigration' -or
    $n -match 'ImportDataDetail|DataMigrationUtility|Login_ErrorLog'
}
$others = $found | Where-Object { $cmtRelated -notcontains $_ }
$byTime = @($cmtRelated | Sort-Object LastWriteTime -Descending) + @($others | Sort-Object LastWriteTime -Descending) | Select-Object -First 10
if ($byTime.Count -eq 0) {
    Write-Host 'Nie znaleziono logow CMT z ostatnich 48 h. Sprawdz recznie:' -ForegroundColor Yellow
    Write-Host '  - Katalog gdzie uruchomiles CRM Configuration Migration (np. Tools\ConfigurationMigration)' -ForegroundColor Gray
    Write-Host '  - %AppData%\Microsoft  lub  %LocalAppData%\Microsoft  (podkatalogi z Migration / Dynamics)' -ForegroundColor Gray
    Write-Host '  - Pliki: DataMigrationUtility.log, ImportDataDetail.log' -ForegroundColor Gray
    exit 0
}

Write-Host "Znalezione logi (najnowsze):" -ForegroundColor Cyan
foreach ($f in $byTime) {
    Write-Host "  $($f.FullName)  ($($f.LastWriteTime))" -ForegroundColor Gray
}

$latest = $byTime[0]
Write-Host ""
Write-Host "=== Ostatnie $Tail linii z: $($latest.Name) ===" -ForegroundColor Cyan
$content = Get-Content -Path $latest.FullName -Tail $Tail -Encoding UTF8 -ErrorAction SilentlyContinue
if (-not $content) {
    Write-Host "(brak danych lub blad odczytu)" -ForegroundColor Yellow
    exit 0
}

if (-not $All) {
    $content = @($content | Where-Object {
        $line = $_
        $lower = $line.ToLowerInvariant()
        foreach ($k in $errorKeywords) {
            if ($lower.Contains($k.ToLowerInvariant())) { return $true }
        }
        return $false
    })
    if ($content.Count -eq 0) {
        Write-Host "W ostatnich $Tail liniach brak wpisow z Error/Failed/Exception. Uzyj -All, zeby zobaczyc caly ogon logu." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Ostatnie 30 linii calego logu:" -ForegroundColor Gray
        $content = Get-Content -Path $latest.FullName -Tail 30 -Encoding UTF8 -ErrorAction SilentlyContinue
    }
}

$content | ForEach-Object { Write-Host $_ }

# Gdy glowny log to ImportDataDetail (tylko "Stage Failed") – odczytaj tez DataMigrationUtility z tego samego folderu
if ($latest.FullName -match 'ImportDataDetail') {
    $cmtDir = Split-Path -Parent $latest.FullName
    $dmuLogs = @(Get-ChildItem -Path $cmtDir -Filter 'DataMigrationUtility_*.log' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
    if ($dmuLogs.Count -gt 0) {
        $dmuFile = $dmuLogs[0]
        Write-Host ""
        Write-Host "=== DataMigrationUtility (ostatnie wpisy z bledami lub ogon logu): $($dmuFile.Name) ===" -ForegroundColor Cyan
        $dmuContent = Get-Content -Path $dmuFile.FullName -Tail 200 -Encoding UTF8 -ErrorAction SilentlyContinue
        if ($dmuContent) {
            $dmuErrors = @($dmuContent | Where-Object {
                $line = $_
                $lower = $line.ToLowerInvariant()
                foreach ($k in $errorKeywords) {
                    if ($lower.Contains($k.ToLowerInvariant())) { return $true }
                }
                return $false
            })
            if ($dmuErrors.Count -gt 0) {
                $dmuErrors | Select-Object -Last 60 | ForEach-Object { Write-Host $_ }
            } else {
                Write-Host "Brak linii z Error/Failed/Exception w ostatnich 200. Ostatnie 40 linii:" -ForegroundColor Gray
                $dmuContent | Select-Object -Last 40 | ForEach-Object { Write-Host $_ }
            }
        } else {
            Write-Host "  (nie udalo sie odczytac)" -ForegroundColor Gray
        }
        Write-Host ""
        Write-Host "Pelna sciezka: $($dmuFile.FullName)" -ForegroundColor DarkGray
    }
}
Write-Host ""
Write-Host "Jesli nadal nie wiesz, co powoduje Stage Failed: otworz DataMigrationUtility_*.log i szukaj Exception / Invalid / required." -ForegroundColor Gray

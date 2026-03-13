<#
.SYNOPSIS
    Interaktywne menu konsolowe do migracji Dataverse.
.DESCRIPTION
    Uruchom: .\Start-MigrationMenu.ps1
    Krok po kroku: połączenia -> lista encji -> migracja lub podgląd.
#>

$ErrorActionPreference = 'Stop'
$script:Root = $PSScriptRoot
if (-not $script:Root -and $MyInvocation.MyCommand.Path) { $script:Root = Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $script:Root) { $script:Root = (Get-Location).Path }
$script:SourceConnStr = ''
$script:TargetConnStr = ''
$script:EntityList = @()

function Read-ConnectionString {
    param([string] $Prompt)
    Write-Host $Prompt -ForegroundColor Cyan
    Write-Host "  (wklej connection string lub ścieżkę do pliku .txt z connection stringiem)" -ForegroundColor DarkGray
    $input = Read-Host
    $trimmed = $input.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) { return $null }
    if (Test-Path -LiteralPath $trimmed -PathType Leaf) {
        return [System.IO.File]::ReadAllText($trimmed).Trim()
    }
    return $trimmed
}

function Get-EntityListFromConnections {
    if ([string]::IsNullOrWhiteSpace($script:SourceConnStr) -or [string]::IsNullOrWhiteSpace($script:TargetConnStr)) {
        Write-Host "Najpierw ustaw połączenia (opcja 1)." -ForegroundColor Yellow
        return $null
    }
    Write-Host "Pobieranie listy encji (źródło + cel)..." -ForegroundColor Gray
    try {
        $list = & (Join-Path $script:Root 'Get-MigrationEntityList.ps1') -SourceConnectionString $script:SourceConnStr -TargetConnectionString $script:TargetConnStr
        return @($list)
    } catch {
        Write-Host "Błąd: $_" -ForegroundColor Red
        return $null
    }
}

function Show-MainMenu {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor DarkCyan
    Write-Host "  Migracja Dataverse - menu główne" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor DarkCyan
    Write-Host "  1. Ustaw połączenia (źródło i cel)" -ForegroundColor White
    Write-Host "  2. Pobierz listę encji do migracji" -ForegroundColor White
    Write-Host "  3. Tylko podgląd (WhatIf) - bez kopiowania" -ForegroundColor White
    Write-Host "  4. Uruchom migrację (wszystkie encje)" -ForegroundColor White
    Write-Host "  5. Uruchom migrację (wybrane encje)" -ForegroundColor White
    Write-Host "  6. Otwórz folder z logami" -ForegroundColor White
    Write-Host "  0. Wyjście" -ForegroundColor White
    Write-Host "----------------------------------------" -ForegroundColor DarkGray
}

function Invoke-Migration {
    param([string[]] $EntityFilter = @(), [switch] $WhatIf)
    $params = @{
        SourceConnectionString = $script:SourceConnStr
        TargetConnectionString = $script:TargetConnStr
        WhatIf = $WhatIf
    }
    if ($EntityFilter -and $EntityFilter.Count -gt 0) {
        $params['EntityFilter'] = $EntityFilter
    }
    & (Join-Path $script:Root 'Start-DataverseMigration.ps1') @params
}

# główna pętla
Write-Host "Narzędzie migracji Dataverse (PowerShell)" -ForegroundColor Green
Write-Host "Skrypt główny: Start-DataverseMigration.ps1" -ForegroundColor DarkGray

while ($true) {
    Show-MainMenu
    $choice = Read-Host "Wybierz opcję (0-6)"

    switch ($choice) {
        '1' {
            Write-Host ""
            $script:SourceConnStr = Read-ConnectionString -Prompt "Connection string ŚRÓDŁA (Dynamics/Dataverse):"
            if ($null -eq $script:SourceConnStr) { Write-Host "Pominięto." -ForegroundColor Yellow; break }
            $script:TargetConnStr = Read-ConnectionString -Prompt "Connection string CELU:"
            if ($null -eq $script:TargetConnStr) { Write-Host "Pominięto." -ForegroundColor Yellow; break }
            Write-Host "Połączenia zapisane w tej sesji." -ForegroundColor Green
        }
        '2' {
            $script:EntityList = Get-EntityListFromConnections
            if ($null -eq $script:EntityList -or $script:EntityList.Count -eq 0) { break }
            Write-Host "Encje do migracji ($($script:EntityList.Count)):" -ForegroundColor Green
            $i = 1
            foreach ($e in $script:EntityList) {
                $bpf = if ($e -match 'process$') { ' [BPF]' } else { '' }
                Write-Host "  $i. $e$bpf"
                $i++
            }
        }
        '3' {
            if ([string]::IsNullOrWhiteSpace($script:SourceConnStr) -or [string]::IsNullOrWhiteSpace($script:TargetConnStr)) {
                Write-Host "Najpierw ustaw połączenia (opcja 1)." -ForegroundColor Yellow
                break
            }
            Write-Host "Uruchamiam WhatIf (podgląd bez migracji)..." -ForegroundColor Cyan
            Invoke-Migration -WhatIf
        }
        '4' {
            if ([string]::IsNullOrWhiteSpace($script:SourceConnStr) -or [string]::IsNullOrWhiteSpace($script:TargetConnStr)) {
                Write-Host "Najpierw ustaw połączenia (opcja 1)." -ForegroundColor Yellow
                break
            }
            Write-Host "Uruchamiam migrację WSZYSTKICH wspólnych encji..." -ForegroundColor Cyan
            $confirm = Read-Host "Kontynuować? (T/N)"
            if ($confirm -match '^[TtYy]') { Invoke-Migration } else { Write-Host "Anulowano." -ForegroundColor Yellow }
        }
        '5' {
            if ([string]::IsNullOrWhiteSpace($script:SourceConnStr) -or [string]::IsNullOrWhiteSpace($script:TargetConnStr)) {
                Write-Host "Najpierw ustaw połączenia (opcja 1)." -ForegroundColor Yellow
                break
            }
            if ($script:EntityList.Count -eq 0) {
                $script:EntityList = Get-EntityListFromConnections
                if ($null -eq $script:EntityList -or $script:EntityList.Count -eq 0) { break }
            }
            Write-Host "Podaj numery encji oddzielone przecinkami (np. 3,5,7) lub nazwy (np. account,contact,opportunity):" -ForegroundColor Cyan
            $userInput = Read-Host
            $selected = @()
            if (-not [string]::IsNullOrWhiteSpace($userInput)) {
                $parts = $userInput -split '[,;\s]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                foreach ($p in $parts) {
                    if ($p -match '^\d+$') {
                        $idx = [int]$p
                        if ($idx -ge 1 -and $idx -le $script:EntityList.Count) {
                            $selected += $script:EntityList[$idx - 1]
                        }
                    } else {
                        if ($script:EntityList -contains $p) { $selected += $p }
                    }
                }
            }
            if ($selected.Count -eq 0) {
                Write-Host "Brak wybranych encji. Uruchamiam migrację wszystkich." -ForegroundColor Yellow
                $selected = $null
            }
            Write-Host "Uruchamiam migrację..." -ForegroundColor Cyan
            $confirm = Read-Host "Kontynuować? (T/N)"
            if ($confirm -match '^[TtYy]') {
                Invoke-Migration -EntityFilter $selected
            } else {
                Write-Host "Anulowano." -ForegroundColor Yellow
            }
        }
        '6' {
            $logDir = Join-Path $script:Root 'Logs'
            if (Test-Path $logDir) {
                Start-Process explorer.exe -ArgumentList $logDir
            } else {
                Write-Host "Folder Logs nie istnieje (brak jeszcze uruchomień migracji)." -ForegroundColor Yellow
            }
        }
        '0' {
            Write-Host "Do widzenia." -ForegroundColor Green
            exit 0
        }
        default {
            Write-Host "Nieznana opcja. Wybierz 0-6." -ForegroundColor Yellow
        }
    }
}

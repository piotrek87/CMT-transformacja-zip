# Menu: tylko zipy z CMT + transformacja (ownerzy, daty, pola nie z celu)
# Uruchom: MigracjaCMT.bat lub MigracjaCMT.exe (Build-Exe.ps1)

$ErrorActionPreference = 'Stop'
try {
if ($MyInvocation.MyCommand.Path) {
    try { $scriptRoot = [System.IO.Path]::GetDirectoryName((Resolve-Path $MyInvocation.MyCommand.Path).Path) }
    catch { $scriptRoot = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path) }
} else {
    $scriptRoot = [System.IO.Path]::GetDirectoryName([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
}
if (-not $scriptRoot -or -not (Test-Path $scriptRoot)) { $scriptRoot = (Get-Location).Path }

$logDir = Join-Path $scriptRoot 'Logs'
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$sessionLogFile = Join-Path $logDir ('CMT_Menu_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))

function Add-MenuLog {
    param([string]$Message, [switch]$NoConsole)
    $line = "{0:yyyy-MM-dd HH:mm:ss} {1}" -f (Get-Date), $Message
    try { Add-Content -Path $sessionLogFile -Value $line -ErrorAction SilentlyContinue } catch { }
    if (-not $NoConsole) { Write-Host $Message }
}

Add-MenuLog -NoConsole "Start menu CMT. scriptRoot=$scriptRoot"

$configPath = Join-Path $scriptRoot 'Config\CMTConfig.ps1'
$outputDir = Join-Path $scriptRoot 'Output'
$inputDir = Join-Path $scriptRoot 'Input'
if (Test-Path $configPath) {
    try {
        $cfg = & $configPath
        if ($cfg.ExportOutputDirectory) { $outputDir = $cfg.ExportOutputDirectory }
    } catch { Add-MenuLog "Blad configu: $_" }
}
if (-not (Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }
if (-not (Test-Path $inputDir)) { New-Item -ItemType Directory -Path $inputDir -Force | Out-Null }

$script:LastZipPath = $null
$script:IdMapPath = Join-Path $outputDir 'CMT_IdMap_SystemUser.json'

function Show-Menu {
    Clear-Host
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  CMT - transformacja zip (ownerzy, daty, pola)" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Eksportujesz w CMT (Export data), zapisujesz zip. Tutaj: User Map + transformacja zipa." -ForegroundColor Yellow
    Write-Host "  Gotowy zip (Output\*_ForTarget.zip) importujesz sam w CMT lub innym narzedziu." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  1. Wybierz zip z CMT (lub wrzuc zipy do folderu Input)" -ForegroundColor White
    Write-Host "  2. Generuj User Map (imie i nazwisko zrodlo->cel) -> User Map XML i IdMap JSON" -ForegroundColor White
    Write-Host "  3. Transformuj zip: ownerzy, daty utworzenia, usun pola nie z celu -> *_ForTarget.zip" -ForegroundColor White
    Write-Host "  4. Pokaz ostatni log bledow CMT (gdy import konczy sie Stage Failed)" -ForegroundColor White
    Write-Host "  0. Wyjscie" -ForegroundColor Gray
    Write-Host ""
    # Status: wybrany zip i data User Map
    if ($script:LastZipPath -and (Test-Path $script:LastZipPath -PathType Leaf)) {
        $zipName = [System.IO.Path]::GetFileName($script:LastZipPath)
        Write-Host "  Aktualny zip: " -NoNewline
        Write-Host $zipName -ForegroundColor Cyan
        Write-Host "    $($script:LastZipPath)" -ForegroundColor DarkGray
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($script:LastZipPath)
        $forTargetPath = Join-Path $outputDir ($baseName + '_ForTarget.zip')
        Write-Host "  Do importu w CMT: " -NoNewline -ForegroundColor Gray
        Write-Host ($baseName + '_ForTarget.zip') -ForegroundColor Yellow
        Write-Host "    (pelna sciezka: $forTargetPath)" -ForegroundColor DarkGray
    } else {
        Write-Host "  Aktualny zip: (nie wybrano)" -ForegroundColor DarkGray
    }
    $idMapFile = Join-Path $outputDir 'CMT_IdMap_SystemUser.json'
    $idMapByDn = Join-Path $outputDir 'CMT_IdMap_ByDisplayName.json'
    $userMapDate = $null
    if (Test-Path $idMapFile -PathType Leaf) { $userMapDate = (Get-Item $idMapFile).LastWriteTime }
    if (Test-Path $idMapByDn -PathType Leaf) {
        $tDn = (Get-Item $idMapByDn).LastWriteTime
        if (-not $userMapDate -or $tDn -gt $userMapDate) { $userMapDate = $tDn }
    }
    if ($userMapDate) {
        Write-Host "  User Map: wygenerowany " -NoNewline
        Write-Host $userMapDate.ToString('dd.MM.yyyy HH:mm') -ForegroundColor Cyan
    } else {
        Write-Host "  User Map: brak (uruchom opcje 2)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

do {
    Show-Menu
    $choice = Read-Host "Wybierz opcje (0-4)"
    Add-MenuLog -NoConsole ("Wybor: " + $choice)
    switch ($choice) {
        '1' {
            Write-Host "Zip z CMT: wybierz numer z listy lub podaj sciezke." -ForegroundColor Gray
            $inputZips = @(Get-ChildItem -Path $inputDir -Filter *.zip -File -ErrorAction SilentlyContinue | Sort-Object Name)
            if ($inputZips.Count -gt 0) {
                Write-Host "Zipy w Input ($inputDir):" -ForegroundColor Cyan
                for ($i = 0; $i -lt $inputZips.Count; $i++) {
                    Write-Host "  $($i + 1). $($inputZips[$i].Name)" -ForegroundColor White
                }
                $prompt = "Numer (1-$($inputZips.Count)) lub sciezke"
                if ($inputZips.Count -ge 1) { $prompt += " [Enter = 1]" }
                $prompt += ": "
            } else {
                Write-Host "Brak zipow w Input. Podaj pelna sciezke do pliku." -ForegroundColor Yellow
                $prompt = "Sciezka do zip (lub Enter = Output\CMT_Export.zip): "
            }
            $p = (Read-Host $prompt).Trim()
            if ([string]::IsNullOrWhiteSpace($p)) {
                if ($inputZips.Count -gt 0) { $script:LastZipPath = $inputZips[0].FullName }
                else { $script:LastZipPath = Join-Path $outputDir 'CMT_Export.zip' }
            } elseif ($inputZips.Count -gt 0 -and $p -match '^\d+$') {
                $idx = [int]$p
                if ($idx -ge 1 -and $idx -le $inputZips.Count) {
                    $script:LastZipPath = $inputZips[$idx - 1].FullName
                } else {
                    $script:LastZipPath = $p
                }
            } else {
                $script:LastZipPath = $p
                if (-not [System.IO.Path]::IsPathRooted($p) -and ($p -notmatch '^[A-Za-z]:')) {
                    $inInput = Join-Path $inputDir $p
                    $inOutput = Join-Path $outputDir $p
                    if (Test-Path $inInput -PathType Leaf) { $script:LastZipPath = $inInput }
                    elseif (Test-Path $inOutput -PathType Leaf) { $script:LastZipPath = $inOutput }
                }
            }
            if (Test-Path $script:LastZipPath -PathType Leaf) {
                Write-Host "Wybrany zip: $script:LastZipPath" -ForegroundColor Green
            } else {
                Write-Host "Plik nie istnieje. Wrzuc zip(y) do $inputDir lub podaj pelna sciezke przy opcji 3." -ForegroundColor Yellow
            }
            pause
        }
        '2' {
            Add-MenuLog -NoConsole "Opcja 2: Generuj User Map"
            $mapScript = Join-Path $scriptRoot 'Lib\New-CMTUserMapByDisplayName.ps1'
            if (-not (Test-Path $mapScript)) {
                Write-Host "Brak Lib\New-CMTUserMapByDisplayName.ps1" -ForegroundColor Red
                pause
                break
            }
            if (-not (Test-Path $configPath)) {
                Write-Host "Brak Config\CMTConfig.ps1 (Polaczenia.txt: Zrodlo/Cel)." -ForegroundColor Red
                pause
                break
            }
            try {
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = 'powershell.exe'
                $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$mapScript`" -ConfigPath `"$configPath`""
                $psi.WorkingDirectory = $scriptRoot
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $false
                $p = [System.Diagnostics.Process]::Start($psi)
                $p.WaitForExit()
                if ($p.ExitCode -eq 0) {
                    Write-Host "User Map i IdMap zapisane w Output. Uzyj opcji 3 do transformacji zip." -ForegroundColor Green
                } else { Write-Host "Kod wyjscia: $($p.ExitCode)" -ForegroundColor Yellow }
            } catch {
                Add-MenuLog ('BLAD opcja 2: ' + $_)
                Write-Host ('BLAD: ' + $_) -ForegroundColor Red
            }
            pause
        }
        '3' {
            Add-MenuLog -NoConsole "Opcja 3: Transformuj zip"
            $transformScript = Join-Path $scriptRoot 'Lib\Transform-CMTZip.ps1'
            if (-not (Test-Path $transformScript)) {
                Write-Host "Brak Lib\Transform-CMTZip.ps1" -ForegroundColor Red
                pause
                break
            }
            $zipPath = $script:LastZipPath
            if (-not $zipPath -or -not (Test-Path $zipPath)) {
                $inputZips = @(Get-ChildItem -Path $inputDir -Filter *.zip -File -ErrorAction SilentlyContinue)
                if ($inputZips.Count -gt 0) {
                    $zipPath = $inputZips[0].FullName
                    $script:LastZipPath = $zipPath
                    Write-Host "Uzyto pierwszego zip z Input: $($inputZips[0].Name)" -ForegroundColor Gray
                } else {
                    $zipPath = Read-Host "Podaj sciezke do zip z CMT"
                    if ([string]::IsNullOrWhiteSpace($zipPath)) {
                        Write-Host "Brak zip. Wrzuc zip do Input lub wybierz opcje 1." -ForegroundColor Yellow
                        pause
                        break
                    }
                }
            }
            if (-not (Test-Path $zipPath)) {
                Write-Host "Plik nie istnieje: $zipPath" -ForegroundColor Red
                pause
                break
            }
            $idMap = $script:IdMapPath
            if (-not (Test-Path $idMap)) { $idMap = '' }
            $outZip = Join-Path $outputDir ([System.IO.Path]::GetFileNameWithoutExtension($zipPath) + '_ForTarget.zip')
            try {
                # Uruchom w osobnym procesie PowerShell – modul Microsoft.Xrm.Data.PowerShell nie laduje sie w hostach Cursor/VS Code
                $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$transformScript`"", '-InputZipPath', "`"$zipPath`"", '-OutputZipPath', "`"$outZip`"")
                if ($idMap) { $argList += '-IdMapPath'; $argList += "`"$idMap`"" }
                if (Test-Path $configPath) { $argList += '-ConfigPath'; $argList += "`"$configPath`"" }
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = 'powershell.exe'
                $psi.Arguments = $argList -join ' '
                $psi.WorkingDirectory = $scriptRoot
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $false
                $p = [System.Diagnostics.Process]::Start($psi)
                $p.WaitForExit()
                if ($p.ExitCode -eq 0) {
                    Write-Host "Gotowy zip: $outZip" -ForegroundColor Green
                    Write-Host "W zipie: ownerzy podmienieni (IdMap), overriddencreatedon = oryginalna data." -ForegroundColor Cyan
                    Write-Host ""
                    Write-Host "W CMT (Import data) w polu Plik ZIP wybierz DOKLADNIE ten plik:" -ForegroundColor Yellow
                    Write-Host "  $outZip" -ForegroundColor White
                    Write-Host 'NIE wybieraj pliku z folderu Input (contact.zip / contact_account.zip) - to powoduje Stage Failed.' -ForegroundColor Yellow
                    Write-Host ""
                    Write-Host "Ten zip mozesz zaimportowac w CRM Configuration Migration (Import data)." -ForegroundColor Gray
                } else { Write-Host "Kod wyjscia: $($p.ExitCode)" -ForegroundColor Yellow }
            } catch {
                Add-MenuLog ('BLAD opcja 3: ' + $_)
                Write-Host ('BLAD: ' + $_) -ForegroundColor Red
            }
            pause
        }
        '4' {
            Add-MenuLog -NoConsole "Opcja 4: Log bledow CMT"
            $logScript = Join-Path $scriptRoot 'Lib\Get-CMTImportErrorLog.ps1'
            if (-not (Test-Path $logScript)) {
                Write-Host 'Brak Lib\Get-CMTImportErrorLog.ps1' -ForegroundColor Red
                pause
                break
            }
            try {
                & $logScript -Tail 120
                Write-Host ""
                Write-Host 'Jesli log jest pusty: sprawdz katalog, w ktorym uruchomiles CRM Configuration Migration (np. plik DataMigrationUtility.log).' -ForegroundColor Gray
            } catch {
                Write-Host ('BLAD: ' + $_) -ForegroundColor Red
            }
            pause
        }
        '0' { Add-MenuLog -NoConsole 'Wyjscie'; exit 0 }
        default { Write-Host 'Wybierz 0, 1, 2, 3 lub 4.' -ForegroundColor Yellow; Start-Sleep -Seconds 2 }
    }
} while ($true)
Add-MenuLog -NoConsole 'Koniec sesji'
} catch {
    Write-Host ""
    Write-Host ('BLAD: ' + $_.Exception.Message) -ForegroundColor Red
    if ($_.ScriptStackTrace) { Write-Host $_.ScriptStackTrace -ForegroundColor Gray }
    Read-Host 'Nacisnij Enter'
    exit 1
}

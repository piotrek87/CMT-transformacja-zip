# Menu: tylko zipy z CMT + transformacja (ownerzy, daty, pola nie z celu)
# Uruchom: MigracjaCMT.bat lub MigracjaCMT.exe (Build-Exe.ps1)
# Zasada: w stringach tylko ASCII - cudzyslowy " i ', myslnik - (nie en-dash). Zapobiega bledom parsera.

$script:CMTMenuVersion = 'v1.4-metadata-cache'
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

function Get-EntitiesFromZip {
    param([string]$ZipPath)
    if (-not $ZipPath -or -not (Test-Path $ZipPath -PathType Leaf)) { return @() }
    $entitySet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
        $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
        try {
            foreach ($entry in $zip.Entries) {
                if (-not $entry.Name -or -not $entry.Name.EndsWith('.xml', [StringComparison]::OrdinalIgnoreCase)) { continue }
                if ($entry.Length -gt 50 * 1024 * 1024) { continue }
                $reader = $null
                try {
                    $stream = $entry.Open()
                    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
                    $content = $reader.ReadToEnd()
                } finally { if ($reader) { $reader.Dispose() }; if ($stream) { $stream.Dispose() } }
                $matches = [regex]::Matches($content, '(?i)<(?:entity|Entity)\s[^>]*\b(?:name|Name)=["'']([^"'']+)["'']')
                foreach ($m in $matches) {
                    if ($m.Groups[1].Success -and $m.Groups[1].Value) {
                        [void]$entitySet.Add($m.Groups[1].Value.Trim())
                    }
                }
            }
        } finally { $zip.Dispose() }
    } catch { return @() }
    $arr = @($entitySet | Sort-Object)
    return $arr
}

function Show-Menu {
    Clear-Host
    Write-Host "  Wersja: $script:CMTMenuVersion  (git: git checkout $script:CMTMenuVersion)" -ForegroundColor DarkCyan
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
    Write-Host "  5. Pobierz metadane celu (cache dla opcji 3, encje z wybranego zipa)" -ForegroundColor White
    Write-Host "  6. Popraw daty Uwag ze zrodla (uruchom PO opcji 3 na zipie *_ForTarget; wynik od razu do importu)" -ForegroundColor White
    Write-Host "  7. Generuj IdMap encji (leady, szanse, kontakty, konta - zrodlo->cel; opcja 3 moze je uzyc)" -ForegroundColor White
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
    $leadIdMapFile = Join-Path $outputDir 'CMT_IdMap_Lead.json'
    if (Test-Path $leadIdMapFile -PathType Leaf) {
        $leadMapDate = (Get-Item $leadIdMapFile).LastWriteTime
        Write-Host "  Lead IdMap: " -NoNewline
        Write-Host $leadMapDate.ToString('dd.MM.yyyy HH:mm') -ForegroundColor Cyan
    } else {
        Write-Host "  Lead IdMap: brak (opcja 7)" -ForegroundColor DarkGray
    }
    $otherMaps = @('CMT_IdMap_Opportunity.json', 'CMT_IdMap_Contact.json', 'CMT_IdMap_Account.json')
    $otherLabels = @('Szanse', 'Kontakty', 'Konta')
    for ($i = 0; $i -lt $otherMaps.Count; $i++) {
        $fp = Join-Path $outputDir $otherMaps[$i]
        if (Test-Path $fp -PathType Leaf) {
            $dt = (Get-Item $fp).LastWriteTime
            Write-Host "  $($otherLabels[$i]) IdMap: " -NoNewline
            Write-Host $dt.ToString('dd.MM.yyyy HH:mm') -ForegroundColor Cyan
        } else {
            Write-Host "  $($otherLabels[$i]) IdMap: brak (opcja 7)" -ForegroundColor DarkGray
        }
    }
    $cacheFiles = @(Get-ChildItem -Path $outputDir -Filter 'TargetMetadata_*.json' -File -ErrorAction SilentlyContinue)
    if ($cacheFiles.Count -gt 0) {
        $newestCache = $cacheFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $cacheDateStr = ''
        try {
            $cacheJson = [System.IO.File]::ReadAllText($newestCache.FullName, [System.Text.Encoding]::UTF8)
            $cacheObj = $cacheJson | ConvertFrom-Json
            if ($cacheObj.FetchedAt) {
                $dt = [DateTime]::Parse($cacheObj.FetchedAt)
                $cacheDateStr = $dt.ToString('dd.MM.yyyy HH:mm')
            }
        } catch { }
        if (-not $cacheDateStr) { $cacheDateStr = $newestCache.LastWriteTime.ToString('dd.MM.yyyy HH:mm') }
        $cacheName = $newestCache.Name
        if ($cacheFiles.Count -gt 1) { $cacheName = $cacheName + " (+$($cacheFiles.Count - 1) inny)" }
        Write-Host "  Cache metadanych: " -NoNewline
        Write-Host $cacheName -NoNewline -ForegroundColor Cyan
        Write-Host " | ostatni pobor: " -NoNewline
        Write-Host $cacheDateStr -ForegroundColor Cyan
        if ($script:LastZipPath -and $cacheObj.EntityAttributes) {
            $selectedZipName = [System.IO.Path]::GetFileName($script:LastZipPath)
            $selectedZipPath = [System.IO.Path]::GetFullPath($script:LastZipPath)
            $scanFile = Join-Path $outputDir 'CMT_SelectedZipEntities.json'
            $zipEnts = $null
            if (Test-Path $scanFile -PathType Leaf) {
                try {
                    $scanJson = [System.IO.File]::ReadAllText($scanFile, [System.Text.Encoding]::UTF8)
                    $scanData = $scanJson | ConvertFrom-Json
                    $scanPath = if ($scanData.ZipPath) { [System.IO.Path]::GetFullPath([string]$scanData.ZipPath) } else { $null }
                    $scanName = if ($scanData.ZipName) { [string]$scanData.ZipName } else { $null }
                    if (($scanName -eq $selectedZipName) -or ($scanPath -eq $selectedZipPath)) {
                        $zipEnts = @($scanData.Entities | ForEach-Object { [string]$_ })
                    }
                } catch { }
            }
            if ($null -eq $zipEnts) {
                $zipEnts = @(Get-EntitiesFromZip -ZipPath $script:LastZipPath)
                try {
                    $scanPayload = @{ ZipPath = $script:LastZipPath; ZipName = $selectedZipName; Entities = $zipEnts }
                    $scanPayload | ConvertTo-Json -Compress:$false | Set-Content -Path $scanFile -Encoding UTF8 -ErrorAction SilentlyContinue
                } catch { }
            }
            if ($zipEnts -and $zipEnts.Count -gt 0) {
                $cacheEnts = @($cacheObj.EntityAttributes.PSObject.Properties | ForEach-Object { $_.Name.Trim().ToLowerInvariant() })
                $missing = @($zipEnts | Where-Object { $e = $_.Trim().ToLowerInvariant(); $cacheEnts -notcontains $e })
                if ($missing.Count -eq 0) {
                    Write-Host "  Zaznaczona paczka: " -NoNewline
                    Write-Host "wszystkie encje w cache." -ForegroundColor Green
                } else {
                    Write-Host "  Zaznaczona paczka: " -NoNewline
                    Write-Host "brakuje $($missing.Count) encji: $($missing -join ', ')" -ForegroundColor Yellow
                }
            } else {
                Write-Host "  Zaznaczona paczka: nie wykryto encji w zipie (lub blad odczytu)." -ForegroundColor DarkGray
            }
        }
    } else {
        Write-Host "  Cache metadanych: brak (uruchom opcje 5)" -ForegroundColor DarkGray
    }
    Write-Host ""
}

do {
    Show-Menu
    $choice = Read-Host "Wybierz opcje (0-7)"
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
            $useLeadIdMap = $true
            $useOpportunityIdMap = $true
            $useContactIdMap = $true
            $useAccountIdMap = $true
            $entityPrompts = @(
                @{ File = 'CMT_IdMap_Lead.json'; Prompt = 'Uzyc mapowania leadow (objectid w uwagach)?'; Var = 'useLeadIdMap' }
                @{ File = 'CMT_IdMap_Opportunity.json'; Prompt = 'Uzyc mapowania szans (opportunity)?'; Var = 'useOpportunityIdMap' }
                @{ File = 'CMT_IdMap_Contact.json'; Prompt = 'Uzyc mapowania kontaktow (contact)?'; Var = 'useContactIdMap' }
                @{ File = 'CMT_IdMap_Account.json'; Prompt = 'Uzyc mapowania kont (account)?'; Var = 'useAccountIdMap' }
            )
            if ($zipPath -match 'annotation') {
                Write-Host "Uwaga: ten zip wyglada na paczke Uwag (annotation). Wybierz T przy mapowaniu leadow, inaczej import w CMT konczy sie bledem 'parent object type was present, but the ID was missing'." -ForegroundColor Yellow
            }
            foreach ($ep in $entityPrompts) {
                $fp = Join-Path $outputDir $ep.File
                if (Test-Path $fp -PathType Leaf) {
                    $r = (Read-Host ($ep.Prompt + ' [T/n]: ')).Trim().ToLowerInvariant()
                    if ($r -eq 'n' -or $r -eq 'no') {
                        switch ($ep.Var) {
                            'useLeadIdMap' { $useLeadIdMap = $false }
                            'useOpportunityIdMap' { $useOpportunityIdMap = $false }
                            'useContactIdMap' { $useContactIdMap = $false }
                            'useAccountIdMap' { $useAccountIdMap = $false }
                        }
                    }
                }
            }
            try {
                # Uruchom w osobnym procesie PowerShell - modul Microsoft.Xrm.Data.PowerShell nie laduje sie w hostach Cursor/VS Code
                $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$transformScript`"", '-InputZipPath', "`"$zipPath`"", '-OutputZipPath', "`"$outZip`"")
                if ($idMap) { $argList += '-IdMapPath'; $argList += "`"$idMap`"" }
                if (Test-Path $configPath) { $argList += '-ConfigPath'; $argList += "`"$configPath`"" }
                if (-not $useLeadIdMap) { $argList += '-NoLeadIdMap' }
                if (-not $useOpportunityIdMap) { $argList += '-NoOpportunityIdMap' }
                if (-not $useContactIdMap) { $argList += '-NoContactIdMap' }
                if (-not $useAccountIdMap) { $argList += '-NoAccountIdMap' }
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = 'powershell.exe'
                $psi.Arguments = $argList -join ' '
                $psi.WorkingDirectory = $scriptRoot
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $false
                Write-Host "Uruchamiam transformacje o $(Get-Date -Format 'HH:mm')..." -ForegroundColor Gray
                $p = [System.Diagnostics.Process]::Start($psi)
                $p.WaitForExit()
                if ($p.ExitCode -eq 0) {
                    Write-Host "Gotowy zip: $outZip" -ForegroundColor Green
                    Write-Host "W zipie: ownerzy podmienieni (IdMap), overriddencreatedon = oryginalna data." -ForegroundColor Cyan
                    if ($useLeadIdMap) {
                        Write-Host "Mapowanie leadow (objectid w uwagach): wlaczone - GUID leadow podmienione na cel." -ForegroundColor Green
                    } else {
                        Write-Host "Uwaga: mapowanie leadow wylaczone. Jesli zip zawiera Uwagi (annotation), import moze konczyc sie bledem 'parent object type was present, but the ID was missing'. Uruchom opcje 3 ponownie i wybierz T przy mapowaniu leadow." -ForegroundColor Yellow
                    }
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
        '5' {
            Add-MenuLog -NoConsole "Opcja 5: Pobierz metadane celu"
            $metaScript = Join-Path $scriptRoot 'Lib\Export-TargetMetadataCache.ps1'
            if (-not (Test-Path $metaScript)) {
                Write-Host "Brak Lib\Export-TargetMetadataCache.ps1" -ForegroundColor Red
                pause
                break
            }
            if (-not (Test-Path $configPath)) {
                Write-Host "Brak Config\CMTConfig.ps1 (Polaczenia.txt: Cel)." -ForegroundColor Red
                pause
                break
            }
            $zipForMeta = $script:LastZipPath
            if (-not $zipForMeta -or -not (Test-Path $zipForMeta)) {
                $inputZips = @(Get-ChildItem -Path $inputDir -Filter *.zip -File -ErrorAction SilentlyContinue)
                if ($inputZips.Count -gt 0) { $zipForMeta = $inputZips[0].FullName }
            }
            try {
                $metaArgStr = "-NoProfile -ExecutionPolicy Bypass -File `"$metaScript`" -ConfigPath `"$configPath`" -OutputDirectory `"$outputDir`""
                if ($zipForMeta -and (Test-Path $zipForMeta)) {
                    $metaArgStr += " -ZipPath `"$zipForMeta`""
                    Write-Host "Encje do pobrania z wybranego zipa: $([System.IO.Path]::GetFileName($zipForMeta))" -ForegroundColor Gray
                } else {
                    Write-Host "Brak wybranego zipa - pobiorę encje domyślne (ustaw zip w opcji 1 dla encji z zipa)." -ForegroundColor Yellow
                }
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = 'powershell.exe'
                $psi.Arguments = $metaArgStr
                $psi.WorkingDirectory = $scriptRoot
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $false
                $p = [System.Diagnostics.Process]::Start($psi)
                $p.WaitForExit()
                if ($p.ExitCode -eq 0) {
                    Write-Host "Cache metadanych zapisany w Output. Opcja 3 uzyje go (szybsza transformacja, bez polaczenia z CRM)." -ForegroundColor Green
                } else { Write-Host "Kod wyjscia: $($p.ExitCode)" -ForegroundColor Yellow }
            } catch {
                Add-MenuLog ('BLAD opcja 5: ' + $_)
                Write-Host ('BLAD: ' + $_) -ForegroundColor Red
            }
            pause
        }
        '6' {
            # Uwaga: w stringach tylko ASCII (cudzyslowy " ', myslnik - nie en-dash)
            Add-MenuLog -NoConsole "Opcja 6: Uzupelnij daty Uwag ze zrodla (wybor zip z Output, zmiana w tym samym pliku)"
            $injectScript = Join-Path $scriptRoot 'Lib\Inject-AnnotationDatesFromSource.ps1'
            if (-not (Test-Path $injectScript)) {
                Write-Host "Brak Lib\Inject-AnnotationDatesFromSource.ps1" -ForegroundColor Red
                pause
                break
            }
            if (-not (Test-Path $configPath)) {
                Write-Host "Brak Config\CMTConfig.ps1 (Polaczenia.txt: Zrodlo)." -ForegroundColor Red
                pause
                break
            }
            $outputZips = @(Get-ChildItem -Path $outputDir -Filter *.zip -File -ErrorAction SilentlyContinue | Sort-Object Name)
            if ($outputZips.Count -eq 0) {
                Write-Host "Brak zipow w Output ($outputDir). Wrzuc zip do tego folderu." -ForegroundColor Yellow
                pause
                break
            }
            Write-Host "Zipy w Output ($outputDir):" -ForegroundColor Cyan
            for ($i = 0; $i -lt $outputZips.Count; $i++) {
                Write-Host "  $($i + 1). $($outputZips[$i].Name)" -ForegroundColor White
            }
            $prompt = "Numer (1-$($outputZips.Count)) lub pelna sciezka do innego zipa"
            if ($outputZips.Count -ge 1) { $prompt += " [Enter = 1]" }
            $prompt += ": "
            $p = (Read-Host $prompt).Trim()
            if ([string]::IsNullOrWhiteSpace($p)) {
                $zipPath = $outputZips[0].FullName
            } elseif ($p -match '^\d+$') {
                $idx = [int]$p
                if ($idx -ge 1 -and $idx -le $outputZips.Count) {
                    $zipPath = $outputZips[$idx - 1].FullName
                } else {
                    $zipPath = $p
                }
            } else {
                $zipPath = $p
                if (-not [System.IO.Path]::IsPathRooted($p) -and ($p -notmatch '^[A-Za-z]:')) {
                    $inOutput = Join-Path $outputDir $p
                    if (Test-Path $inOutput -PathType Leaf) { $zipPath = $inOutput }
                }
            }
            if (-not (Test-Path $zipPath -PathType Leaf)) {
                Write-Host "Plik nie istnieje: $zipPath" -ForegroundColor Red
                pause
                break
            }
            try {
                $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$injectScript`"", '-InputZipPath', "`"$zipPath`"", '-OutputZipPath', "`"$zipPath`"", '-ConfigPath', "`"$configPath`"")
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = 'powershell.exe'
                $psi.Arguments = $argList -join ' '
                $psi.WorkingDirectory = $scriptRoot
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $false
                Write-Host "Pobieranie createdon/modifiedon ze zrodla i zapis w wybranym zipie..." -ForegroundColor Gray
                $p = [System.Diagnostics.Process]::Start($psi)
                $p.WaitForExit()
                if ($p.ExitCode -eq 0) {
                    Write-Host "Zaktualizowano zip (w tym samym pliku): $zipPath" -ForegroundColor Green
                    Write-Host "Mozesz teraz uruchomic opcje 3 (Transform) dla tego zipa." -ForegroundColor Cyan
                } else {
                    Write-Host ("Kod wyjscia: " + $p.ExitCode) -ForegroundColor Yellow
                }
            } catch {
                Add-MenuLog ('BLAD opcja 6: ' + $_)
                Write-Host ('BLAD: ' + $_) -ForegroundColor Red
            }
            pause
        }
        '7' {
            Add-MenuLog -NoConsole "Opcja 7: Generuj IdMap encji (leady, szanse, kontakty, konta)"
            $leadMapScript = Join-Path $scriptRoot 'Lib\New-CMTLeadIdMap.ps1'
            $otherMapScript = Join-Path $scriptRoot 'Lib\New-CMTOtherEntityIdMaps.ps1'
            if (-not (Test-Path $leadMapScript)) {
                Write-Host "Brak Lib\New-CMTLeadIdMap.ps1" -ForegroundColor Red
                pause
                break
            }
            if (-not (Test-Path $otherMapScript)) {
                Write-Host "Brak Lib\New-CMTOtherEntityIdMaps.ps1" -ForegroundColor Red
                pause
                break
            }
            if (-not (Test-Path $configPath)) {
                Write-Host "Brak Config\CMTConfig.ps1 (Polaczenia.txt: Zrodlo + Cel)." -ForegroundColor Red
                pause
                break
            }
            $allOk = $true
            try {
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = 'powershell.exe'
                $psi.UseShellExecute = $false
                $psi.CreateNoWindow = $false
                $psi.WorkingDirectory = $scriptRoot
                $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$leadMapScript`" -ConfigPath `"$configPath`""
                $p = [System.Diagnostics.Process]::Start($psi)
                $p.WaitForExit()
                if ($p.ExitCode -ne 0) {
                    Write-Host "Lead IdMap: kod wyjscia $($p.ExitCode)" -ForegroundColor Yellow
                    $allOk = $false
                }
                $psi.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$otherMapScript`" -ConfigPath `"$configPath`""
                $p = [System.Diagnostics.Process]::Start($psi)
                $p.WaitForExit()
                if ($p.ExitCode -ne 0) {
                    Write-Host "IdMap szans/kontaktow/kont: kod wyjscia $($p.ExitCode)" -ForegroundColor Yellow
                    $allOk = $false
                }
                if ($allOk) {
                    Write-Host "Wszystkie IdMap (Lead, Szanse, Kontakty, Konta) zapisane w Output. Przy opcji 3 mozesz wlaczyc/wylaczyc kazde z osobna." -ForegroundColor Green
                }
            } catch {
                Add-MenuLog ('BLAD opcja 7: ' + $_)
                Write-Host ('BLAD: ' + $_) -ForegroundColor Red
            }
            pause
        }
        '0' { Add-MenuLog -NoConsole 'Wyjscie'; exit 0 }
        default { Write-Host 'Wybierz 0-7.' -ForegroundColor Yellow; Start-Sleep -Seconds 2 }
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

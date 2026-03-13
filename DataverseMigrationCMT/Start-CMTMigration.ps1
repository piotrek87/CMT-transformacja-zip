# Migracja danych Dataverse przez Configuration Migration Tool (CMT)
# Używa: Microsoft.Xrm.Tooling.ConfigurationMigration (Export-CrmDataFile, Import-CrmDataFile)
# Schemat: wyciągasz ze źródła (lub używasz istniejącego); migrujemy tylko to, co pasuje do schematu.
# CMT obsługuje m.in. daty utworzenia i mapowanie właścicieli (User Map).

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false)]
    [string] $ConfigPath,
    [Parameter(Mandatory = $false)]
    [string] $SchemaFilePath,
    [Parameter(Mandatory = $false)]
    [string] $SourceConnectionString,
    [Parameter(Mandatory = $false)]
    [string] $TargetConnectionString,
    [Parameter(Mandatory = $false)]
    [switch] $ExportOnly,
    [Parameter(Mandatory = $false)]
    [switch] $ImportOnly,
    [Parameter(Mandatory = $false)]
    [string] $ImportDataFile,
    [Parameter(Mandatory = $false)]
    [string] $UserMapFilePath,
    [Parameter(Mandatory = $false)]
    [switch] $DisableTelemetry = $true,
    [Parameter(Mandatory = $false)]
    [string] $LogDirectory
)

$ErrorActionPreference = 'Stop'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$configDir = Join-Path $scriptRoot 'Config'
$libDir = Join-Path $scriptRoot 'Lib'
$defaultLogDir = Join-Path $scriptRoot 'Logs'

# Załaduj biblioteki
. (Join-Path $libDir 'Connect-CMT.ps1')

# Konfiguracja
if (-not $ConfigPath) { $ConfigPath = Join-Path $configDir 'CMTConfig.ps1' }
if (-not (Test-Path $ConfigPath)) {
    Write-Error "Brak pliku konfiguracji: $ConfigPath"
}
$config = & $ConfigPath

$srcConnStr = if ($SourceConnectionString) { $SourceConnectionString } else { $config.SourceConnectionString }
$tgtConnStr = if ($TargetConnectionString) { $TargetConnectionString } else { $config.TargetConnectionString }
$schemaPath = if ($SchemaFilePath) { $SchemaFilePath } else { $config.SchemaFilePath }
$logDir = if ($LogDirectory) { $LogDirectory } else { $defaultLogDir }
$userMap = if ($UserMapFilePath) { $UserMapFilePath } else { $config.UserMapFilePath }

$exportDir = $config.ExportOutputDirectory
if (-not [string]::IsNullOrWhiteSpace($exportDir)) {
    if (-not (Test-Path $exportDir)) { New-Item -ItemType Directory -Path $exportDir -Force | Out-Null }
}
$dataFileExportPath = Join-Path $exportDir $config.ImportDataFileName
$exportZip = if ($ImportDataFile) { $ImportDataFile } else { $dataFileExportPath }
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

function Write-Log {
    param([string]$Message)
    $line = "{0:yyyy-MM-dd HH:mm:ss} {1}" -f (Get-Date), $Message
    Write-Host $line
    $logFile = Join-Path $logDir ("CMT_{0:yyyyMMdd}.log" -f (Get-Date))
    Add-Content -Path $logFile -Value $line -ErrorAction SilentlyContinue
}

# Szuka dziennika CMT w AppData i zwraca linie z bledami (lub ostatnie linie)
function Get-CMTErrorLogContent {
    $maxMinutes = 30
    $appDataMicrosoft = Join-Path $env:APPDATA 'Microsoft'
    if (-not (Test-Path $appDataMicrosoft)) { return @() }
    $allLogs = Get-ChildItem -Path $appDataMicrosoft -Filter *.log -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like '*DataMigration*' -or $_.Name -like '*Configuration*' -or $_.Name -like '*Migration*' } |
        Where-Object { ((Get-Date) - $_.LastWriteTime).TotalMinutes -lt $maxMinutes } |
        Sort-Object LastWriteTime -Descending
    $latestLog = $allLogs | Select-Object -First 1
    if (-not $latestLog -or -not (Test-Path $latestLog.FullName)) { return @() }
    $content = Get-Content -Path $latestLog.FullName -Encoding UTF8 -ErrorAction SilentlyContinue
    if (-not $content) { return @() }
    $errorKeywords = @('Error', 'Failed', 'Fail', 'Missing', 'EXIT EXPORT', 'EXIT IMPORT', 'Validation', 'Schema', 'Exception', 'blad', 'nie powiodl')
    $errorLines = @($content | Where-Object {
        $line = $_
        $lineLower = $line.ToLowerInvariant()
        foreach ($k in $errorKeywords) {
            if ($lineLower.Contains($k.ToLowerInvariant())) { return $true }
        }
        $false
    })
    if ($errorLines.Count -gt 0) {
        return @($errorLines | Select-Object -First 40)
    }
    return @($content | Select-Object -Last 35)
}

# Encje wymagane przez walidacje CMT – dopisujemy do schematu, zeby uniknac bledu "Missing entities"
$script:CMTRequiredEntities = @(
    @{ name = 'activitypointer'; pk = 'activityid'; pn = 'subject'; etc = 4200 }
    @{ name = 'workflow'; pk = 'workflowid'; pn = 'name'; etc = 4703 }
    @{ name = 'processstage'; pk = 'processstageid'; pn = 'stagename'; etc = 4724 }
    @{ name = 'listmember'; pk = 'listmemberid'; pn = 'name'; etc = 4301 }
)
function Get-SchemaPathWithRequiredEntities {
    param([string]$SchemaPath)
    if (-not $SchemaPath -or -not (Test-Path $SchemaPath)) { return $SchemaPath }
    $xmlContent = [System.IO.File]::ReadAllText($SchemaPath, [System.Text.UTF8Encoding]::new($false))
    $missing = @()
    try {
        $doc = [xml]$xmlContent
        $existingNames = @($doc.entities.entity | ForEach-Object { $_.name })
        $existingSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($n in $existingNames) { [void]$existingSet.Add($n) }
        $missing = @($script:CMTRequiredEntities | Where-Object { -not $existingSet.Contains($_.name) })
    } catch {
        $xmlLower = $xmlContent.ToLowerInvariant()
        $missing = @($script:CMTRequiredEntities | Where-Object {
            -not ($xmlLower -match ('<entity\s[^>]*name="' + [regex]::Escape($_.name) + '"'))
        })
    }
    $stubs = New-Object System.Text.StringBuilder
    foreach ($m in $missing) {
        [void]$stubs.AppendLine("  <entity name=`"$($m.name)`" displayname=`"$($m.name)`" etc=`"$($m.etc)`" primaryidfield=`"$($m.pk)`" primarynamefield=`"$($m.pn)`" disableplugins=`"false`">")
        [void]$stubs.AppendLine('    <fields>')
        [void]$stubs.AppendLine("      <field displayname=`"$($m.pk)`" name=`"$($m.pk)`" type=`"guid`" updateCompare=`"true`" primaryKey=`"true`" />")
        [void]$stubs.AppendLine('    </fields>')
        [void]$stubs.AppendLine('    <relationships />')
        [void]$stubs.AppendLine('  </entity>')
    }
    $insert = $stubs.ToString().TrimEnd()
    $patched = if ($insert) { $xmlContent -replace '</entities>', "$insert`r`n</entities>" } else { $xmlContent }
    $tempFile = [System.IO.Path]::Combine($env:TEMP, 'CMTSchemaPatched_' + [Guid]::NewGuid().ToString('N').Substring(0, 8) + '.xml')
    [System.IO.File]::WriteAllText($tempFile, $patched, [System.Text.UTF8Encoding]::new($false))
    if ($missing.Count -gt 0) {
        Write-Log "Schemat: dopisano brakujace encje CMT ($($missing.Count)): $($missing.name -join ', ') -> $tempFile"
    } else {
        Write-Log "Schemat: kopia w TEMP (bez spacji): $tempFile"
    }
    return $tempFile
}

# Walidacja
if (-not $ExportOnly -and -not $ImportOnly -and -not $schemaPath) {
    Write-Error "Podaj sciezke do pliku schematu CMT (SchemaFilePath w config lub -SchemaFilePath). Schemat mozesz utworzyc w Configuration Migration Tool (GUI) lub uzyc istniejacego."
}
if (-not $ExportOnly -and -not $ImportOnly -and -not $srcConnStr) {
    Write-Error "Podaj SourceConnectionString (config lub -SourceConnectionString)."
}
if (-not $ExportOnly -and -not $ImportOnly -and -not $tgtConnStr) {
    Write-Error "Podaj TargetConnectionString (config lub -TargetConnectionString)."
}

# Instalacja/ładowanie modułu CMT
Install-CMTModuleIfNeeded -MinimumVersion '1.0.0.88' | Out-Null
Import-Module Microsoft.Xrm.Tooling.ConfigurationMigration -Force

if ($ExportOnly) {
    Write-Log "Eksport CMT: start (Start-CMTMigration.ps1)"
    if (-not $schemaPath -or -not (Test-Path $schemaPath)) {
        Write-Error "Eksport wymaga istniejacego pliku schematu: $schemaPath"
    }
    $schemaPath = (Resolve-Path -LiteralPath $schemaPath -ErrorAction Stop).Path
    Write-Log "Przygotowanie schematu (encje CMT: activitypointer, workflow, processstage, listmember)..."
    $schemaPathToUse = Get-SchemaPathWithRequiredEntities -SchemaPath $schemaPath
    if (-not $srcConnStr) { Write-Error "Eksport wymaga SourceConnectionString." }
    Write-Log "Polaczenie ze zrodlem..."
    $srcConn = New-CMTConnection -ConnectionString $srcConnStr
    Write-Log "Eksport danych (schemat: $schemaPathToUse) -> $exportZip"
    $exportParams = @{
        CrmConnection     = $srcConn
        SchemaFile        = $schemaPathToUse
        DataFile          = $exportZip
        LogWriteDirectory = $logDir
        DisableTelemetry  = $DisableTelemetry
        EmitLogToConsole  = $true
    }
    $transcriptPath = Join-Path $logDir ("CMT_Export_Transcript_{0:yyyyMMdd_HHmmss}.txt" -f (Get-Date))
    $transcriptStarted = $false
    $exportDone = $false
    try {
        try { Start-Transcript -Path $transcriptPath -Append -ErrorAction Stop; $transcriptStarted = $true } catch { Write-Log "Transkrypt niedostepny: $($_.Exception.Message)" }
        if ($PSCmdlet.ShouldProcess($exportZip, 'Export-CrmDataFile')) {
            Export-CrmDataFile @exportParams
            $exportDone = $true
        }
        if ($exportDone) { Write-Log "Eksport zakonczony: $exportZip" }
    } catch {
        # Jednorazowa proba z sciezkami w TEMP (bez spacji/OneDrive) - diagnoza
        if (-not $exportDone -and ($schemaPathToUse -match '\s|OneDrive' -or $exportZip -match '\s|OneDrive')) {
            $tempDir = [System.IO.Path]::Combine($env:TEMP, 'CMTExport_' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
            try {
                New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
                $tempSchema = Join-Path $tempDir 'data_schema.xml'
                $tempZip = Join-Path $tempDir 'CMT_Export.zip'
                Copy-Item -LiteralPath $schemaPathToUse -Destination $tempSchema -Force
                Write-Log "Proba z sciezkami w TEMP (bez spacji): $tempSchema -> $tempZip"
                Export-CrmDataFile -CrmConnection $srcConn -SchemaFile $tempSchema -DataFile $tempZip -LogWriteDirectory $tempDir -DisableTelemetry:$DisableTelemetry -EmitLogToConsole:$true
                Copy-Item -LiteralPath $tempZip -Destination $exportZip -Force
                Write-Log "Eksport zakonczony (przez TEMP): $exportZip"
                $exportDone = $true
            } catch {
                Write-Log "Eksport z TEMP tez nie powiodl sie: $($_.Exception.Message)"
            } finally {
                if (Test-Path $tempDir) { Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
            }
        }
        if (-not $exportDone) {
            if ($transcriptStarted) { try { Stop-Transcript } catch { } }
            $errMsg = $_.Exception.Message
            $inner = $_.Exception.InnerException
            while ($inner) {
                $errMsg += " | Inner: " + $inner.Message
                $inner = $inner.InnerException
            }
            Write-Log "BLAD eksportu CMT: $errMsg"
            Write-Log "Pelny wyjatek: $($_.Exception.ToString())"
            if ($_.ErrorDetails.Message) { Write-Log "ErrorDetails: $($_.ErrorDetails.Message)" }
            if ($_.Exception.StackTrace) { Write-Log "StackTrace: $($_.Exception.StackTrace)" }
            if (Test-Path $transcriptPath) {
                Write-Log "Transkrypt eksportu: $transcriptPath"
                Get-Content -Path $transcriptPath -Tail 80 -Encoding UTF8 -ErrorAction SilentlyContinue | ForEach-Object { Write-Log "  TX: $_" }
            }
            $cmtLogs = Get-ChildItem -Path $logDir -Filter *.log -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 3
            foreach ($lf in $cmtLogs) {
                Write-Log "Log CMT: $($lf.FullName)"
                Get-Content -Path $lf.FullName -Tail 50 -Encoding UTF8 -ErrorAction SilentlyContinue | ForEach-Object { Write-Log "  $_" }
            }
            # Zaciagnij dziennik bledow CMT z AppData (DataMigrationUtility tam zapisuje)
            $cmtErrorLines = Get-CMTErrorLogContent
            $appDataMicrosoft = Join-Path $env:APPDATA 'Microsoft'
            $cmtLogPath = $null
            if (Test-Path $appDataMicrosoft) {
                $cmtLogs = Get-ChildItem -Path $appDataMicrosoft -Filter *.log -Recurse -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -like '*DataMigration*' -or $_.Name -like '*Configuration*' } |
                    Sort-Object LastWriteTime -Descending
                $latest = $cmtLogs | Where-Object { ((Get-Date) - $_.LastWriteTime).TotalMinutes -lt 30 } | Select-Object -First 1
                if ($latest) { $cmtLogPath = $latest.FullName }
            }
            if ($cmtLogPath) { Write-Log "Dziennik CMT (AppData): $cmtLogPath" }
            if ($cmtErrorLines -and $cmtErrorLines.Count -gt 0) {
                Write-Log "Wybrane wpisy z dziennika bledow CMT:"
                foreach ($line in $cmtErrorLines) { Write-Log "  CMT: $line" }
            }
            throw
        }
    }
    if ($transcriptStarted) { try { Stop-Transcript } catch { } }
    return
}

if ($ImportOnly) {
    $fileToImport = if ($ImportDataFile) { $ImportDataFile } else { $dataFileExportPath }
    if (-not (Test-Path $fileToImport)) {
        Write-Error "Plik do importu nie istnieje: $fileToImport"
    }
    if (-not $tgtConnStr) { Write-Error "Import wymaga TargetConnectionString." }
    Write-Log "Polaczenie z celem..."
    $tgtConn = New-CMTConnection -ConnectionString $tgtConnStr
    Write-Log "Import danych: $fileToImport"
    $importParams = @{
        CrmConnection     = $tgtConn
        DataFile          = $fileToImport
        LogWriteDirectory = $logDir
        DisableTelemetry  = $DisableTelemetry
    }
    if (-not [string]::IsNullOrWhiteSpace($userMap) -and (Test-Path $userMap)) {
        $importParams['UserMapFile'] = $userMap
        Write-Log "Mapowanie uzytkownikow: $userMap"
    }
    if ($PSCmdlet.ShouldProcess($fileToImport, 'Import-CrmDataFile')) {
        Import-CrmDataFile @importParams
    }
    Write-Log "Import zakonczony."
    return
}

# Pełna migracja: eksport ze źródła -> import do celu
if (-not $schemaPath -or -not (Test-Path $schemaPath)) {
    Write-Error "Podaj istniejacy plik schematu CMT (SchemaFilePath)."
}
Write-Log "=== Migracja CMT: zrodlo -> cel ==="
Write-Log "Schemat: $schemaPath"
Write-Log "Katalog wyjsciowy: $exportDir"

Write-Log "Polaczenie ze zrodlem..."
$srcConn = New-CMTConnection -ConnectionString $srcConnStr
Write-Log "Eksport ze zrodla..."
try {
    $verboseBackup = $VerbosePreference
    $VerbosePreference = 'Continue'
    $exportOut = Export-CrmDataFile -CrmConnection $srcConn -SchemaFile $schemaPath -DataFile $exportZip -LogWriteDirectory $logDir -DisableTelemetry:$DisableTelemetry -Verbose 4>&1
    $VerbosePreference = $verboseBackup
    if ($exportOut) { $exportOut | ForEach-Object { Write-Log "CMT: $_" } }
    Write-Log "Eksport zapisany: $exportZip"
} catch {
    $errMsg = $_.Exception.Message
    $inner = $_.Exception.InnerException
    while ($inner) {
        $errMsg += " | Inner: " + $inner.Message
        $inner = $inner.InnerException
    }
    Write-Log "BLAD eksportu CMT: $errMsg"
    Write-Log "Pelny wyjatek: $($_.Exception.ToString())"
    if ($_.ErrorDetails.Message) { Write-Log "ErrorDetails: $($_.ErrorDetails.Message)" }
    if ($_.Exception.StackTrace) { Write-Log "StackTrace: $($_.Exception.StackTrace)" }
    # Log CMT z AppData (np. schema validation failed, missing fields)
    $microsoftRoaming = Join-Path $env:APPDATA 'Microsoft'
    if (Test-Path $microsoftRoaming) {
        $cmtFolders = Get-ChildItem -Path $microsoftRoaming -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like '*Configuration Migration*' -or $_.Name -like '*Dynamics*CRM*Migration*' }
        foreach ($cmtDir in $cmtFolders) {
            $allLogs = Get-ChildItem -Path $cmtDir.FullName -Filter *.log -Recurse -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
            $latestLog = $allLogs | Select-Object -First 1
            if ($latestLog -and ((Get-Date) - $latestLog.LastWriteTime).TotalMinutes -lt 15) {
                Write-Log "Ostatni log CMT (AppData): $($latestLog.FullName)"
                Get-Content -Path $latestLog.FullName -Tail 80 -Encoding UTF8 -ErrorAction SilentlyContinue | ForEach-Object { Write-Log "  CMT: $_" }
                break
            }
        }
    }
    throw
}

Write-Log "Polaczenie z celem..."
$tgtConn = New-CMTConnection -ConnectionString $tgtConnStr
$importParams = @{
    CrmConnection     = $tgtConn
    DataFile          = $exportZip
    LogWriteDirectory = $logDir
    DisableTelemetry  = $DisableTelemetry
}
if (-not [string]::IsNullOrWhiteSpace($userMap) -and (Test-Path $userMap)) {
    $importParams['UserMapFile'] = $userMap
}
Write-Log "Import do celu..."
Import-CrmDataFile @importParams
Write-Log "=== Migracja CMT zakonczona ==="

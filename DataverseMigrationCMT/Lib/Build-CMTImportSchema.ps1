# Porownuje schemat zrodla i celu, buduje schemat importu (tylko encje i atrybuty obecne w celu).
# Wynik: data_schema_import.xml do uzycia z CMT Export ze zrodla, potem transformacja zip i import w CMT.
# Uruchom w osobnym procesie PowerShell (modul Xrm wymaga standardowego hosta).

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string] $ConfigPath,
    [Parameter(Mandatory = $false)]
    [string] $OutputPath
)

$ErrorActionPreference = 'Stop'
$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$libDir = $scriptRoot
$schemaScript = Join-Path $libDir 'Get-CMTSchemaFromSource.ps1'
if (-not (Test-Path $schemaScript)) { throw "Brak Get-CMTSchemaFromSource.ps1 w $libDir" }
. $schemaScript

$configDir = Join-Path $scriptRoot '..\Config'
if (-not $ConfigPath) { $ConfigPath = Join-Path $configDir 'CMTConfig.ps1' }
if (-not (Test-Path $ConfigPath)) { throw "Brak configu: $ConfigPath" }
$config = & $ConfigPath

$logsDir = Join-Path $scriptRoot '..\Logs'
if (-not [System.IO.Directory]::Exists($logsDir)) { [System.IO.Directory]::CreateDirectory($logsDir) | Out-Null }
$script:SchemaLogPath = Join-Path $logsDir ("CMT_BuildSchema_{0:yyyyMMdd_HHmmss}.log" -f [DateTime]::Now)

Write-SchemaLog 'Build-CMTImportSchema: porownanie zrodlo + cel, budowa schematu importu.'
if ([string]::IsNullOrWhiteSpace($config.SourceConnectionString) -or [string]::IsNullOrWhiteSpace($config.TargetConnectionString)) {
    throw "W configu ustaw SourceConnectionString i TargetConnectionString (Config\Polaczenia.txt lub CMTConfig.ps1)."
}

$entityFilter = @($config.SchemaEntityIncludeOnly)
if (-not $entityFilter -or $entityFilter.Count -eq 0) {
    throw "W configu ustaw SchemaEntityIncludeOnly (lista encji do migracji)."
}
Write-SchemaLog ('Encje do porownania: ' + $entityFilter.Count + ' (SchemaEntityIncludeOnly).')

Write-SchemaLog 'Pobieranie metadanych ze zrodla...'
$sourceMeta = Get-SourceMetadataForCMTSchema -ConnectionString $config.SourceConnectionString -EntityFilter $entityFilter
Write-SchemaLog ('Zrodlo: ' + $sourceMeta.Count + ' encji.')

Write-SchemaLog 'Pobieranie metadanych z celu...'
$targetMeta = Get-SourceMetadataForCMTSchema -ConnectionString $config.TargetConnectionString -EntityFilter $entityFilter
Write-SchemaLog ('Cel: ' + $targetMeta.Count + ' encji.')

$exclude = @($config.SystemEntitiesToSkip)
if (-not $exclude) { $exclude = @() }
$common = Get-CommonEntitiesForCMT -SourceMetadata $sourceMeta -TargetMetadata $targetMeta -ExcludeEntities $exclude
Write-SchemaLog ('Wspolne encje (atrybuty z celu): ' + $common.Count + '.')

# CMT walidacja: kazda encja referencowana w lookupach musi byc w schemacie – dopelnij brakujace
$referencedEntities = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($entName in $common.Keys) {
    $attrs = $common[$entName].Attributes
    if (-not $attrs) { continue }
    foreach ($a in $attrs) {
        if ($a.Targets -and $a.Targets.Count -gt 0) {
            foreach ($t in $a.Targets) {
                if (-not [string]::IsNullOrWhiteSpace($t)) { [void]$referencedEntities.Add($t.Trim()) }
            }
        }
    }
}
$missingRefs = @($referencedEntities | Where-Object { -not $common.ContainsKey($_) })
if ($missingRefs.Count -gt 0) {
    Write-SchemaLog ('Brakujace encje referencowane w lookupach (' + $missingRefs.Count + '): ' + ($missingRefs -join ', '))
    Write-SchemaLog 'Pobieranie metadanych brakujacych encji ze zrodla...'
    $missingMeta = Get-SourceMetadataForCMTSchema -ConnectionString $config.SourceConnectionString -EntityFilter $missingRefs
    foreach ($name in $missingMeta.Keys) {
        if ($common.ContainsKey($name)) { continue }
        $common[$name] = $missingMeta[$name]
        Write-SchemaLog ('  Dodano encje zalezna: ' + $name)
    }
    # Na poczatek kolejnosci (zaleznosci pierwsze)
    $entityOrder = @($config.SchemaEntityOrder)
    if (-not $entityOrder -or $entityOrder.Count -eq 0) { $entityOrder = @() }
    $existingSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($e in $entityOrder) { [void]$existingSet.Add($e) }
    foreach ($m in $missingRefs) {
        if (-not $existingSet.Contains($m)) { $entityOrder = @(,$m) + $entityOrder; [void]$existingSet.Add($m) }
    }
    $config = @{ } + $config
    $config['SchemaEntityOrder'] = $entityOrder
}

# CMT walidacja czesto wymaga encji bazowych/zaleznych, ktorych API nie zwraca w Targets – dopelnij z listy znanych
$cmtRequiredExtra = @('activitypointer', 'workflow', 'processstage', 'listmember')
$missingCmt = @($cmtRequiredExtra | Where-Object { -not $common.ContainsKey($_) })
if ($missingCmt.Count -gt 0) {
    Write-SchemaLog ('Encje wymagane przez CMT (brak w schemacie): ' + ($missingCmt -join ', '))
    Write-SchemaLog 'Pobieranie metadanych tych encji ze zrodla...'
    $extraMeta = Get-SourceMetadataForCMTSchema -ConnectionString $config.SourceConnectionString -EntityFilter $missingCmt
    foreach ($name in $extraMeta.Keys) {
        if ($common.ContainsKey($name)) { continue }
        $common[$name] = $extraMeta[$name]
        Write-SchemaLog ('  Dodano encje dla CMT: ' + $name)
    }
    $entityOrder = @($config.SchemaEntityOrder)
    if (-not $entityOrder -or $entityOrder.Count -eq 0) { $entityOrder = @() }
    $existingSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($e in $entityOrder) { [void]$existingSet.Add($e) }
    foreach ($m in $missingCmt) {
        if (-not $existingSet.Contains($m)) { $entityOrder = @(,$m) + $entityOrder; [void]$existingSet.Add($m) }
    }
    $config = @{ } + $config
    $config['SchemaEntityOrder'] = $entityOrder
}

# Jesli API nie zwrocilo ktorejs z encji wymaganych przez CMT – dopisz minimalna definicje (stub), zeby walidacja schematu przeszla
$stillMissingCmt = @($cmtRequiredExtra | Where-Object { -not $common.ContainsKey($_) })
if ($stillMissingCmt.Count -gt 0) {
    $stubEntities = @{
        activitypointer = @{ pk = 'activityid'; pn = 'subject'; otc = 4200 }
        workflow        = @{ pk = 'workflowid'; pn = 'name'; otc = 4703 }
        processstage    = @{ pk = 'processstageid'; pn = 'stagename'; otc = 4724 }
        listmember      = @{ pk = 'listmemberid'; pn = 'name'; otc = 4301 }
    }
    foreach ($name in $stillMissingCmt) {
        $stub = $stubEntities[$name]
        if (-not $stub) { $stub = @{ pk = $name + 'id'; pn = 'name'; otc = '' } }
        $pkAttr = $stub.pk
        $common[$name] = @{
            LogicalName        = $name
            Attributes         = @(@{ LogicalName = $pkAttr; Type = 'Uniqueidentifier'; Targets = @() })
            PrimaryIdAttribute = $pkAttr
            PrimaryNameAttribute = $stub.pn
            ObjectTypeCode     = $stub.otc
        }
        Write-SchemaLog ('  Dodano stub encji dla CMT (brak w API): ' + $name)
    }
    $entityOrder = @($config.SchemaEntityOrder)
    if (-not $entityOrder -or $entityOrder.Count -eq 0) { $entityOrder = @() }
    $existingSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($e in $entityOrder) { [void]$existingSet.Add($e) }
    foreach ($m in $stillMissingCmt) {
        if (-not $existingSet.Contains($m)) { $entityOrder = @(,$m) + $entityOrder; [void]$existingSet.Add($m) }
    }
    $config = @{ } + $config
    $config['SchemaEntityOrder'] = $entityOrder
}

if ($common.Count -eq 0) {
    throw "Brak wspolnych encji. Sprawdz polaczenia i SchemaEntityIncludeOnly."
}

# Koncowa gwarancja: 4 encje wymagane przez walidacje CMT musza byc w schemacie (nawet jako stub)
$cmtRequiredFinal = @('activitypointer', 'workflow', 'processstage', 'listmember')
$stubDefs = @{
    activitypointer = @{ pk = 'activityid'; pn = 'subject'; otc = 4200 }
    workflow        = @{ pk = 'workflowid'; pn = 'name'; otc = 4703 }
    processstage    = @{ pk = 'processstageid'; pn = 'stagename'; otc = 4724 }
    listmember      = @{ pk = 'listmemberid'; pn = 'name'; otc = 4301 }
}
$entityOrder = @($config.SchemaEntityOrder)
if (-not $entityOrder -or $entityOrder.Count -eq 0) { $entityOrder = @() }
foreach ($name in $cmtRequiredFinal) {
    if ($common.ContainsKey($name)) { continue }
    $stub = $stubDefs[$name]
    if (-not $stub) { $stub = @{ pk = $name + 'id'; pn = 'name'; otc = '' } }
    $common[$name] = @{
        LogicalName          = $name
        Attributes           = @(@{ LogicalName = $stub.pk; Type = 'Uniqueidentifier'; Targets = @() })
        PrimaryIdAttribute   = $stub.pk
        PrimaryNameAttribute = $stub.pn
        ObjectTypeCode       = $stub.otc
    }
    $entityOrder = @(,$name) + @($entityOrder | Where-Object { $_ -ne $name })
    Write-SchemaLog ('Gwarancja CMT: dodano encje do schematu: ' + $name)
}
$config = @{ } + $config
$config['SchemaEntityOrder'] = $entityOrder

if (-not $OutputPath) {
    $outDir = $config.ExportOutputDirectory
    if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = Join-Path $scriptRoot '..\Output' }
    if (-not (Test-Path $outDir)) { [System.IO.Directory]::CreateDirectory($outDir) | Out-Null }
    $OutputPath = Join-Path $outDir 'data_schema_import.xml'
}

$entityOrder = @($config.SchemaEntityOrder)
if (-not $entityOrder -or $entityOrder.Count -eq 0) { $entityOrder = $null }
Export-CMTSchemaXml -EntityMetadata $common -OutputPath $OutputPath -EntityOrder $entityOrder
Write-SchemaLog ('Schemat importu zapisany: ' + $OutputPath)
# Zapis sciezki do uzycia w menu (eksport CMT uzywa tego schematu)
$lastSchemaFile = Join-Path ([System.IO.Path]::GetDirectoryName($ConfigPath)) 'LastGeneratedSchema.txt'
try { [System.IO.File]::WriteAllText($lastSchemaFile, $OutputPath, [System.Text.UTF8Encoding]::new($false)) } catch { }
Write-Host "Schemat importu (zrodlo vs cel): $OutputPath"
return $OutputPath

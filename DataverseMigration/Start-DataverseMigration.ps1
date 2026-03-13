<#
.SYNOPSIS
    Narzędzie migracyjne Dataverse - przenoszenie danych między środowiskami Dynamics 365 / Dataverse.
.DESCRIPTION
    Łączy się ze środowiskiem źródłowym i docelowym, wykrywa wspólne encje,
    migruje dane z zachowaniem relacji, CreatedOn (overriddencreatedon), statecode/statuscode, ownerid.
    Encje BPF (sufiks "process") są migrowane po encjach głównych.
.EXAMPLE
    .\Start-DataverseMigration.ps1 -SourceConnectionString "AuthType=OAuth;..." -TargetConnectionString "AuthType=OAuth;..." -Interactive:$false
.EXAMPLE
    .\Start-DataverseMigration.ps1 -Interactive
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string] $SourceConnectionString,
    [Parameter(Mandatory = $false)]
    [string] $TargetConnectionString,
    [Parameter(Mandatory = $false)]
    $SourceConn,
    [Parameter(Mandatory = $false)]
    $TargetConn,
    [Parameter(Mandatory = $false)]
    [switch] $Interactive,
    [Parameter(Mandatory = $false)]
    [string[]] $EntityFilter = @(),
    [Parameter(Mandatory = $false)]
    [string[]] $EntityExcludeFilter = @(),
    [Parameter(Mandatory = $false)]
    [string] $ConfigPath = "$PSScriptRoot\Config\MigrationConfig.ps1",
    [Parameter(Mandatory = $false)]
    [ValidateSet('Create','Update','Upsert')]
    [string] $MigrationMode = '',
    [Parameter(Mandatory = $false)]
    [ValidateSet('Id','IdThenName','Name','Custom')]
    [string] $MatchBy = '',
    [Parameter(Mandatory = $false)]
    [string] $CustomMatchAttribute = '',
    [Parameter(Mandatory = $false)]
    [switch] $OnlyEntitiesWithRecordsAndDependencies,
    [Parameter(Mandatory = $false)]
    [string] $EntityDefaultTargetLookupStr = '',
    [Parameter(Mandatory = $false)]
    [string[]] $EntityLookupResolveByName = @(),
    [Parameter(Mandatory = $false)]
    [int] $MaxRecordsPerEntity = 0,
    [Parameter(Mandatory = $false)]
    [string[]] $EntityIncludeOnly = @(),
    [Parameter(Mandatory = $false)]
    [switch] $WhatIf
)

$ErrorActionPreference = 'Stop'
$script:LogFile = $null
if ($env:DATAVERSE_MIGRATION_WHATIF -eq "1") { $WhatIf = $true }
if (-not [string]::IsNullOrWhiteSpace($env:DATAVERSE_MIGRATION_ENTITY_FILTER)) {
    $EntityFilter = @($env:DATAVERSE_MIGRATION_ENTITY_FILTER -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Write-MigrationLog {
    param([string] $Message, [string] $Level = 'INFO')
    $line = "{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}" -f (Get-Date), $Level, $Message
    Write-Host $line
    if ($script:LogFile) {
        Add-Content -Path $script:LogFile -Value $line -ErrorAction SilentlyContinue
    }
}

# Ładowanie konfiguracji
if (Test-Path $ConfigPath) {
    $Config = . $ConfigPath
} else {
    $Config = @{
        PageSize = 5000
        BatchSize = 100
        MaxRetryCount = 3
        RetryDelaySeconds = 5
        SystemFieldsToSkip = @('createdon','modifiedon','createdby','modifiedby','versionnumber','utcconversiontimezonecode','timezoneruleversionnumber','importsequencenumber')
        SystemEntitiesToSkip = @('bulkdeleteoperation','asyncoperation','workflow','pluginassembly')
        EntityOrderPriority = @('systemuser','team','account','contact','lead','opportunity','activitypointer','email','task','appointment')
        BpfEntitySuffix = 'process'
        MigrationMode = 'Upsert'
        MatchBy = 'IdThenName'
        IdMapPath = ''
        CustomMatchAttribute = ''
        EntityMatchKey = @{}
        LogFolder = ".\Logs"
        LogFileName = "Migration_{0:yyyyMMdd_HHmmss}.log"
    }
}
if (-not [string]::IsNullOrWhiteSpace($MigrationMode)) { $Config.MigrationMode = $MigrationMode }
if (-not [string]::IsNullOrWhiteSpace($MatchBy)) { $Config.MatchBy = $MatchBy }
if (-not [string]::IsNullOrWhiteSpace($CustomMatchAttribute)) { $Config.CustomMatchAttribute = $CustomMatchAttribute }
if (-not [string]::IsNullOrWhiteSpace($EntityDefaultTargetLookupStr)) {
    $defaultLookup = @{}
    $EntityDefaultTargetLookupStr -split "[\r\n;]+" | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^(\w+)\s*=\s*([\w\-]+)$' } | ForEach-Object {
        $defaultLookup[$Matches[1].ToLowerInvariant()] = $Matches[2].Trim()
    }
    $Config.EntityDefaultTargetLookup = $defaultLookup
}
if ($EntityLookupResolveByName -and $EntityLookupResolveByName.Count -gt 0) {
    $Config.EntityLookupResolveByName = @($EntityLookupResolveByName | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ })
}
if ($MaxRecordsPerEntity -gt 0) {
    $Config.MaxRecordsPerEntity = $MaxRecordsPerEntity
}
if ($EntityIncludeOnly -and $EntityIncludeOnly.Count -gt 0) {
    $Config.EntityIncludeOnly = @($EntityIncludeOnly | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

# Inicjalizacja logu
$logDir = $Config.LogFolder
if (-not [string]::IsNullOrWhiteSpace($logDir)) {
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $script:LogFile = Join-Path $logDir ($Config.LogFileName -f (Get-Date))
}
Write-MigrationLog "Start migracji Dataverse"

# Modul Xrm musi byc zaladowany w scope globalnym (gdy skrypt wywolany przez & z GUI)
Import-Module Microsoft.Xrm.Data.PowerShell -Force -Scope Global -ErrorAction Stop

# Zaladuj biblioteki (dot-source)
$libPath = Join-Path $PSScriptRoot 'Lib'
. (Join-Path $libPath 'Connect-Dataverse.ps1')
. (Join-Path $libPath 'Get-EntityMetadata.ps1')
. (Join-Path $libPath 'Get-MigrationOrder.ps1')
. (Join-Path $libPath 'Migrate-EntityData.ps1')

$createCmd = Get-Command -Name Add-CrmRecord -ErrorAction SilentlyContinue
if (-not $createCmd) { $createCmd = Get-Command -Name New-CrmRecord -ErrorAction SilentlyContinue }
if (-not $createCmd) {
    Write-MigrationLog "Blad: modul nie eksportuje Add-CrmRecord ani New-CrmRecord. Zainstaluj: Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser" "ERROR"
    throw "Modul Microsoft.Xrm.Data.PowerShell nie eksportuje Add-CrmRecord ani New-CrmRecord. Zainstaluj: Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser"
}

# Polaczenia
Write-MigrationLog "Nawiazywanie polaczen..."
if ($SourceConn -and $TargetConn) {
    Write-MigrationLog "Uzyto przekazanych polaczen (np. z Run-InteractiveMigration)."
} elseif ($Interactive) {
    $SourceConn = Connect-DataverseEnvironment -Interactive
    $TargetConn = $SourceConn
    Write-MigrationLog "Tryb interaktywny: jedno polaczenie. Dla dwoch srodowisk uzyj Run-InteractiveMigration.ps1."
} else {
    if ([string]::IsNullOrWhiteSpace($SourceConnectionString)) { $SourceConnectionString = $env:DATAVERSE_MIGRATION_SRC }
    if ([string]::IsNullOrWhiteSpace($TargetConnectionString)) { $TargetConnectionString = $env:DATAVERSE_MIGRATION_TGT }
    if ([string]::IsNullOrWhiteSpace($SourceConnectionString)) { $SourceConnectionString = $Config.SourceConnectionString }
    if ([string]::IsNullOrWhiteSpace($TargetConnectionString)) { $TargetConnectionString = $Config.TargetConnectionString }
    if ([string]::IsNullOrWhiteSpace($SourceConnectionString) -or [string]::IsNullOrWhiteSpace($TargetConnectionString)) {
        throw "Podaj SourceConnectionString i TargetConnectionString (lub ustaw w konfiguracji / zmiennych srodowiskowych). Lub uruchom Run-InteractiveMigration.ps1 z konsoli PowerShell."
    }
    $SourceConn = Connect-DataverseEnvironment -ConnectionString $SourceConnectionString
    $TargetConn = Connect-DataverseEnvironment -ConnectionString $TargetConnectionString
}
if (-not (Test-DataverseConnection -Connection $SourceConn)) { throw "Polaczenie ze zrodlem nie powiodlo sie." }
if (-not (Test-DataverseConnection -Connection $TargetConn)) { throw "Polaczenie z celem nie powiodlo sie." }
Write-MigrationLog "Polaczenia OK."

# Metadane (przy wykluczeniach pobierz wszystkie encje, potem odfiltruj)
$metaFilter = $EntityFilter
if ($EntityExcludeFilter -and $EntityExcludeFilter.Count -gt 0) { $metaFilter = @() }
Write-MigrationLog "Pobieranie metadanych encji ze zrodla..."
$sourceMeta = Get-EntityMetadataFromEnv -Connection $SourceConn -EntityLogicalNames $metaFilter
Write-MigrationLog "Pobieranie metadanych encji z celu..."
$targetMeta = Get-EntityMetadataFromEnv -Connection $TargetConn -EntityLogicalNames $metaFilter

$commonEntities = Get-CommonEntities -SourceMetadata $sourceMeta -TargetMetadata $targetMeta -ExcludeEntities $Config.SystemEntitiesToSkip
$orderedEntities = @(Get-MigrationOrderByDependencies -CommonEntities $commonEntities -Config $Config -BpfSuffix $Config.BpfEntitySuffix | Select-Object -Unique)

if ($Config.EntityIncludeOnly -and $Config.EntityIncludeOnly.Count -gt 0) {
    $orderedEntities = @(Get-EntitiesFromWhitelistAndDependencies -CommonEntities $commonEntities -Whitelist $Config.EntityIncludeOnly -Config $Config -BpfSuffix $Config.BpfEntitySuffix)
    Write-MigrationLog "Tryb: tylko encje biznesowe (EntityIncludeOnly) + zaleznosci: $($orderedEntities.Count) encji."
}

if ($OnlyEntitiesWithRecordsAndDependencies) {
    Write-MigrationLog "Liczenie rekordow w zrodle (tylko encje z danymi i ich zaleznosci)..."
    $recordCounts = Get-EntityRecordCounts -Connection $SourceConn -EntityLogicalNames $orderedEntities -Logger ${function:Write-MigrationLog}
    $orderedEntities = @(Get-EntitiesWithRecordsAndDependencies -CommonEntities $commonEntities -RecordCounts $recordCounts -Config $Config -BpfSuffix $Config.BpfEntitySuffix | Select-Object -Unique)
    Write-MigrationLog "Po filtracji (encje z rekordami + zaleznosci): $($orderedEntities.Count) encji."
}

if ($EntityExcludeFilter -and $EntityExcludeFilter.Count -gt 0) {
    $depGraph = Get-EntityDependencyGraph -CommonEntities $commonEntities -Config $Config -BpfSuffix $Config.BpfEntitySuffix
    $fullExclude = [System.Collections.ArrayList]::new(@($EntityExcludeFilter | Select-Object -Unique))
    $changed = $true
    while ($changed) {
        $changed = $false
        foreach ($e in $orderedEntities) {
            if ($e -in $fullExclude) { continue }
            $deps = $depGraph[$e]
            if (-not $deps) { continue }
            foreach ($d in $deps) {
                if ($d -in $fullExclude) {
                    $null = $fullExclude.Add($e)
                    $changed = $true
                    break
                }
            }
        }
    }
    $orderedEntities = @($orderedEntities | Where-Object { $_ -notin $fullExclude })
    $dependentExcluded = @($fullExclude | Where-Object { $_ -notin $EntityExcludeFilter })
    if ($dependentExcluded.Count -gt 0) {
        Write-MigrationLog "Wykluczono encje: $($EntityExcludeFilter -join ', '); rowniez zalezne: $($dependentExcluded -join ', ')"
    } else {
        Write-MigrationLog "Wykluczono encje: $($EntityExcludeFilter -join ', ')"
    }
}

Write-MigrationLog "Wspolne encje do migracji: $($orderedEntities.Count)"
foreach ($e in $orderedEntities) {
    $isBpf = $e -match "$($Config.BpfEntitySuffix)`$"
    Write-MigrationLog "  - $e $(if($isBpf){'(BPF)'})"
}

if ($WhatIf) {
    Write-MigrationLog "WhatIf: zakonczono bez migracji danych."
    exit 0
}

# Migracja po kolei (z ponownymi przebiegami, jesli sa bledy – zeby dopasowac relacje po utworzeniu brakujacych)
$stats = @{}
$entityTotal = $orderedEntities.Count
$retryPasses = if ($Config.RetryFailedRecordPasses -and [int]$Config.RetryFailedRecordPasses -gt 0) { [int]$Config.RetryFailedRecordPasses } else { 1 }
$useAutoPull = $OnlyEntitiesWithRecordsAndDependencies -or ($Config.EntityIncludeOnly -and $Config.EntityIncludeOnly.Count -gt 0)
if ($useAutoPull) {
    Write-MigrationLog "Wlaczone: AutoMigrateMissingLookups (dociaganie brakujacych rekordow ze zrodla przy relacjach)."
}
for ($pass = 1; $pass -le $retryPasses; $pass++) {
    if ($pass -gt 1) {
        $totalFailed = ($stats.Values | ForEach-Object { if ($_.Failed) { $_.Failed } else { 0 } } | Measure-Object -Sum).Sum
        if ($totalFailed -eq 0) { break }
        Write-MigrationLog ('Ponowny przebieg ' + $pass + '/' + $retryPasses + ' (po bledach - IdMap uzupelniony, ponawiam probe).')
    }
    $entityNum = 0
    foreach ($entityName in $orderedEntities) {
        $entityNum++
        $entityPct = if ($entityTotal -gt 0) { [int](($entityNum - 1) / $entityTotal * 100) } else { 0 }
        Write-Progress -Activity "Migracja Dataverse" -Status "Encja $entityNum z ${entityTotal}: $entityName$(if ($pass -gt 1) { " (przebieg $pass)" })" -PercentComplete $entityPct
        Write-MigrationLog "Migracja encji ($entityNum/$entityTotal): $entityName$(if ($pass -gt 1) { " [przebieg $pass]" })"
        $entityMeta = $commonEntities[$entityName]
        $targetAttrNames = @()
        if ($targetMeta -and $targetMeta.ContainsKey($entityName) -and $targetMeta[$entityName].Attributes) {
            $targetAttrNames = @($targetMeta[$entityName].Attributes | ForEach-Object { $_.LogicalName } | Select-Object -Unique)
        }
        try {
            $copyParams = @{
                SourceConn            = $SourceConn
                TargetConn            = $TargetConn
                EntityLogicalName     = $entityName
                EntityMeta            = $entityMeta
                Config                = $Config
                TargetAttributeNames  = $targetAttrNames
                TargetMeta            = $targetMeta
                LogInfo               = ${function:Write-MigrationLog}
            }
            if ($useAutoPull -and $commonEntities -and $orderedEntities) {
                $copyParams['EntitiesInScope'] = $orderedEntities
                $copyParams['CommonEntities'] = $commonEntities
                $copyParams['AutoMigrateMissingLookups'] = $true
            }
            $result = Copy-EntityRecords @copyParams
            if ($pass -eq 1) {
                $stats[$entityName] = $result
            } else {
                $prev = $stats[$entityName]
                $stats[$entityName] = @{
                    Created  = $prev.Created + $result.Created
                    Updated  = $prev.Updated + $result.Updated
                    Skipped  = $prev.Skipped + $result.Skipped
                    Failed   = $result.Failed
                    Total    = $prev.Total
                }
            }
            $upd = if ($null -ne $result.Updated) { $result.Updated } else { 0 }
            Write-MigrationLog "  Zakonczono: utworzono=$($result.Created), zaktualizowano=$upd, pominieto=$($result.Skipped), bledy=$($result.Failed)"
        } catch {
            $msg = $_.Exception.Message
            if ($msg -match "RetrieveMultiple.*does not support entities") {
                Write-MigrationLog "  Encja nie obsluguje pobierania rekordow (RetrieveMultiple) - pomijam."
                if (-not $stats[$entityName]) { $stats[$entityName] = @{ Created = 0; Updated = 0; Skipped = 0; Failed = 0; Total = 0 } }
            } else {
                throw
            }
        }
    }
}
Write-Progress -Activity "Migracja Dataverse" -Completed
$idMapPath = $Config.IdMapPath
if (-not $idMapPath -and $Config.LogFolder) { $idMapPath = Join-Path $Config.LogFolder 'IdMap_latest.json' }
if ($idMapPath) {
    Export-MigrationIdMap -FilePath $idMapPath
    Write-MigrationLog "Zapisano mape ID: $idMapPath"
}
Write-MigrationLog ('Migracja zakonczona. Log: ' + $script:LogFile)

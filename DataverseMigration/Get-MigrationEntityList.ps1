<#
.SYNOPSIS
    Zwraca listę encji do migracji (wspólne dla źródła i celu, w kolejności migracji).
    Używane przez menu i GUI.
#>
param(
    [Parameter(Mandatory = $false)]
    [string] $SourceConnectionString,
    [Parameter(Mandatory = $false)]
    [string] $TargetConnectionString,
    [string] $ConfigPath = "$PSScriptRoot\Config\MigrationConfig.ps1"
)

$ErrorActionPreference = 'Stop'
if (-not $PSScriptRoot -and $MyInvocation.MyCommand.Path) { $PSScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $PSScriptRoot) { $PSScriptRoot = (Get-Location).Path }

# Jesli wywolane z GUI/exe - parametry moga byc w zmiennych srodowiskowych (nowy proces powershell.exe)
if ([string]::IsNullOrWhiteSpace($SourceConnectionString)) { $SourceConnectionString = $env:DATAVERSE_MIGRATION_SRC }
if ([string]::IsNullOrWhiteSpace($TargetConnectionString)) { $TargetConnectionString = $env:DATAVERSE_MIGRATION_TGT }
if ([string]::IsNullOrWhiteSpace($SourceConnectionString) -or [string]::IsNullOrWhiteSpace($TargetConnectionString)) {
    throw "Brak connection stringow. Podaj -SourceConnectionString i -TargetConnectionString lub ustaw DATAVERSE_MIGRATION_SRC i DATAVERSE_MIGRATION_TGT."
}

if (Test-Path $ConfigPath) { $Config = . $ConfigPath }
else {
    $Config = @{
        SystemEntitiesToSkip = @('bulkdeleteoperation','asyncoperation','workflow','pluginassembly')
        EntityOrderPriority = @('systemuser','team','account','contact','lead','opportunity','activitypointer','email','task','appointment')
        BpfEntitySuffix = 'process'
    }
}

# Connection string musi zawierac AuthType= i Url= (pelny connection string, nie sam adres URL)
$connStrCheck = $SourceConnectionString.Trim()
if ($connStrCheck.StartsWith("http://") -or $connStrCheck.StartsWith("https://")) {
    throw "To wyglada na sam adres URL. Potrzebny jest pelny connection string (np. AuthType=OAuth;Url=https://...;AppId=...;RedirectUri=...). Connection string mozesz wygenerowac w Power Platform Admin Center lub przez aplikacje Power Apps."
}
$connStrCheck = $TargetConnectionString.Trim()
if ($connStrCheck.StartsWith("http://") -or $connStrCheck.StartsWith("https://")) {
    throw "Connection string CELU wyglada na sam adres URL. Potrzebny jest pelny connection string (AuthType=OAuth;Url=...;AppId=...;RedirectUri=...)."
}

$libPath = Join-Path $PSScriptRoot 'Lib'
. (Join-Path $libPath 'Connect-Dataverse.ps1')
. (Join-Path $libPath 'Get-EntityMetadata.ps1')
. (Join-Path $libPath 'Get-MigrationOrder.ps1')

$SourceConn = Connect-DataverseEnvironment -ConnectionString $SourceConnectionString
$TargetConn = Connect-DataverseEnvironment -ConnectionString $TargetConnectionString
if ($null -eq $SourceConn) { throw "Polaczenie ze zrodlem nie zostalo nawiazane (Get-CrmConnection zwrocil null). Sprawdz connection string." }
if ($null -eq $TargetConn) { throw "Polaczenie z celem nie zostalo nawiazane (Get-CrmConnection zwrocil null). Sprawdz connection string." }
if (-not (Test-DataverseConnection -Connection $SourceConn)) { throw "Test polaczenia ze zrodlem nie powiodl sie." }
if (-not (Test-DataverseConnection -Connection $TargetConn)) { throw "Test polaczenia z celem nie powiodl sie." }

$sourceMeta = Get-EntityMetadataFromEnv -Connection $SourceConn
$targetMeta = Get-EntityMetadataFromEnv -Connection $TargetConn
$commonEntities = Get-CommonEntities -SourceMetadata $sourceMeta -TargetMetadata $targetMeta -ExcludeEntities $Config.SystemEntitiesToSkip
$orderedEntities = Get-MigrationOrderByDependencies -CommonEntities $commonEntities -Config $Config -BpfSuffix $Config.BpfEntitySuffix

Write-Output $orderedEntities

# Połączenie do Dataverse dla CMT (Microsoft.Xrm.Tooling.ConfigurationMigration)
# CMT używa CrmServiceClient z Microsoft.Xrm.Tooling.Connector

function Install-CMTModuleIfNeeded {
    <#
    .SYNOPSIS
        Instaluje moduł Microsoft.Xrm.Tooling.ConfigurationMigration, jeśli nie jest zainstalowany.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string] $MinimumVersion = '1.0.0.88'
    )
    $mod = Get-Module -ListAvailable -Name 'Microsoft.Xrm.Tooling.ConfigurationMigration' | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $mod) {
        Write-Host "Instalowanie modulu Microsoft.Xrm.Tooling.ConfigurationMigration (MinimumVersion $MinimumVersion)..."
        Install-Module -Name 'Microsoft.Xrm.Tooling.ConfigurationMigration' -MinimumVersion $MinimumVersion -Scope CurrentUser -Force -AllowClobber
        $mod = Get-Module -ListAvailable -Name 'Microsoft.Xrm.Tooling.ConfigurationMigration' | Sort-Object Version -Descending | Select-Object -First 1
    }
    if (-not $mod) {
        throw "Nie udalo sie zainstalowac modulu Microsoft.Xrm.Tooling.ConfigurationMigration. Sprobuj: Install-Module -Name Microsoft.Xrm.Tooling.ConfigurationMigration -Scope CurrentUser"
    }
    if ($mod.Version -lt [version]$MinimumVersion) {
        Write-Host "Aktualizowanie modulu do wersji $MinimumVersion lub nowszej..."
        Update-Module -Name 'Microsoft.Xrm.Tooling.ConfigurationMigration' -Force -ErrorAction SilentlyContinue
    }
    Write-Host "Modul CMT: $($mod.Name) $($mod.Version)"
    return $mod
}

function New-CMTConnection {
    <#
    .SYNOPSIS
        Tworzy połączenie CrmServiceClient używane przez cmdlety CMT (Export-CrmDataFile, Import-CrmDataFile).
    .PARAMETER ConnectionString
        Connection string (np. AuthType=OAuth;Url=...;Username=...;Password=...).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string] $ConnectionString
    )
    if (-not (Get-Module -Name 'Microsoft.Xrm.Tooling.ConfigurationMigration')) {
        Import-Module 'Microsoft.Xrm.Tooling.ConfigurationMigration' -Force -ErrorAction Stop
    }
    # Jawnie zaladuj zestaw Connector (w niektorych hostach modul go nie laduje, co daje "Cannot find type CrmServiceClient")
    $cmtModule = Get-Module -Name 'Microsoft.Xrm.Tooling.ConfigurationMigration'
    $moduleBase = $cmtModule.ModuleBase
    $connectorDll = Join-Path $moduleBase 'Microsoft.Xrm.Tooling.Connector.dll'
    if (Test-Path $connectorDll) {
        $alreadyLoaded = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.Location -and $_.Location.EndsWith('Microsoft.Xrm.Tooling.Connector.dll', [StringComparison]::OrdinalIgnoreCase) }
        if (-not $alreadyLoaded) {
            # Zaladuj zaleznosci w kolejnosci (Connector wymaga Sdk i Proxy)
            foreach ($dll in @('Microsoft.Xrm.Sdk.dll', 'Microsoft.Crm.Sdk.Proxy.dll')) {
                $path = Join-Path $moduleBase $dll
                if (Test-Path $path) {
                    $loaded = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.Location -and (Split-Path $_.Location -Leaf) -eq $dll }
                    if (-not $loaded) { Add-Type -Path $path -ErrorAction SilentlyContinue }
                }
            }
            Add-Type -Path $connectorDll -ErrorAction Stop
        }
    }
    $conn = New-Object Microsoft.Xrm.Tooling.Connector.CrmServiceClient($ConnectionString)
    if (-not $conn.IsReady) {
        $err = $conn.LastCrmError
        throw "Polaczenie CMT nie jest gotowe: $err"
    }
    return $conn
}

function Test-CMTConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $CrmConnection
    )
    if ($null -eq $CrmConnection) { return $false }
    try {
        return $CrmConnection.IsReady -eq $true
    } catch {
        return $false
    }
}

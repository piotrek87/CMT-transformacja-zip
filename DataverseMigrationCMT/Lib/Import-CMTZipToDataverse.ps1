# Importuje zip CMT (po transformacji) do Dataverse przez API – z zachowaniem ownerid i overriddencreatedon.
# Uzycie: .\Import-CMTZipToDataverse.ps1 -ZipPath C:\...\CMT_Export_ForTarget.zip [-ConfigPath ...\CMTConfig.ps1]
# Wymaga: Microsoft.Xrm.Data.PowerShell (Get-CrmConnection, Add-CrmRecord).

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $ZipPath,
    [Parameter(Mandatory = $false)]
    [string] $ConfigPath,
    [Parameter(Mandatory = $false)]
    [string] $TargetConnectionString,
    [Parameter(Mandatory = $false)]
    [switch] $WhatIf,
    [Parameter(Mandatory = $false)]
    [int] $MaxRecordsPerEntity = 0
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop

if (-not (Test-Path $ZipPath -PathType Leaf)) {
    throw "Plik zip nie istnieje: $ZipPath"
}

# Polaczenie do celu
$conn = $null
if (-not [string]::IsNullOrWhiteSpace($TargetConnectionString)) {
    $connStr = $TargetConnectionString
} elseif (-not [string]::IsNullOrWhiteSpace($ConfigPath) -and (Test-Path $ConfigPath)) {
    $cfg = & $ConfigPath
    $connStr = $cfg.TargetConnectionString
} else {
    throw "Podaj -TargetConnectionString lub -ConfigPath z polaczeniem do celu."
}
if ([string]::IsNullOrWhiteSpace($connStr)) {
    throw "Brak connection string do celu (Config\Polaczenia.txt: CelUrl, CelLogin, CelHaslo)."
}

if (-not (Get-Module -ListAvailable -Name 'Microsoft.Xrm.Data.PowerShell')) {
    throw "Zainstaluj modul: Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser"
}
Import-Module Microsoft.Xrm.Data.PowerShell -Force -ErrorAction Stop
$addCmd = Get-Command -Name Add-CrmRecord -ErrorAction SilentlyContinue
if (-not $addCmd) { $addCmd = Get-Command -Name New-CrmRecord -ErrorAction SilentlyContinue }
if (-not $addCmd) {
    throw "Modul nie eksportuje Add-CrmRecord ani New-CrmRecord."
}

try {
    $conn = Get-CrmConnection -ConnectionString $connStr
} catch {
    throw "Polaczenie do celu nie powiodlo sie: $($_.Exception.Message)"
}
if (-not $conn) { throw "Get-CrmConnection zwrocil null." }

# Rozpakuj zip
$tempDir = Join-Path $env:TEMP ([System.Guid]::NewGuid().ToString('N'))
[System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $tempDir)

$dataFile = Get-ChildItem -Path $tempDir -Recurse -Filter 'data.xml' -File -ErrorAction SilentlyContinue | Select-Object -First 1
$schemaFile = Get-ChildItem -Path $tempDir -Recurse -Filter 'data_schema.xml' -File -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $dataFile -or -not (Test-Path $dataFile.FullName)) {
    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    throw "W zipie nie znaleziono data.xml"
}

# Kolejnosc encji i typy pol ze schematu (opcjonalnie)
$entityOrder = [System.Collections.ArrayList]::new()
$schemaLookups = @{}  # entity -> { fieldName -> targetEntity (jedna) }
if ($schemaFile -and (Test-Path $schemaFile.FullName)) {
    try {
        [xml]$schemaXml = [System.IO.File]::ReadAllText($schemaFile.FullName, [System.Text.Encoding]::UTF8)
        $schemaEntities = $schemaXml.SelectNodes('//*[local-name()="entities"]/*[local-name()="entity" and @name]') + $schemaXml.SelectNodes('//*[local-name()="Entities"]/*[local-name()="Entity" and @name]')
        foreach ($se in $schemaEntities) {
            $en = $se.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($en)) { $en = $se.GetAttribute('Name') }
            if (-not [string]::IsNullOrWhiteSpace($en)) { [void]$entityOrder.Add($en) }
            $fieldsParent = $se.SelectSingleNode('.//*[local-name()="fields"]')
            if (-not $fieldsParent) { $fieldsParent = $se.SelectSingleNode('.//*[local-name()="Fields"]') }
            if ($fieldsParent) {
                $lookups = @{}
                foreach ($fn in $fieldsParent.ChildNodes) {
                    if ($fn.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                    $fname = $fn.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($fname)) { $fname = $fn.GetAttribute('Name') }
                    $ftype = $fn.GetAttribute('type'); if ([string]::IsNullOrWhiteSpace($ftype)) { $ftype = $fn.GetAttribute('Type') }
                    $lookupType = $fn.GetAttribute('lookupType'); if ([string]::IsNullOrWhiteSpace($lookupType)) { $lookupType = $fn.GetAttribute('lookupTarget') }
                    if ($ftype -match 'entityreference|lookup' -and -not [string]::IsNullOrWhiteSpace($lookupType)) {
                        $firstTarget = ($lookupType -split '[|,]')[0].Trim()
                        if (-not [string]::IsNullOrWhiteSpace($firstTarget)) { $lookups[$fname] = $firstTarget }
                    }
                }
                if ($lookups.Count -gt 0) { $schemaLookups[$en] = $lookups }
            }
        }
    } catch { Write-Host "Ostrzezenie: nie udalo sie wczytac data_schema.xml: $_" -ForegroundColor Yellow }
}

# Wczytaj data.xml
[xml]$dataXml = [System.IO.File]::ReadAllText($dataFile.FullName, [System.Text.Encoding]::UTF8)
$entityNodes = $dataXml.SelectNodes('//*[local-name()="entity" and @name]') + $dataXml.SelectNodes('//*[local-name()="Entity" and @name]')
if ($entityOrder.Count -eq 0) {
    foreach ($e in $entityNodes) {
        $n = $e.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($n)) { $n = $e.GetAttribute('Name') }
        if (-not [string]::IsNullOrWhiteSpace($n)) { [void]$entityOrder.Add($n) }
    }
}

function Get-PrimaryKeyAttribute {
    param([string]$EntityLogicalName)
    $pk = switch -Regex ($EntityLogicalName) {
        '^(contact|account|lead|opportunity|systemuser|team|businessunit)$' { $EntityLogicalName + 'id' }
        default { $EntityLogicalName + 'id' }
    }
    return $pk
}

function Convert-RecordToFields {
    param(
        [System.Xml.XmlElement] $Record,
        [string] $EntityLogicalName,
        [hashtable] $LookupTargets
    )
    $pkAttr = Get-PrimaryKeyAttribute -EntityLogicalName $EntityLogicalName
    $fields = @{}
    $createdOnVal = $null
    foreach ($child in $Record.ChildNodes) {
        if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
        $cname = $child.GetAttribute('name')
        if ([string]::IsNullOrWhiteSpace($cname)) { $cname = $child.GetAttribute('Name') }
        if ([string]::IsNullOrWhiteSpace($cname)) { $cname = $child.LocalName }
        $cname = $cname.Trim().ToLowerInvariant()
        $val = $child.InnerText
        if ($null -ne $val) { $val = $val.Trim() }

        if ($cname -eq $pkAttr -or $cname -eq 'id') { continue }
        if ($cname -eq 'createdon') { $createdOnVal = $val; continue }

        if ($cname -eq 'ownerid') {
            if (-not [string]::IsNullOrWhiteSpace($val)) {
                $g = [guid]::Empty
                if ([guid]::TryParse($val, [ref]$g)) {
                    try {
                        $fields['ownerid'] = [Microsoft.Xrm.Sdk.EntityReference]::new('systemuser', $g)
                    } catch { Write-Warning "ownerid $val : $_" }
                }
            }
            continue
        }
        if ($cname -eq 'overriddencreatedon') {
            if (-not [string]::IsNullOrWhiteSpace($val)) {
                $dt = [DateTime]::MinValue
                if ([DateTime]::TryParse($val, [ref]$dt)) {
                    $fields['overriddencreatedon'] = $dt
                }
            }
            continue
        }
        if ($cname -eq 'statecode' -or $cname -eq 'statuscode') {
            if (-not [string]::IsNullOrWhiteSpace($val)) {
                $intVal = 0
                if ([int]::TryParse($val, [ref]$intVal)) {
                    $fields[$cname] = [Microsoft.Xrm.Sdk.OptionSetValue]::new($intVal)
                }
            }
            continue
        }
        # Lookup – z schematu lub po konwencji
        if ($cname -match 'id$' -and -not [string]::IsNullOrWhiteSpace($val)) {
            $g = [guid]::Empty
            if ([guid]::TryParse($val, [ref]$g)) {
                $targetEntity = $null
                if ($LookupTargets -and $LookupTargets[$cname]) {
                    $targetEntity = $LookupTargets[$cname]
                } else {
                    $targetEntity = switch -Regex ($cname) {
                        '^parentaccountid$' { 'account' }
                        '^primarycontactid$' { 'contact' }
                        '^accountid$' { 'account' }
                        '^contactid$' { 'contact' }
                        '^ownerid$' { 'systemuser' }
                        '^transactioncurrencyid$' { 'transactioncurrency' }
                        '^businessunitid$' { 'businessunit' }
                        '^originatingleadid$' { 'lead' }
                        '^opportunityid$' { 'opportunity' }
                        '^leadid$' { 'lead' }
                        default { $null }
                    }
                }
                if ($targetEntity) {
                    try {
                        $fields[$cname] = [Microsoft.Xrm.Sdk.EntityReference]::new($targetEntity, $g)
                    } catch { }
                }
                if ($fields.ContainsKey($cname)) { continue }
            }
        }
        # Domyslnie: string (pomin puste jesli nie wymagane)
        if ([string]::IsNullOrWhiteSpace($val)) { continue }
        $fields[$cname] = $val
    }
    if ($null -ne $createdOnVal -and -not $fields.ContainsKey('overriddencreatedon')) {
        $dt = [DateTime]::MinValue
        if ([DateTime]::TryParse($createdOnVal, [ref]$dt)) {
            $fields['overriddencreatedon'] = $dt
        }
    }
    return $fields
}

$totalCreated = 0
$totalErrors = 0
$processedEntities = @{}
$orderIndex = @{}
for ($i = 0; $i -lt $entityOrder.Count; $i++) { $orderIndex[$entityOrder[$i]] = $i }
$dataEntityNames = @($entityNodes | ForEach-Object {
    $n = $_.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($n)) { $n = $_.GetAttribute('Name') }
    $n
} | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
$sortedEntityNames = @($dataEntityNames | Sort-Object {
    if ($orderIndex.ContainsKey($_)) { $orderIndex[$_] } else { 99999 }
})

foreach ($entName in $sortedEntityNames) {
    $entNode = $entityNodes | Where-Object {
        $n = $_.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($n)) { $n = $_.GetAttribute('Name') }
        $n -eq $entName
    } | Select-Object -First 1
    if (-not $entNode) { continue }

    $recordNodes = $entNode.SelectNodes('.//*[local-name()="record"]') + $entNode.SelectNodes('.//*[local-name()="Record"]')
    $lookups = if ($schemaLookups[$entName]) { $schemaLookups[$entName] } else { @{} }
    $count = 0
    $limit = if ($MaxRecordsPerEntity -gt 0) { $MaxRecordsPerEntity } else { [int]::MaxValue }
    Write-Host "Encja: $entName (rekordow: $($recordNodes.Count))" -ForegroundColor Cyan
    foreach ($rec in $recordNodes) {
        if ($count -ge $limit) { break }
        $fields = Convert-RecordToFields -Record $rec -EntityLogicalName $entName -LookupTargets $lookups
        if ($fields.Count -eq 0) { continue }
        if ($WhatIf) {
            Write-Host "  [WhatIf] Utworzylbym rekord $entName z polami: $($fields.Keys -join ', ')" -ForegroundColor Gray
            $count++
            $totalCreated++
            continue
        }
        try {
            $newId = & $addCmd -conn $conn -EntityLogicalName $entName -Fields $fields
            $count++
            $totalCreated++
            if ($totalCreated % 50 -eq 0) { Write-Host "  Utworzono $totalCreated rekordow..." -ForegroundColor Gray }
        } catch {
            $totalErrors++
            Write-Warning "  Rekord $entName : $($_.Exception.Message)"
        }
    }
    $processedEntities[$entName] = $count
}

Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Zakonczono. Utworzono: $totalCreated, bledow: $totalErrors." -ForegroundColor $(if ($totalErrors -gt 0) { 'Yellow' } else { 'Green' })

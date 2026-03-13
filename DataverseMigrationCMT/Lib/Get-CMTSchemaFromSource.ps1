# Pobiera metadane ze zrodla (Dataverse) i generuje plik schematu CMT (data_schema.xml).
# Wymaga: Microsoft.Xrm.Data.PowerShell (do pobrania metadanych). Connection string jak w CMT (zrodlo).
# Logi: zapis do pliku Logs\CMT_Schema_*.log oraz na konsole (pomaga ustalic problem z pustymi atrybutami).

$script:SchemaLogPath = $null

function Write-SchemaLog {
    param([string]$Message, [string]$Level = 'INFO')
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] $Message"
    Write-Host $line
    if ($script:SchemaLogPath -and (Test-Path (Split-Path $script:SchemaLogPath -Parent))) {
        try {
            Add-Content -LiteralPath $script:SchemaLogPath -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
        } catch { Write-Host ('Nie mozna zapisac do logu: ' + $_.Exception.Message) -ForegroundColor Yellow }
    }
}

function Get-SourceMetadataForCMTSchema {
    param(
        [Parameter(Mandatory = $true)]
        [string] $ConnectionString,
        [Parameter(Mandatory = $false)]
        [string[]] $EntityFilter = $null
    )
    $filterInfo = if ($EntityFilter -and $EntityFilter.Count -gt 0) { ($EntityFilter[0..2] -join ',') } else { 'wszystkie' }
    Write-SchemaLog ('Get-SourceMetadataForCMTSchema: start. EntityFilter count=' + $EntityFilter.Count + ', pierwsze: ' + $filterInfo)
    if (-not (Get-Module -ListAvailable -Name 'Microsoft.Xrm.Data.PowerShell')) {
        throw "Do pobrania schematu ze zrodla potrzebny jest modul Microsoft.Xrm.Data.PowerShell. Zainstaluj: Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser"
    }
    Import-Module Microsoft.Xrm.Data.PowerShell -Force -ErrorAction Stop
    Write-SchemaLog 'Polaczenie ze zrodlem (Get-CrmConnection)...'
    $conn = Get-CrmConnection -ConnectionString $ConnectionString -ErrorAction Stop
    Write-SchemaLog 'Polaczenie OK.'
    # EntityFilters::Attributes – zeby Get-CrmEntityAllMetadata zwrocil encje Z atrybutami (bez tego czesto Attributes = null)
    $entityFiltersAttributes = $null
    try {
        $entityFiltersAttributes = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::Attributes
        Write-SchemaLog 'Wywolanie Get-CrmEntityAllMetadata z EntityFilters::Attributes (pelne atrybuty)...'
    } catch {
        Write-SchemaLog ('Brak typu EntityFilters (uzycie Get-CrmEntityAllMetadata bez EntityFilters): ' + $_.Exception.Message) -Level WARN
    }
    $entities = @{}
    try {
        if ($null -ne $entityFiltersAttributes) {
            try {
                $raw = Get-CrmEntityAllMetadata -conn $conn -EntityFilters $entityFiltersAttributes -ErrorAction Stop
            } catch {
                Write-SchemaLog ('Get-CrmEntityAllMetadata z EntityFilters nie dzialal, proba bez: ' + $_.Exception.Message) -Level WARN
                $raw = Get-CrmEntityAllMetadata -conn $conn -ErrorAction Stop
            }
        } else {
            $raw = Get-CrmEntityAllMetadata -conn $conn -ErrorAction Stop
        }
        $rawType = if ($null -eq $raw) { 'null' } else { $raw.GetType().FullName }
        $allMeta = @()
        if ($null -eq $raw) { $allMeta = @() }
        elseif ($raw -is [Array]) { $allMeta = $raw }
        elseif ($raw.PSObject.Properties.Name -contains 'EntityMetadata') { $allMeta = @($raw.EntityMetadata) }
        else { $allMeta = @($raw) }
        Write-SchemaLog ('Get-CrmEntityAllMetadata zwrocil: rawType=' + $rawType + ', allMeta.Count=' + $allMeta.Count)
        if ($allMeta.Count -gt 0) {
            $first = $allMeta[0]
            $firstProps = ($first.PSObject.Properties.Name | Sort-Object) -join ','
            $firstAttrs = $first.Attributes
            $firstAttrsCount = if ($null -eq $firstAttrs) { 'null' } elseif ($firstAttrs -is [Array]) { $firstAttrs.Count } else { 'nie-tablica' }
            Write-SchemaLog ('Pierwsza encja: LogicalName=' + $first.LogicalName + ', wlasciwosci=[' + $firstProps + '], Attributes=' + $firstAttrsCount)
        }
        foreach ($entityMeta in $allMeta) {
            if ($null -eq $entityMeta -or [string]::IsNullOrWhiteSpace($entityMeta.LogicalName)) { continue }
            $logicalName = $entityMeta.LogicalName
            if ($EntityFilter -and $EntityFilter.Count -gt 0 -and $logicalName -notin $EntityFilter) { continue }
            $attrs = @()
            $attrList = $entityMeta.Attributes
            if ($null -eq $attrList) { $attrList = @() }
            foreach ($attr in $attrList) {
                if ($attr.IsLogical -eq $false -and ($attr.IsValidForCreate -eq $true -or $attr.IsValidForUpdate -eq $true)) {
                    $typeVal = $attr.AttributeType; if ($null -ne $typeVal -and $typeVal -isnot [string]) { $typeVal = [string]$typeVal }
                    $targets = @()
                    if ($typeVal -match '^(Lookup|Customer|Owner)$' -and $null -ne $attr) {
                        if ($attr.Targets -and $attr.Targets.Count -gt 0) { $targets = @($attr.Targets) }
                        elseif ($attr.TargetEntityLogicalName) { $targets = @($attr.TargetEntityLogicalName) }
                    }
                    $attrs += @{ LogicalName = $attr.LogicalName; Type = $typeVal; Targets = $targets }
                }
            }
            if ($attrs.Count -eq 0) {
                foreach ($attr in $attrList) {
                    if ($attr.IsLogical -eq $false) {
                        $typeVal = $attr.AttributeType; if ($null -ne $typeVal -and $typeVal -isnot [string]) { $typeVal = [string]$typeVal }
                        $targets = @()
                        if ($typeVal -match '^(Lookup|Customer|Owner)$' -and $null -ne $attr) {
                            if ($attr.Targets -and $attr.Targets.Count -gt 0) { $targets = @($attr.Targets) }
                            elseif ($attr.TargetEntityLogicalName) { $targets = @($attr.TargetEntityLogicalName) }
                        }
                        $attrs += @{ LogicalName = $attr.LogicalName; Type = $typeVal; Targets = $targets }
                    }
                }
            }
            $pkAttr = if ($entityMeta.PrimaryIdAttribute) { $entityMeta.PrimaryIdAttribute } else { $logicalName + 'id' }
            $pnAttr = if ($entityMeta.PrimaryNameAttribute) { $entityMeta.PrimaryNameAttribute } else { 'name' }
            $otc = if ($null -ne $entityMeta.ObjectTypeCode) { $entityMeta.ObjectTypeCode } else { '' }
            $entities[$logicalName] = @{ LogicalName = $logicalName; Attributes = $attrs; PrimaryIdAttribute = $pkAttr; PrimaryNameAttribute = $pnAttr; ObjectTypeCode = $otc }
        }
        $withAttrsNow = @($entities.Keys | Where-Object { $entities[$_].Attributes -and $entities[$_].Attributes.Count -gt 0 })
        Write-SchemaLog ('Po AllMetadata: encje=' + $entities.Count + ', z atrybutami=' + $withAttrsNow.Count)
        # Gdy encje maja puste atrybuty (czesto Get-CrmEntityAllMetadata nie zwraca Attributes) – pobierz pojedynczo
        $needAttributes = @($entities.Keys | Where-Object { -not $entities[$_].Attributes -or $entities[$_].Attributes.Count -eq 0 })
        if ($needAttributes.Count -gt 0) {
            Write-SchemaLog ('Pobieranie atrybutow pojedynczo (Get-CrmEntityMetadata) dla ' + $needAttributes.Count + ' encji...')
            $entityFiltersAttr = $null
            try {
                $entityFiltersAttr = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::Attributes
                Write-SchemaLog 'Uzycie EntityFilters::Attributes przy Get-CrmEntityMetadata (pelne atrybuty).'
            } catch {
                Write-SchemaLog ('Brak typu EntityFilters w zakresie (uzywam domyslnego wywolania): ' + $_.Exception.Message) -Level WARN
            }
            foreach ($logicalName in $needAttributes) {
                try {
                    if ($null -ne $entityFiltersAttr) {
                        try {
                            $meta = Get-CrmEntityMetadata -Conn $conn -EntityLogicalName $logicalName -EntityFilters $entityFiltersAttr -ErrorAction Stop
                        } catch {
                            Write-SchemaLog ('  Get-CrmEntityMetadata z EntityFilters nie dzialal, proba bez: ' + $_.Exception.Message) -Level WARN
                            $entityFiltersAttr = $null
                            $meta = Get-CrmEntityMetadata -Conn $conn -EntityLogicalName $logicalName -ErrorAction Stop
                        }
                    } else {
                        $meta = Get-CrmEntityMetadata -Conn $conn -EntityLogicalName $logicalName -ErrorAction Stop
                    }
                    $attrCount = if ($meta -and $meta.Attributes) { $meta.Attributes.Count } else { 0 }
                    Write-SchemaLog ('  Get-CrmEntityMetadata ' + $logicalName + ' -> Attributes.Count=' + $attrCount)
                    $pkA = if ($meta -and $meta.PrimaryIdAttribute) { $meta.PrimaryIdAttribute } else { $logicalName + 'id' }
                    $pnA = if ($meta -and $meta.PrimaryNameAttribute) { $meta.PrimaryNameAttribute } else { 'name' }
                    $otc = if ($meta -and $null -ne $meta.ObjectTypeCode) { $meta.ObjectTypeCode } else { '' }
                    if ($meta -and $meta.Attributes) {
                        $attrs = @()
                        foreach ($attr in $meta.Attributes) {
                            if ($attr.IsLogical -eq $false) {
                                $typeVal = $attr.AttributeType; if ($null -ne $typeVal -and $typeVal -isnot [string]) { $typeVal = [string]$typeVal }
                                $targets = @()
                                if ($typeVal -match '^(Lookup|Customer|Owner)$' -and $null -ne $attr) {
                                    if ($attr.Targets -and $attr.Targets.Count -gt 0) { $targets = @($attr.Targets) }
                                    elseif ($attr.TargetEntityLogicalName) { $targets = @($attr.TargetEntityLogicalName) }
                                }
                                $attrs += @{ LogicalName = $attr.LogicalName; Type = $typeVal; Targets = $targets }
                            }
                        }
                        if ($attrs.Count -gt 0) {
                            $entities[$logicalName] = @{ LogicalName = $logicalName; Attributes = $attrs; PrimaryIdAttribute = $pkA; PrimaryNameAttribute = $pnA; ObjectTypeCode = $otc }
                        } else {
                            $entities[$logicalName] = @{ LogicalName = $logicalName; Attributes = @(); PrimaryIdAttribute = $pkA; PrimaryNameAttribute = $pnA; ObjectTypeCode = $otc }
                        }
                    } else {
                        $entities[$logicalName] = @{ LogicalName = $logicalName; Attributes = @(); PrimaryIdAttribute = $pkA; PrimaryNameAttribute = $pnA; ObjectTypeCode = $otc }
                    }
                } catch {
                    Write-SchemaLog ('  ' + $logicalName + ' blad: ' + $_.Exception.Message) -Level WARN
                    $entities[$logicalName] = @{ LogicalName = $logicalName; Attributes = @(); PrimaryIdAttribute = ($logicalName + 'id'); PrimaryNameAttribute = 'name'; ObjectTypeCode = '' }
                }
            }
        }
    } catch {
        Write-SchemaLog ('Get-CrmEntityAllMetadata wyjatek: ' + $_.Exception.Message) -Level WARN
        Write-SchemaLog 'Proba pojedynczych encji (Get-CrmEntityMetadata)...'
        $entities = @{}
    }
    # Encje z listy (EntityFilter), ktorych nie mamy lub maja puste atrybuty – dopobierz pojedynczo
    if ($EntityFilter -and $EntityFilter.Count -gt 0) {
        $missing = @()
        foreach ($name in $EntityFilter) {
            if (-not $entities.ContainsKey($name) -or -not $entities[$name].Attributes -or $entities[$name].Attributes.Count -eq 0) { $missing += $name }
        }
        if ($missing.Count -gt 0) {
            Write-SchemaLog ('Dopobieranie z listy (EntityFilter): ' + $missing.Count + ' encji - ' + ($missing -join ', '))
            $entityFiltersAttr = $null
            try { $entityFiltersAttr = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::Attributes } catch { }
            foreach ($logicalName in $missing) {
                try {
                    if ($null -ne $entityFiltersAttr) {
                        try {
                            $meta = Get-CrmEntityMetadata -Conn $conn -EntityLogicalName $logicalName -EntityFilters $entityFiltersAttr -ErrorAction Stop
                        } catch {
                            $entityFiltersAttr = $null
                            $meta = Get-CrmEntityMetadata -Conn $conn -EntityLogicalName $logicalName -ErrorAction Stop
                        }
                    } else {
                        $meta = Get-CrmEntityMetadata -Conn $conn -EntityLogicalName $logicalName -ErrorAction Stop
                    }
                    $attrCount = if ($meta -and $meta.Attributes) { $meta.Attributes.Count } else { 0 }
                    Write-SchemaLog ('  ' + $logicalName + ' -> Attributes.Count=' + $attrCount)
                    $pkA = if ($meta -and $meta.PrimaryIdAttribute) { $meta.PrimaryIdAttribute } else { $logicalName + 'id' }
                    $pnA = if ($meta -and $meta.PrimaryNameAttribute) { $meta.PrimaryNameAttribute } else { 'name' }
                    $otc = if ($meta -and $null -ne $meta.ObjectTypeCode) { $meta.ObjectTypeCode } else { '' }
                    if ($meta -and $meta.Attributes) {
                        $attrs = @()
                        foreach ($attr in $meta.Attributes) {
                            if ($attr.IsLogical -eq $false) {
                                $typeVal = $attr.AttributeType; if ($null -ne $typeVal -and $typeVal -isnot [string]) { $typeVal = [string]$typeVal }
                                $targets = @()
                                if ($typeVal -match '^(Lookup|Customer|Owner)$' -and $null -ne $attr) {
                                    if ($attr.Targets -and $attr.Targets.Count -gt 0) { $targets = @($attr.Targets) }
                                    elseif ($attr.TargetEntityLogicalName) { $targets = @($attr.TargetEntityLogicalName) }
                                }
                                $attrs += @{ LogicalName = $attr.LogicalName; Type = $typeVal; Targets = $targets }
                            }
                        }
                        if ($attrs.Count -gt 0) {
                            $entities[$logicalName] = @{ LogicalName = $logicalName; Attributes = $attrs; PrimaryIdAttribute = $pkA; PrimaryNameAttribute = $pnA; ObjectTypeCode = $otc }
                        } else {
                            $entities[$logicalName] = @{ LogicalName = $logicalName; Attributes = @(); PrimaryIdAttribute = $pkA; PrimaryNameAttribute = $pnA; ObjectTypeCode = $otc }
                        }
                    } else {
                        $entities[$logicalName] = @{ LogicalName = $logicalName; Attributes = @(); PrimaryIdAttribute = $pkA; PrimaryNameAttribute = $pnA; ObjectTypeCode = $otc }
                    }
                } catch {
                    Write-SchemaLog ('  ' + $logicalName + ' blad: ' + $_.Exception.Message) -Level WARN
                    $entities[$logicalName] = @{ LogicalName = $logicalName; Attributes = @(); PrimaryIdAttribute = ($logicalName + 'id'); PrimaryNameAttribute = 'name'; ObjectTypeCode = '' }
                }
            }
        }
    }
    # Jesli nadal brak encji z atrybutami – proba listy domyslnej
    $withAttrs = @($entities.Keys | Where-Object { $entities[$_].Attributes -and $entities[$_].Attributes.Count -gt 0 })
    if ($withAttrs.Count -eq 0) {
        Write-SchemaLog 'Zadna encja nie ma atrybutow - proba listy domyslnej (Get-CrmEntityMetadata pojedynczo)...'
        $defaultEntities = if ($EntityFilter -and $EntityFilter.Count -gt 0) {
            @($EntityFilter)
        } else {
            @('account','contact','lead','opportunity','systemuser','team','businessunit','transactioncurrency','subject','activitypointer','email','task','appointment','phonecall','letter','fax','annotation')
        }
        $entityFiltersAttr = $null
        try { $entityFiltersAttr = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::Attributes } catch { }
        foreach ($logicalName in $defaultEntities) {
            if ($EntityFilter -and $EntityFilter.Count -gt 0 -and $logicalName -notin $EntityFilter) { continue }
            try {
                if ($null -ne $entityFiltersAttr) {
                    try {
                        $meta = Get-CrmEntityMetadata -Conn $conn -EntityLogicalName $logicalName -EntityFilters $entityFiltersAttr -ErrorAction Stop
                    } catch {
                        $entityFiltersAttr = $null
                        $meta = Get-CrmEntityMetadata -Conn $conn -EntityLogicalName $logicalName -ErrorAction Stop
                    }
                } else {
                    $meta = Get-CrmEntityMetadata -Conn $conn -EntityLogicalName $logicalName -ErrorAction Stop
                }
                $attrCount = if ($meta -and $meta.Attributes) { $meta.Attributes.Count } else { 0 }
                Write-SchemaLog ('  ' + $logicalName + ' -> Attributes.Count=' + $attrCount)
                $pkA = if ($meta -and $meta.PrimaryIdAttribute) { $meta.PrimaryIdAttribute } else { $logicalName + 'id' }
                $pnA = if ($meta -and $meta.PrimaryNameAttribute) { $meta.PrimaryNameAttribute } else { 'name' }
                $otc = if ($meta -and $null -ne $meta.ObjectTypeCode) { $meta.ObjectTypeCode } else { '' }
                if ($meta -and $meta.Attributes) {
                    $attrs = @()
                    foreach ($attr in $meta.Attributes) {
                        if ($attr.IsLogical -eq $false) {
                            $typeVal = $attr.AttributeType; if ($null -ne $typeVal -and $typeVal -isnot [string]) { $typeVal = [string]$typeVal }
                            $targets = @()
                            if ($typeVal -match '^(Lookup|Customer|Owner)$' -and $null -ne $attr) {
                                if ($attr.Targets -and $attr.Targets.Count -gt 0) { $targets = @($attr.Targets) }
                                elseif ($attr.TargetEntityLogicalName) { $targets = @($attr.TargetEntityLogicalName) }
                            }
                            $attrs += @{ LogicalName = $attr.LogicalName; Type = $typeVal; Targets = $targets }
                        }
                    }
                    if ($attrs.Count -gt 0) {
                        $entities[$logicalName] = @{ LogicalName = $logicalName; Attributes = $attrs; PrimaryIdAttribute = $pkA; PrimaryNameAttribute = $pnA; ObjectTypeCode = $otc }
                    } else {
                        $entities[$logicalName] = @{ LogicalName = $logicalName; Attributes = @(); PrimaryIdAttribute = $pkA; PrimaryNameAttribute = $pnA; ObjectTypeCode = $otc }
                    }
                } else {
                    $entities[$logicalName] = @{ LogicalName = $logicalName; Attributes = @(); PrimaryIdAttribute = $pkA; PrimaryNameAttribute = $pnA; ObjectTypeCode = $otc }
                }
            } catch {
                Write-SchemaLog ('  ' + $logicalName + ' blad: ' + $_.Exception.Message) -Level WARN
                $entities[$logicalName] = @{ LogicalName = $logicalName; Attributes = @(); PrimaryIdAttribute = ($logicalName + 'id'); PrimaryNameAttribute = 'name'; ObjectTypeCode = '' }
            }
        }
    }
    $finalWithAttrs = @($entities.Keys | Where-Object { $entities[$_].Attributes -and $entities[$_].Attributes.Count -gt 0 })
    $finalWithoutAttrs = @($entities.Keys | Where-Object { -not $entities[$_].Attributes -or $entities[$_].Attributes.Count -eq 0 })
    Write-SchemaLog ('Koniec Get-SourceMetadataForCMTSchema: encje lacznie=' + $entities.Count + ', z atrybutami=' + $finalWithAttrs.Count + ', bez atrybutow (syntetyczny PK w XML)=' + $finalWithoutAttrs.Count)
    if ($finalWithoutAttrs.Count -gt 0 -and $finalWithoutAttrs.Count -le 20) {
        Write-SchemaLog ('Encje bez atrybutow: ' + ($finalWithoutAttrs -join ', '))
    } elseif ($finalWithoutAttrs.Count -gt 20) {
        Write-SchemaLog ('Encje bez atrybutow (pierwsze 20): ' + (($finalWithoutAttrs | Select-Object -First 20) -join ', ') + '...')
    }
    return $entities
}

function Get-CommonEntitiesForCMT {
    <#
    .SYNOPSIS
        Zwraca encje istniejace w obu srodowiskach oraz atrybuty obecne w ZAROWNO w zrodle, jak i w celu.
        CMT eksportuje ze zrodla – schemat musi zawierac tylko pola, ktore zrodlo ma (w przeciwnym razie "Export failed" / schema validation).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $SourceMetadata,
        [Parameter(Mandatory = $true)]
        $TargetMetadata,
        [Parameter(Mandatory = $false)]
        [string[]] $ExcludeEntities = @()
    )
    if ($null -eq $SourceMetadata) { $SourceMetadata = @{} }
    if ($null -eq $TargetMetadata) { $TargetMetadata = @{} }
    $common = @{}
    foreach ($name in $SourceMetadata.Keys) {
        if ($name -in $ExcludeEntities) { continue }
        if (-not $TargetMetadata.ContainsKey($name)) { continue }
        $srcAttrs = $SourceMetadata[$name].Attributes
        $tgtAttrs = $TargetMetadata[$name].Attributes
        if ($null -eq $srcAttrs) { $srcAttrs = @() }
        if ($null -eq $tgtAttrs) { $tgtAttrs = @() }
        $srcAttrNames = @{}
        $srcAttrByLn = @{}
        foreach ($a in $srcAttrs) {
            $srcAttrNames[[string]$a.LogicalName] = $true
            $srcAttrByLn[[string]$a.LogicalName] = $a
        }
        # Tylko atrybuty istniejace w zrodle I w celu (z typem z celu; Targets z zrodla dla lookupow)
        $commonAttrs = @()
        foreach ($t in $tgtAttrs) {
            if (-not $srcAttrNames.ContainsKey([string]$t.LogicalName)) { continue }
            $srcA = $srcAttrByLn[[string]$t.LogicalName]
            $entry = @{ LogicalName = $t.LogicalName; Type = $t.Type }
            if ($srcA -and $srcA.Targets -and $srcA.Targets.Count -gt 0) { $entry.Targets = $srcA.Targets }
            $commonAttrs += $entry
        }
        $srcEnt = $SourceMetadata[$name]
        $pkA = if ($srcEnt.PrimaryIdAttribute) { $srcEnt.PrimaryIdAttribute } else { $name + 'id' }
        $pnA = if ($srcEnt.PrimaryNameAttribute) { $srcEnt.PrimaryNameAttribute } else { 'name' }
        $otc = if ($null -ne $srcEnt.ObjectTypeCode) { $srcEnt.ObjectTypeCode } else { '' }
        $common[$name] = @{ LogicalName = $name; Attributes = $commonAttrs; PrimaryIdAttribute = $pkA; PrimaryNameAttribute = $pnA; ObjectTypeCode = $otc }
    }
    return $common
}

function Get-CMTFieldType {
    param([string]$Type)
    if ([string]::IsNullOrWhiteSpace($Type)) { return 'string' }
    $t = $Type.Trim()
    if ($t -eq 'Uniqueidentifier' -or $t -eq 'Guid') { return 'guid' }
    if ($t -eq 'Lookup' -or $t -eq 'Customer' -or $t -eq 'Owner') { return 'entityreference' }
    if ($t -eq 'String' -or $t -eq 'Memo') { return 'string' }
    if ($t -eq 'DateTime') { return 'datetime' }
    if ($t -eq 'Integer' -or $t -eq 'Double' -or $t -eq 'Int') { return 'number' }
    if ($t -eq 'Boolean') { return 'bool' }
    if ($t -eq 'Decimal' -or $t -eq 'Money') { return 'decimal' }
    if ($t -eq 'OptionSetValue' -or $t -eq 'Picklist') { return 'optionsetvalue' }
    if ($t -eq 'State') { return 'state' }
    if ($t -eq 'Status') { return 'status' }
    return 'string'
}

function Export-CMTSchemaXml {
    <#
    .SYNOPSIS
        Generuje plik data_schema.xml w formacie CMT (entities/entity/fields/field jak w narzedziu GUI).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable] $EntityMetadata,
        [Parameter(Mandatory = $true)]
        [string] $OutputPath,
        [Parameter(Mandatory = $false)]
        [string[]] $EntityOrder = $null
    )
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine('<?xml version="1.0" encoding="utf-8"?>')
    [void]$sb.AppendLine('<entities>')
    $ordered = if ($EntityOrder -and $EntityOrder.Count -gt 0) {
        $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($e in $EntityOrder) { [void]$set.Add($e) }
        $first = @($EntityOrder | Where-Object { $EntityMetadata.ContainsKey($_) })
        $rest = @($EntityMetadata.Keys | Where-Object { -not $set.Contains($_) } | Sort-Object)
        $first + $rest
    } else {
        @($EntityMetadata.Keys | Sort-Object)
    }
    $writtenWithAttrs = 0
    $writtenSyntheticPk = 0
    foreach ($entityName in $ordered) {
        $ent = $EntityMetadata[$entityName]
        $attrs = $ent.Attributes
        if (-not $attrs -or $attrs.Count -eq 0) {
            $pkAttr = $entityName + 'id'
            $attrs = @(@{ LogicalName = $pkAttr; Type = 'Uniqueidentifier' })
            $writtenSyntheticPk++
        } else {
            $writtenWithAttrs++
        }
        $pkAttr = if ($ent.PrimaryIdAttribute) { $ent.PrimaryIdAttribute } else { $entityName + 'id' }
        $primaryName = if ($ent.PrimaryNameAttribute) { $ent.PrimaryNameAttribute } else { 'name' }
        $etc = if ($null -ne $ent.ObjectTypeCode -and $ent.ObjectTypeCode -ne '') { $ent.ObjectTypeCode } else { '' }
        $etcAttr = if ([string]::IsNullOrWhiteSpace($etc)) { '' } else { " etc=`"$etc`"" }
        [void]$sb.AppendLine("  <entity name=`"$entityName`" displayname=`"$entityName`"$etcAttr primaryidfield=`"$pkAttr`" primarynamefield=`"$primaryName`" disableplugins=`"false`">")
        [void]$sb.AppendLine('    <fields>')
        foreach ($a in $attrs) {
            $ln = $a.LogicalName
            $type = $a.Type
            $isPk = ($ln -eq $pkAttr)
            $cmtType = Get-CMTFieldType -Type $type
            $typeAttr = " type=`"$cmtType`""
            $lookupAttr = ''
            if ($cmtType -eq 'entityreference' -and $a.Targets -and $a.Targets.Count -gt 0) {
                $lookupAttr = " lookupType=`"" + ($a.Targets -join '|') + "`""
            }
            $updCompare = if ($isPk) { ' updateCompare="true"' } else { '' }
            $pkAttrField = if ($isPk) { ' primaryKey="true"' } else { '' }
            [void]$sb.AppendLine("      <field displayname=`"$ln`" name=`"$ln`"$typeAttr$lookupAttr$updCompare$pkAttrField />")
        }
        [void]$sb.AppendLine('    </fields>')
        [void]$sb.AppendLine('    <relationships />')
        [void]$sb.AppendLine('  </entity>')
    }
    [void]$sb.AppendLine('</entities>')
    $dir = [System.IO.Path]::GetDirectoryName($OutputPath)
    if (-not [string]::IsNullOrEmpty($dir) -and -not [System.IO.Directory]::Exists($dir)) {
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
    }
    [System.IO.File]::WriteAllText($OutputPath, $sb.ToString(), [System.Text.UTF8Encoding]::new($false))
    if ($script:SchemaLogPath) {
        Write-SchemaLog ('Export-CMTSchemaXml: zapisano ' + $ordered.Count + ' encji (z atrybutami: ' + $writtenWithAttrs + ', tylko syntetyczny PK: ' + $writtenSyntheticPk + ') -> ' + $OutputPath)
    }
}

function Get-CMTSchemaFromSource {
    <#
    .SYNOPSIS
        Laczy sie ze zrodlem, pobiera metadane encji i zapisuje plik schematu CMT (data_schema.xml).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string] $ConnectionString,
        [Parameter(Mandatory = $false)]
        [string] $ConfigPath,
        [Parameter(Mandatory = $false)]
        [string] $OutputPath,
        [Parameter(Mandatory = $false)]
        [string[]] $EntityFilter = $null
    )
    $scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
    $logsDir = Join-Path $scriptRoot '..\Logs'
    if (-not [System.IO.Directory]::Exists($logsDir)) { [System.IO.Directory]::CreateDirectory($logsDir) | Out-Null }
    $script:SchemaLogPath = Join-Path $logsDir ("CMT_Schema_{0:yyyyMMdd_HHmmss}.log" -f [DateTime]::Now)
    Write-SchemaLog ('Get-CMTSchemaFromSource: start. ConfigPath=' + $ConfigPath)
    $configDir = Join-Path $scriptRoot '..\Config'
    if (-not $ConfigPath) { $ConfigPath = Join-Path $configDir 'CMTConfig.ps1' }
    if (-not (Test-Path $ConfigPath)) { throw "Brak configu: $ConfigPath" }
    $config = & $ConfigPath
    $connOk = if ($config.SourceConnectionString) { 'tak' } else { 'nie' }
    Write-SchemaLog ('Config zaladowany. SourceConnectionString obecny: ' + $connOk)
    $connStr = if ($ConnectionString) { $ConnectionString } else { $config.SourceConnectionString }
    if ([string]::IsNullOrWhiteSpace($connStr)) { throw "Brak SourceConnectionString w configu lub -ConnectionString." }
    if (-not $OutputPath) {
        $outDir = $config.ExportOutputDirectory
        if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = Join-Path $scriptRoot '..\Output' }
        $OutputPath = Join-Path $outDir 'data_schema_generated.xml'
    }
    Write-SchemaLog ('OutputPath=' + $OutputPath)
    $entityFilter = $EntityFilter
    if (-not $entityFilter -and $config.SchemaEntityIncludeOnly -and $config.SchemaEntityIncludeOnly.Count -gt 0) {
        $entityFilter = @($config.SchemaEntityIncludeOnly)
        Write-SchemaLog ('Schemat ograniczony do ' + $entityFilter.Count + ' encji z configu (SchemaEntityIncludeOnly).')
    }
    Write-SchemaLog 'Polaczenie ze zrodlem i pobieranie metadanych (Get-SourceMetadataForCMTSchema)...'
    $meta = Get-SourceMetadataForCMTSchema -ConnectionString $connStr -EntityFilter $entityFilter
    if ($meta.Count -eq 0) { throw "Nie pobrano zadnych encji. Sprawdz polaczenie, uprawnienia i liste encji w configu. Log: $script:SchemaLogPath" }
    Write-SchemaLog ('Pobrano encje: ' + $meta.Count + '. Eksport do ' + $OutputPath + ' (Export-CMTSchemaXml)...')
    $entityOrder = $null
    if ($config.SchemaEntityOrder -and $config.SchemaEntityOrder.Count -gt 0) {
        $entityOrder = @($config.SchemaEntityOrder)
    }
    Export-CMTSchemaXml -EntityMetadata $meta -OutputPath $OutputPath -EntityOrder $entityOrder
    Write-SchemaLog ('Schemat zapisany: ' + $OutputPath + '. Pelny log: ' + $script:SchemaLogPath)
    return $OutputPath
}

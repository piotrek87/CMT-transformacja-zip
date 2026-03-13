# Moduł migracji danych encji z retry, paging, mapowaniem lookup
# Używa: overriddencreatedon, statecode, statuscode, mapowanie ownerid

# W Dataverse encje aktywności (email, task, appointment, …) mają PK activityid, nie entitynameid
$script:ActivityEntityNames = @{ 'activitypointer' = $true; 'email' = $true; 'task' = $true; 'appointment' = $true; 'phonecall' = $true; 'letter' = $true; 'fax' = $true; 'campaignresponse' = $true; 'campaignactivity' = $true }

$script:LookupIdMap = @{}  # klucz: "entityname|sourceId", wartość: targetId (GUID)
$script:SkipFields = @()
# Stos encji|id przy domigrowaniu brakujacych lookupow (ochrona przed petla)
$script:AutoMigratePullStack = [System.Collections.ArrayList]::new()
# Niektore wersje/forki modulu eksportuja New-CrmRecord zamiast Add-CrmRecord
$script:NewCrmRecordCmd = Get-Command -Name Add-CrmRecord -ErrorAction SilentlyContinue
if (-not $script:NewCrmRecordCmd) { $script:NewCrmRecordCmd = Get-Command -Name New-CrmRecord -ErrorAction SilentlyContinue }

function Get-PrimaryKeyAttributeName {
    param([string] $EntityLogicalName)
    if ($script:ActivityEntityNames -and $script:ActivityEntityNames.ContainsKey($EntityLogicalName)) { return 'activityid' }
    return $EntityLogicalName + 'id'
}

function Initialize-LookupMap {
    param([string] $EntityLogicalName, [guid] $SourceId, [guid] $TargetId)
    $key = "${EntityLogicalName}|$SourceId"
    $script:LookupIdMap[$key] = $TargetId
}

function Get-MappedLookupId {
    param([string] $EntityLogicalName, [guid] $SourceId)
    $key = "${EntityLogicalName}|$SourceId"
    if ($script:LookupIdMap.ContainsKey($key)) {
        return $script:LookupIdMap[$key]
    }
    return $null
}

# Zapis mapy source->target do pliku (po migracji); odczyt przy MatchBy Id/IdThenName
function Export-MigrationIdMap {
    param([string] $FilePath)
    if ([string]::IsNullOrWhiteSpace($FilePath)) { return }
    $byEntity = @{}
    foreach ($key in $script:LookupIdMap.Keys) {
        if ($key -match '^([^|]+)\|(.+)$') {
            $ent = $Matches[1]
            $srcId = $Matches[2]
            if (-not $byEntity[$ent]) { $byEntity[$ent] = @{} }
            $byEntity[$ent][$srcId] = [string]$script:LookupIdMap[$key]
        }
    }
    $dir = Split-Path -Parent $FilePath
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $byEntity | ConvertTo-Json -Depth 5 | Set-Content -Path $FilePath -Encoding UTF8 -ErrorAction SilentlyContinue
}

function Import-MigrationIdMap {
    param([string] $FilePath)
    if ([string]::IsNullOrWhiteSpace($FilePath) -or -not (Test-Path $FilePath)) { return @{} }
    try {
        $json = Get-Content -Path $FilePath -Raw -Encoding UTF8
        $byEntity = $json | ConvertFrom-Json
        $out = @{}
        $byEntity.PSObject.Properties | ForEach-Object {
            $ent = $_.Name
            if (-not $out.ContainsKey($ent)) { $out[$ent] = @{} }
            $_.Value.PSObject.Properties | ForEach-Object {
                $sid = $_.Name
                $out[$ent][$sid] = [guid]$_.Value
            }
        }
        return $out
    } catch {
        return @{}
    }
}

# Dla atrybutow bez metadanych (Type=String): proba konwersji string->wlasciwy typ na podstawie nazwy atrybutu
function Normalize-ValueByAttributeName {
    param([string] $AttributeName, $Value)
    if ($null -eq $Value) { return $null }
    if ($Value -isnot [string]) { return $Value }
    $ln = $AttributeName.ToLowerInvariant().Trim()
    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return $Value }

    # Atrybuty typowo Boolean (bit)
    $boolPatterns = @('donot', 'follow', 'participates', 'creditonhold', 'decisionmaker', 'evaluatefit', 'marketingonly', 'isbackoffice', 'confirminterest', 'merged', 'msdyn_gdproptout', 'msdyn_isminor', 'msdyn_disablewebtracking', 'xvo_flag', 'adx_', 'isdefault', 'systemmanaged', 'issastokenset')
    foreach ($p in $boolPatterns) {
        if ($ln -like ($p + '*') -or $ln -eq $p) {
            $t = $s.Trim().ToLowerInvariant()
            if ($t -in 'true','1','yes','t','y') { return $true }
            if ($t -in 'false','0','no','f','n','') { return $false }
            break
        }
    }

    # Atrybuty typowo OptionSet / integer (code, type, precision, mask)
    $optionSetPatterns = @('code', 'type', 'precision', 'mask', 'state', 'status', 'addresstype', 'shippingmethod', 'featuremask', 'currencyprecision', 'userlicensetype', 'prioritycode', 'customertypecode', 'territorycode', 'preferredappointmenttimecode', 'haschildrencode')
    foreach ($p in $optionSetPatterns) {
        if ($ln -like ('*' + $p) -or $ln -eq $p) {
            $i = 0
            if ([int]::TryParse($s.Trim(), [ref]$i)) { return [Microsoft.Xrm.Sdk.OptionSetValue]::new($i) }
            break
        }
    }

    # Atrybuty typowo Decimal/Double (quantity, amount, value, rate)
    $decimalPatterns = @('quantity', 'amount', 'value', 'rate', 'exchangerate', 'total', 'estimated', 'precision')
    foreach ($p in $decimalPatterns) {
        if ($ln -like ('*' + $p + '*') -or $ln -eq $p) {
            $d = [decimal]::Zero
            if ([decimal]::TryParse($s.Trim().Replace(',', [cultureinfo]::CurrentCulture.NumberFormat.NumberDecimalSeparator), [ref]$d)) { return $d }
            break
        }
    }

    return $Value
}

# Zwraca wartosc nadajaca sie do serializacji (bez cykli) – klonuje obiekty SDK zeby zlamac referencje
function Get-SerializableValue {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [string] -or $Value -is [int] -or $Value -is [long] -or $Value -is [bool] -or $Value -is [decimal] -or $Value -is [double] -or $Value -is [DateTime] -or $Value -is [guid]) { return $Value }
    if ($Value -is [Microsoft.Xrm.Sdk.EntityReference]) {
        try { return [Microsoft.Xrm.Sdk.EntityReference]::new($Value.LogicalName, $Value.Id) } catch { return $null }
    }
    if ($Value -is [Microsoft.Xrm.Sdk.OptionSetValue]) {
        try { return [Microsoft.Xrm.Sdk.OptionSetValue]::new($Value.Value) } catch { return $null }
    }
    if ($Value -is [Microsoft.Xrm.Sdk.Money]) {
        try { return [Microsoft.Xrm.Sdk.Money]::new($Value.Value) } catch { return $null }
    }
    if ($Value -is [hashtable] -or $Value -is [System.Collections.IDictionary]) { return $null }
    if ($Value -is [PSObject]) { try { return [string]$Value } catch { return $null } }
    if ($Value.GetType().Name -like 'KeyValuePair*') { try { return Get-SerializableValue -Value $Value.Value } catch { return $null } }
    try { return [string]$Value } catch { return $null }
}

function Get-MigratableAttributes {
    param(
        [hashtable] $EntityMeta,
        [string[]] $SystemFieldsToSkip
    )
    $out = @()
    foreach ($a in $EntityMeta.Attributes) {
        $ln = $a.LogicalName
        if ($ln -in $SystemFieldsToSkip) { continue }
        if ($ln -eq 'overriddencreatedon') { continue }  # ustawiamy z createdon
        $out += $a
    }
    return $out
}

function Convert-RecordToTargetAttributes {
    param(
        [hashtable] $SourceRecord,
        [array] $MigratableAttrs,
        [string[]] $SystemFieldsToSkip,
        [hashtable] $LookupIdMap,
        [string] $EntityLogicalName,
        [object] $SourceConn = $null,
        [object] $TargetConn = $null,
        [hashtable] $Config = $null,
        [string[]] $EntitiesInScope = $null,
        [hashtable] $CommonEntities = $null,
        [hashtable] $TargetMeta = $null,
        [bool] $AutoMigrateMissingLookups = $false,
        [scriptblock] $MigrateSingleRecordIfMissing = $null
    )
    $targetAttrs = @{}
    # Mapy typow atrybutow na celu – uzywane do konwersji (cel moze miec inne typy niz zrodlo)
    $targetTypeMap = @{}
    $targetLookupTargets = @{}
    if ($TargetMeta -and $TargetMeta.ContainsKey($EntityLogicalName) -and $TargetMeta[$EntityLogicalName].Attributes) {
        foreach ($a in $TargetMeta[$EntityLogicalName].Attributes) {
            $targetTypeMap[[string]$a.LogicalName.ToLowerInvariant()] = [string]$a.Type
            $t = [string]$a.Type
            if ($t -eq 'Lookup' -or $t -eq 'Customer' -or $t -eq 'Owner') {
                $targetLookupTargets[[string]$a.LogicalName.ToLowerInvariant()] = if ([string]::IsNullOrWhiteSpace($a.Target)) { if ($t -eq 'Owner') { 'systemuser,team' } else { 'account,contact' } } else { $a.Target }
            }
        }
    }
    # Pobierz oryginalna date utworzenia (createdon nie jest w MigratableAttrs bo jest w SystemFieldsToSkip)
    $createdOn = $null
    if ($SourceRecord.ContainsKey('createdon')) { $createdOn = $SourceRecord['createdon'] }
    if ($null -eq $createdOn -and $SourceRecord.ContainsKey('createdon_property')) {
        $coProp = $SourceRecord['createdon_property']
        try { if ($null -ne $coProp -and $coProp.Count -ge 2) { $createdOn = $coProp[1] } } catch { }
    }

    foreach ($attr in $MigratableAttrs) {
        $ln = $attr.LogicalName
        if ([string]::IsNullOrWhiteSpace($ln)) { continue }
        $ln = $ln.ToLowerInvariant()
        if ($ln -in $SystemFieldsToSkip) { continue }
        if (-not $SourceRecord.ContainsKey($ln)) { continue }

        $typeStr = [string]$attr.Type
        if ($targetTypeMap.ContainsKey($ln)) { $typeStr = $targetTypeMap[$ln] }
        $val = $SourceRecord[$ln]
        if ($null -eq $val -and $SourceRecord.ContainsKey($ln + '_property')) {
            $prop = $SourceRecord[$ln + '_property']
            if ($null -ne $prop) {
                try {
                    if ($prop.GetType().Name -like 'KeyValuePair*') { $val = $prop.Value }
                    elseif ($prop -is [Array] -and $prop.Length -ge 2) { $val = $prop[1] }
                    elseif ($null -ne $prop.Count -and $prop.Count -ge 2) { $val = $prop[1] }
                    elseif ($prop.PSObject.Properties['Value']) { $val = $prop.Value }
                    elseif ($prop -is [System.Collections.IList] -and $prop.Count -ge 2) { $val = $prop[1] }
                    else {
                        $arr = @($prop)
                        if ($arr.Count -ge 2) { $val = $arr[1] }
                    }
                } catch { }
            }
        }
        # SDK moze zwracac wartosc bezposrednio jako KeyValuePair (np. fullname) – rozpakuj
        try {
            if ($null -ne $val -and $val.GetType().Name -like 'KeyValuePair*' -and $null -ne $val.Value) { $val = $val.Value }
        } catch { }
        if ($null -eq $val) { continue }

        if ($ln -eq 'createdon') {
            $createdOn = $val
            continue
        }

        # ownerid przy braku metadanych (Type=String) – ze zrodla moze byc string/GUID; cel wymaga EntityReference
        if ($ln -eq 'ownerid') {
            $sourceId = $null
            $targetEntity = 'systemuser'
            if ($val -is [Microsoft.Xrm.Sdk.EntityReference]) {
                $sourceId = $val.Id
                $targetEntity = $val.LogicalName
            } elseif ($val -is [guid]) {
                $sourceId = $val
            } elseif ($val -is [string]) {
                $g = [guid]::Empty
                if ([guid]::TryParse([string]$val.Trim(), [ref]$g)) { $sourceId = $g }
            }
            if ($null -ne $sourceId) {
                $mapped = Get-MappedLookupId -EntityLogicalName $targetEntity -SourceId $sourceId
                if ($null -eq $mapped -and $SourceConn -and $TargetConn -and $Config) {
                    $mapped = ResolveLookupId -SourceConn $SourceConn -TargetConn $TargetConn -Config $Config -TargetEntity $targetEntity -SourceId $sourceId -EntitiesInScope $EntitiesInScope -CommonEntities $CommonEntities -TargetMeta $TargetMeta -AutoMigrateMissingLookups $AutoMigrateMissingLookups -MigrateSingleRecordIfMissing $MigrateSingleRecordIfMissing
                }
                if ($null -ne $mapped) {
                    $targetAttrs[$ln] = [Microsoft.Xrm.Sdk.EntityReference]::new($targetEntity, $mapped)
                }
            }
            continue
        }

        # Typy obsługiwane (porownanie po stringu – SDK zwraca enum)
        if ($typeStr -eq 'Lookup' -or $typeStr -eq 'Customer' -or $typeStr -eq 'Owner') {
            $ref = $val
            $sourceId = $null
            $targetEntity = $null
            if ($ref -is [Microsoft.Xrm.Sdk.EntityReference]) {
                $targetEntity = $ref.LogicalName
                $sourceId = $ref.Id
            } elseif ($val -is [hashtable] -and $val.id) {
                $targetEntity = if ($val.logicalname) { $val.logicalname } else { $val.LogicalName }
                $sourceId = [guid]$val.id
            } elseif ($val -is [guid]) {
                $sourceId = $val
                $targetEntity = $null
            } elseif ($val -is [string]) {
                $g = [guid]::Empty
                if ([guid]::TryParse([string]$val.Trim(), [ref]$g)) { $sourceId = $g; $targetEntity = $null }
            }
            if ($null -ne $sourceId) {
                $mapped = $null
                if ($null -ne $targetEntity) {
                    $mapped = Get-MappedLookupId -EntityLogicalName $targetEntity -SourceId $sourceId
                    if ($null -eq $mapped -and $SourceConn -and $TargetConn -and $Config) {
                        $mapped = ResolveLookupId -SourceConn $SourceConn -TargetConn $TargetConn -Config $Config -TargetEntity $targetEntity -SourceId $sourceId -EntitiesInScope $EntitiesInScope -CommonEntities $CommonEntities -TargetMeta $TargetMeta -AutoMigrateMissingLookups $AutoMigrateMissingLookups -MigrateSingleRecordIfMissing $MigrateSingleRecordIfMissing
                    }
                    if ($null -ne $mapped) {
                        $targetAttrs[$ln] = [Microsoft.Xrm.Sdk.EntityReference]::new($targetEntity, $mapped)
                    }
                } else {
                    $targets = if ($targetLookupTargets.ContainsKey($ln)) { $targetLookupTargets[$ln] -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } } else { @('account','contact') }
                    foreach ($te in $targets) {
                        $mapped = Get-MappedLookupId -EntityLogicalName $te -SourceId $sourceId
                        if ($null -eq $mapped -and $SourceConn -and $TargetConn -and $Config) {
                            $mapped = ResolveLookupId -SourceConn $SourceConn -TargetConn $TargetConn -Config $Config -TargetEntity $te -SourceId $sourceId -EntitiesInScope $EntitiesInScope -CommonEntities $CommonEntities -TargetMeta $TargetMeta -AutoMigrateMissingLookups $AutoMigrateMissingLookups -MigrateSingleRecordIfMissing $MigrateSingleRecordIfMissing
                        }
                        if ($null -ne $mapped) {
                            $targetAttrs[$ln] = [Microsoft.Xrm.Sdk.EntityReference]::new($te, $mapped)
                            break
                        }
                    }
                }
            }
            continue
        }

        if ($typeStr -eq 'OptionSet' -or $typeStr -eq 'State' -or $typeStr -eq 'Status') {
            if ($val -is [Microsoft.Xrm.Sdk.OptionSetValue]) {
                $targetAttrs[$ln] = $val
            } elseif ($val -is [hashtable] -and $null -ne $val.Value) {
                $targetAttrs[$ln] = [Microsoft.Xrm.Sdk.OptionSetValue]::new([int]$val.Value)
            } elseif ($val -is [int]) {
                $targetAttrs[$ln] = [Microsoft.Xrm.Sdk.OptionSetValue]::new($val)
            } else {
                $intVal = 0
                if ($null -ne $val -and [int]::TryParse([string]$val, [ref]$intVal)) {
                    $targetAttrs[$ln] = [Microsoft.Xrm.Sdk.OptionSetValue]::new($intVal)
                }
            }
            continue
        }

        if ($typeStr -eq 'Money') {
            if ($val -is [Microsoft.Xrm.Sdk.Money]) {
                $targetAttrs[$ln] = $val
            } elseif ($val -is [hashtable] -and $null -ne $val.Value) {
                $targetAttrs[$ln] = [Microsoft.Xrm.Sdk.Money]::new([decimal]$val.Value)
            } elseif ($val -is [decimal] -or $val -is [double]) {
                $targetAttrs[$ln] = [Microsoft.Xrm.Sdk.Money]::new([decimal]$val)
            }
            continue
        }

        # Typy proste
        if ($typeStr -eq 'DateTime') {
            $dtVal = $val
            if ($val -is [string]) {
                $parsed = [DateTime]::MinValue
                if ([DateTime]::TryParse($val, [ref]$parsed)) {
                    $dtVal = $parsed
                } else {
                    $dtVal = [DateTime]::Parse($val)
                }
            }
            $targetAttrs[$ln] = $dtVal
            continue
        }
        if ($typeStr -in @('String','Integer','BigInt','Double','Decimal','Boolean','Uniqueidentifier')) {
            if ($typeStr -eq 'String' -and $val -is [string]) {
                $targetAttrs[$ln] = Normalize-ValueByAttributeName -AttributeName $ln -Value $val
            } elseif ($typeStr -eq 'Boolean' -and $val -is [string]) {
                $t = [string]$val.Trim().ToLowerInvariant()
                $targetAttrs[$ln] = ($t -in 'true','1','yes','t','y')
            } elseif (($typeStr -eq 'Double' -or $typeStr -eq 'Decimal') -and $val -is [string]) {
                $d = [decimal]::Zero
                if ([decimal]::TryParse([string]$val.Trim().Replace(',', [cultureinfo]::CurrentCulture.NumberFormat.NumberDecimalSeparator), [ref]$d)) { $targetAttrs[$ln] = $d } else { $targetAttrs[$ln] = $val }
            } elseif (($typeStr -eq 'Integer' -or $typeStr -eq 'BigInt') -and $val -is [Microsoft.Xrm.Sdk.OptionSetValue]) {
                # Cel ma typ Integer/BigInt (np. currencytype) – API oczekuje int, nie OptionSetValue
                $targetAttrs[$ln] = $val.Value
            } else {
                $targetAttrs[$ln] = $val
            }
            continue
        }

        # EntityName, Memo, itd.
        if ($typeStr -in @('EntityName','Memo')) {
            $targetAttrs[$ln] = $val
            continue
        }

        # Fallback (brak typu w metadanych): okresl typ po wartosci ze zrodla i przypisz to samo w celu
        if ($val -is [hashtable] -or $val -is [System.Collections.IDictionary]) {
            if ($val.id -and ($val.logicalname -or $val.LogicalName)) {
                $targetEntity = if ($val.logicalname) { $val.logicalname } else { $val.LogicalName }
                $sourceId = [guid]$val.id
                $mapped = Get-MappedLookupId -EntityLogicalName $targetEntity -SourceId $sourceId
                if ($null -eq $mapped -and $SourceConn -and $TargetConn -and $Config) {
                    $mapped = ResolveLookupId -SourceConn $SourceConn -TargetConn $TargetConn -Config $Config -TargetEntity $targetEntity -SourceId $sourceId -EntitiesInScope $EntitiesInScope -CommonEntities $CommonEntities -TargetMeta $TargetMeta -AutoMigrateMissingLookups $AutoMigrateMissingLookups -MigrateSingleRecordIfMissing $MigrateSingleRecordIfMissing
                }
                if ($null -ne $mapped) {
                    $targetAttrs[$ln] = [Microsoft.Xrm.Sdk.EntityReference]::new($targetEntity, $mapped)
                }
            }
            continue
        }
        if ($val -is [Microsoft.Xrm.Sdk.OptionSetValue]) {
            # Jesli cel oczekuje Integer/BigInt (np. transactioncurrency.currencytype), wysylamy int
            if ($targetTypeMap.ContainsKey($ln) -and $targetTypeMap[$ln] -in 'Integer','BigInt') {
                $targetAttrs[$ln] = $val.Value
            } else {
                $targetAttrs[$ln] = [Microsoft.Xrm.Sdk.OptionSetValue]::new($val.Value)
            }
            continue
        }
        if ($val -is [Microsoft.Xrm.Sdk.EntityReference]) {
            $targetEntity = $val.LogicalName
            $sourceId = $val.Id
            $mapped = Get-MappedLookupId -EntityLogicalName $targetEntity -SourceId $sourceId
            if ($null -eq $mapped -and $SourceConn -and $TargetConn -and $Config) {
                $mapped = ResolveLookupId -SourceConn $SourceConn -TargetConn $TargetConn -Config $Config -TargetEntity $targetEntity -SourceId $sourceId -EntitiesInScope $EntitiesInScope -CommonEntities $CommonEntities -TargetMeta $TargetMeta -AutoMigrateMissingLookups $AutoMigrateMissingLookups -MigrateSingleRecordIfMissing $MigrateSingleRecordIfMissing
            }
            if ($null -ne $mapped) {
                $targetAttrs[$ln] = [Microsoft.Xrm.Sdk.EntityReference]::new($targetEntity, $mapped)
            }
            continue
        }
        if ($val -is [Microsoft.Xrm.Sdk.Money]) {
            $targetAttrs[$ln] = [Microsoft.Xrm.Sdk.Money]::new($val.Value)
            continue
        }
        if ($val -is [DateTime]) {
            $targetAttrs[$ln] = $val
            continue
        }
        if ($val -is [int] -or $val -is [long] -or $val -is [bool] -or $val -is [decimal] -or $val -is [double] -or $val -is [guid]) {
            $targetAttrs[$ln] = $val
            continue
        }
        if ($val -is [string]) {
            $s = [string]$val.Trim()
            if ($targetTypeMap.ContainsKey($ln)) {
                $t = $targetTypeMap[$ln]
                if ($t -eq 'Boolean') {
                    $targetAttrs[$ln] = ($s -in 'true','1','yes','t','y')
                    continue
                }
                if ($t -in 'OptionSet','State','Status') {
                    $intVal = 0
                    if ([int]::TryParse($s, [ref]$intVal)) {
                        $targetAttrs[$ln] = [Microsoft.Xrm.Sdk.OptionSetValue]::new($intVal)
                        continue
                    }
                }
                if ($t -eq 'Money' -or $t -eq 'Double' -or $t -eq 'Decimal') {
                    $d = [decimal]::Zero
                    if ([decimal]::TryParse($s.Replace(',', [cultureinfo]::CurrentCulture.NumberFormat.NumberDecimalSeparator), [ref]$d)) {
                        if ($t -eq 'Money') { $targetAttrs[$ln] = [Microsoft.Xrm.Sdk.Money]::new($d) } else { $targetAttrs[$ln] = $d }
                        continue
                    }
                }
                if ($t -eq 'Lookup' -or $t -eq 'Customer' -or $t -eq 'Owner') {
                    $g = [guid]::Empty
                    if ([guid]::TryParse($s, [ref]$g)) {
                        $targets = if ($targetLookupTargets.ContainsKey($ln)) { $targetLookupTargets[$ln] -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } } else { @('account','contact') }
                        foreach ($te in $targets) {
                            $mapped = Get-MappedLookupId -EntityLogicalName $te -SourceId $g
                            if ($null -eq $mapped -and $SourceConn -and $TargetConn -and $Config) {
                                $mapped = ResolveLookupId -SourceConn $SourceConn -TargetConn $TargetConn -Config $Config -TargetEntity $te -SourceId $g -EntitiesInScope $EntitiesInScope -CommonEntities $CommonEntities -TargetMeta $TargetMeta -AutoMigrateMissingLookups $AutoMigrateMissingLookups -MigrateSingleRecordIfMissing $MigrateSingleRecordIfMissing
                            }
                            if ($null -ne $mapped) {
                                $targetAttrs[$ln] = [Microsoft.Xrm.Sdk.EntityReference]::new($te, $mapped)
                                break
                            }
                        }
                        if ($targetAttrs.ContainsKey($ln)) { continue }
                    }
                }
            }
            $targetAttrs[$ln] = Normalize-ValueByAttributeName -AttributeName $ln -Value $val
            continue
        }
        if ($val -is [PSObject]) {
            try {
                $targetAttrs[$ln] = [string]$val
            } catch { }
            continue
        }
        $intVal = 0
        if ($null -ne $val -and [int]::TryParse([string]$val, [ref]$intVal)) {
            $targetAttrs[$ln] = [Microsoft.Xrm.Sdk.OptionSetValue]::new($intVal)
            continue
        }
        try { $targetAttrs[$ln] = [string]$val } catch { }
    }

    if ($null -ne $createdOn) {
        # API wymaga DateTime; FetchXML moze zwracac createdon jako string
        $dt = $createdOn
        if ($createdOn -is [string]) {
            $parsed = [DateTime]::MinValue
            if ([DateTime]::TryParse($createdOn, [ref]$parsed)) {
                $dt = $parsed
            } else {
                $dt = [DateTime]::Parse($createdOn)
            }
        }
        $targetAttrs['overriddencreatedon'] = $dt
    }
    return $targetAttrs, $createdOn
}

function Get-SourceRecordsPaged {
    param(
        $Conn,
        [string] $EntityLogicalName,
        [string[]] $Fields,
        [int] $PageSize = 5000,
        [scriptblock] $Logger
    )
    $allRecords = [System.Collections.ArrayList]::new()
    $fieldsClean = @($Fields | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $attrList = ($fieldsClean | ForEach-Object { "<attribute name='$_' />" }) -join "`n"
    $page = 1
    $cookie = $null
    do {
        if ($page -eq 1) {
            $fetch = @"
<fetch version="1.0" mapping="logical" returntotalrecordcount="true" count="$PageSize" page="1">
  <entity name="$EntityLogicalName">
    $attrList
  </entity>
</fetch>
"@
        } else {
            $encCookie = [System.Web.HttpUtility]::UrlEncode($cookie)
            $fetch = @"
<fetch version="1.0" mapping="logical" returntotalrecordcount="true" count="$PageSize" page="$page" paging-cookie="$encCookie">
  <entity name="$EntityLogicalName">
    $attrList
  </entity>
</fetch>
"@
        }
        $result = Get-CrmRecordsByFetch -conn $Conn -Fetch $fetch -ErrorAction Stop
        $records = $result.CrmRecords
        if ($records) {
            foreach ($r in $records) { [void]$allRecords.Add($r) }
        }
        $cookie = $result.PagingCookie
        $more = $result.NextPage -eq $true
        $page++
        if ($Logger -and $more) { & $Logger ('  Strona ' + $page + ' ...') }
    } while ($more)
    # Zwroc jako tablice, zeby foreach w Copy-EntityRecords zawsze mial pewna enumeracje
    return @($allRecords)
}

function Invoke-Retry {
    param(
        [scriptblock] $Action,
        [int] $MaxRetries = 3,
        [int] $DelaySeconds = 5,
        [scriptblock] $Logger
    )
    $attempt = 0
    while ($true) {
        try {
            return & $Action
        } catch {
            $attempt++
            if ($Logger) { & $Logger ('  Proba ' + $attempt + '/' + $MaxRetries + ' nie powiodla sie: ' + $_) }
            if ($attempt -ge $MaxRetries) { throw }
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

function Get-SourceRecordKeyValue {
    param($SourceConn, [string] $EntityLogicalName, [guid] $SourceId, [string] $KeyAttribute)
    if ([string]::IsNullOrWhiteSpace($KeyAttribute)) { return $null }
    $pkAttr = Get-PrimaryKeyAttributeName -EntityLogicalName $EntityLogicalName
    $attrsToFetch = @($KeyAttribute)
    if ($EntityLogicalName -eq 'systemuser' -and $KeyAttribute -eq 'fullname') {
        $attrsToFetch = @('fullname', 'firstname', 'lastname')
    }
    $attrList = ($attrsToFetch | ForEach-Object { "<attribute name='$_' />" }) -join "`n"
    $fetch = @"
<fetch version="1.0" mapping="logical" top="1" no-lock="true">
  <entity name="$EntityLogicalName">
    $attrList
    <filter><condition attribute="$pkAttr" operator="eq" value="$SourceId" /></filter>
  </entity>
</fetch>
"@
    try {
        $result = Get-CrmRecordsByFetch -conn $SourceConn -Fetch $fetch -ErrorAction Stop
        $recs = $result.CrmRecords
        if ($recs -and $recs.Count -gt 0) {
            $r = $recs[0]
            $getVal = {
                param($name)
                if ($r.PSObject.Properties[$name]) { return $r.$name }
                $prefixed = 'returnproperty_' + $name.ToLowerInvariant()
                if ($r.PSObject.Properties[$prefixed]) { return $r.$prefixed }
                return $null
            }
            $val = & $getVal $KeyAttribute
            if ($null -ne $val -and -not [string]::IsNullOrWhiteSpace([string]$val)) { return $val }
            if ($EntityLogicalName -eq 'systemuser' -and $KeyAttribute -eq 'fullname') {
                $fn = & $getVal 'firstname'
                $ln = & $getVal 'lastname'
                $combined = (@($fn, $ln) | Where-Object { $_ }) -join ' '
                if (-not [string]::IsNullOrWhiteSpace($combined)) { return $combined.Trim() }
            }
        }
    } catch { }
    return $null
}

function Get-CurrentUserId {
    param($Conn)
    if (-not $Conn) { return $null }
    try {
        $req = New-Object Microsoft.Crm.Sdk.Messages.WhoAmIRequest
        $resp = $Conn.Execute($req)
        if ($resp -and $resp.UserId) { return $resp.UserId }
    } catch { }
    return $null
}

function ResolveLookupId {
    param(
        $SourceConn,
        $TargetConn,
        [hashtable] $Config,
        [string] $TargetEntity,
        [guid] $SourceId,
        [string[]] $EntitiesInScope = $null,
        [hashtable] $CommonEntities = $null,
        [hashtable] $TargetMeta = $null,
        [bool] $AutoMigrateMissingLookups = $false,
        [scriptblock] $MigrateSingleRecordIfMissing = $null
    )
    $mapped = Get-MappedLookupId -EntityLogicalName $TargetEntity -SourceId $SourceId
    if ($null -ne $mapped) { return $mapped }
    $defaults = if ($Config.EntityDefaultTargetLookup) { $Config.EntityDefaultTargetLookup } else { @{} }
    if ($defaults -and $defaults.ContainsKey($TargetEntity) -and -not [string]::IsNullOrWhiteSpace($defaults[$TargetEntity])) {
        return [guid]$defaults[$TargetEntity]
    }
    $resolveByName = if ($Config.EntityLookupResolveByName) { @($Config.EntityLookupResolveByName) } else { @() }
    if ($TargetEntity -in $resolveByName) {
        $matchKeys = if ($Config.EntityMatchKey) { $Config.EntityMatchKey } else { @{} }
        $keyAttr = if ($matchKeys.ContainsKey($TargetEntity)) { $matchKeys[$TargetEntity] } else { 'name' }
        $keyVal = Get-SourceRecordKeyValue -SourceConn $SourceConn -EntityLogicalName $TargetEntity -SourceId $SourceId -KeyAttribute $keyAttr
        if ($null -ne $keyVal -and -not [string]::IsNullOrWhiteSpace([string]$keyVal)) {
            $targetId = Get-TargetRecordIdByKey -Conn $TargetConn -EntityLogicalName $TargetEntity -KeyAttribute $keyAttr -KeyValue $keyVal
            if ($null -ne $targetId) {
                Initialize-LookupMap -EntityLogicalName $TargetEntity -SourceId $SourceId -TargetId $targetId
                return $targetId
            }
        }
    }
    if ($TargetEntity -eq 'systemuser') {
        $currentUserId = Get-CurrentUserId -Conn $TargetConn
        if ($null -ne $currentUserId) { return $currentUserId }
    }
    # Auto-domigrowanie brakujacego rekordu ze zrodla (tylko gdy OnlyEntitiesWithRecordsAndDependencies)
    if ($AutoMigrateMissingLookups -and $MigrateSingleRecordIfMissing -and $EntitiesInScope -and $TargetEntity -in $EntitiesInScope) {
        $stackKey = "${TargetEntity}|$SourceId"
        if ($script:AutoMigratePullStack -notcontains $stackKey) {
            try {
                $pulled = & $MigrateSingleRecordIfMissing $TargetEntity $SourceId $MigrateSingleRecordIfMissing
                if ($null -ne $pulled -and $pulled -is [guid]) { return $pulled }
            } catch { }
        }
    }
    return $null
}

function Migrate-SingleRecordFromSource {
    <#
    .SYNOPSIS
        Pobiera jeden rekord ze zrodla, konwertuje i tworzy w celu (dla spinania lookupow przy OnlyEntitiesWithRecordsAndDependencies).
    #>
    param(
        $SourceConn,
        $TargetConn,
        [string] $EntityLogicalName,
        [guid] $SourceId,
        [hashtable] $Config,
        [hashtable] $CommonEntities,
        [hashtable] $TargetMeta,
        [string[]] $EntitiesInScope,
        [scriptblock] $LogInfo,
        [scriptblock] $MigrateSingleRecordIfMissing,
        [bool] $ForceRecreate = $false
    )
    if (-not $EntitiesInScope -or $EntityLogicalName -notin $EntitiesInScope) { return $null }
    $stackKey = "${EntityLogicalName}|$SourceId"
    if ($script:AutoMigratePullStack -contains $stackKey) { return $null }
    if (-not $ForceRecreate) {
        $mapped = Get-MappedLookupId -EntityLogicalName $EntityLogicalName -SourceId $SourceId
        if ($null -ne $mapped) { return $mapped }
    }

    [void]$script:AutoMigratePullStack.Add($stackKey)
    try {
        $entityMeta = if ($CommonEntities -and $CommonEntities.ContainsKey($EntityLogicalName)) { $CommonEntities[$EntityLogicalName] } else { $null }
        if (-not $entityMeta) { return $null }
        $skipFields = if ($Config -and $Config.SystemFieldsToSkip) { @($Config.SystemFieldsToSkip) } else { @() }
        $getCrmRec = Get-Command -Name Get-CrmRecord -ErrorAction SilentlyContinue
        if (-not $getCrmRec) { return $null }
        $fullRec = & $getCrmRec -conn $SourceConn -EntityLogicalName $EntityLogicalName -Id $SourceId -Fields '*' -ErrorAction Stop
        if (-not $fullRec) { return $null }

        $recHash = @{}
        if ($fullRec.PSObject.Properties['original'] -and $null -ne $fullRec.original -and ($fullRec.original -is [hashtable] -or $fullRec.original -is [System.Collections.IDictionary])) {
            foreach ($k in $fullRec.original.Keys) {
                $strKey = [string]$k
                $ln = if ($strKey -match '^(.+)_Property$') { $Matches[1].ToLowerInvariant() } else { $strKey.ToLowerInvariant() }
                if (-not [string]::IsNullOrWhiteSpace($ln) -and $ln -notlike '*_property') { $recHash[$ln] = $fullRec.original[$k] }
            }
        } elseif ($fullRec.Attributes -and ($fullRec -is [Microsoft.Xrm.Sdk.Entity] -or $null -ne $fullRec.Attributes)) {
            foreach ($a in $fullRec.Attributes) {
                $keyRaw = if ($null -ne $a.Key) { $a.Key } elseif ($a.PSObject.Properties['Key']) { $a.Key } else { $null }
                $val = if ($null -ne $a.Value) { $a.Value } elseif ($a.PSObject.Properties['Value']) { $a.Value } else { $null }
                $ln = if ($keyRaw) { ([string]$keyRaw).ToLowerInvariant() } else { $null }
                if ($ln) { $recHash[$ln] = $val }
            }
            if ($fullRec.Id -ne [guid]::Empty) { $recHash['id'] = $fullRec.Id }
        } else {
            $prefix = 'returnproperty_'
            $fullRec.PSObject.Properties | ForEach-Object {
                $name = $_.Name
                $key = if ($name.Length -gt $prefix.Length -and $name.Substring(0, $prefix.Length).ToLowerInvariant() -eq $prefix) {
                    $name.Substring($prefix.Length).ToLowerInvariant()
                } else { $name.ToLowerInvariant() }
                if ($key -and $key -ne 'attributes') { $recHash[$key] = $_.Value }
            }
            if (-not $recHash.ContainsKey('id') -and $fullRec.PSObject.Properties['id']) { $recHash['id'] = $fullRec.id }
            if (-not $recHash.ContainsKey('id') -and $fullRec.PSObject.Properties['returnProperty_Id']) { $recHash['id'] = $fullRec.returnProperty_Id }
        }
        if (-not $recHash.ContainsKey('id')) { $recHash['id'] = $SourceId }

        # Jesli w celu jest juz rekord o tym samym kluczu (np. name, fullname) – uzyj go zamiast tworzyc duplikat
        $matchKeys = if ($Config -and $Config.EntityMatchKey) { $Config.EntityMatchKey } else { @{} }
        $keyAttr = if ($matchKeys.ContainsKey($EntityLogicalName)) { $matchKeys[$EntityLogicalName] } else { 'name' }
        if ($EntityLogicalName -eq 'systemuser' -and ($keyAttr -eq 'name' -or -not $matchKeys.ContainsKey($EntityLogicalName))) { $keyAttr = 'fullname' }
        if ($keyAttr -and $recHash.ContainsKey($keyAttr)) {
            $keyVal = $recHash[$keyAttr]
            if ($null -ne $keyVal) {
                if ($keyVal -is [Microsoft.Xrm.Sdk.OptionSetValue]) { $keyVal = $keyVal.Value }
                if ($keyVal -is [Microsoft.Xrm.Sdk.Money]) { $keyVal = $keyVal.Value }
                $keyValStr = [string]$keyVal
                if ($EntityLogicalName -eq 'systemuser' -and $keyAttr -eq 'fullname' -and [string]::IsNullOrWhiteSpace($keyValStr.Trim()) -and ($recHash.ContainsKey('firstname') -or $recHash.ContainsKey('lastname'))) {
                    $keyValStr = ([string]$recHash['firstname']).Trim() + ' ' + ([string]$recHash['lastname']).Trim()
                }
                if (-not [string]::IsNullOrWhiteSpace($keyValStr)) {
                    $existingTargetId = Get-TargetRecordIdByKey -Conn $TargetConn -EntityLogicalName $EntityLogicalName -KeyAttribute $keyAttr -KeyValue $keyValStr
                    if ($null -ne $existingTargetId) {
                        Initialize-LookupMap -EntityLogicalName $EntityLogicalName -SourceId $SourceId -TargetId $existingTargetId
                        if ($LogInfo) { & $LogInfo ("  [Auto-pull] Znaleziono istniejacy rekord w celu (pasuje po $keyAttr): $EntityLogicalName - uzyto zamiast duplikatu.") }
                        return $existingTargetId
                    }
                }
            }
        }

        $migratableAttrs = Get-MigratableAttributes -EntityMeta $entityMeta -SystemFieldsToSkip $skipFields
        $targetAttrSet = $null
        if ($TargetMeta -and $TargetMeta.ContainsKey($EntityLogicalName) -and $TargetMeta[$EntityLogicalName].Attributes) {
            $targetAttrSet = @{}
            foreach ($a in $TargetMeta[$EntityLogicalName].Attributes) { $targetAttrSet[[string]$a.LogicalName.ToLowerInvariant()] = $true }
        }
        if ($migratableAttrs.Count -eq 0 -and $targetAttrSet -and $targetAttrSet.Count -gt 0) {
            $exclude = @('id', 'createdon', 'overriddencreatedon') + @($skipFields)
            $migratableAttrs = @( $recHash.Keys | Where-Object {
                $k = $_; $k -ne 'id' -and $k -notlike '*_property' -and $k -notin $exclude -and $targetAttrSet.ContainsKey($k)
            } | ForEach-Object { [PSCustomObject]@{ LogicalName = $_; Type = 'String' } } )
        }
        if ($migratableAttrs.Count -eq 0) { return $null }

        $targetAttrs, $createdOn = Convert-RecordToTargetAttributes -SourceRecord $recHash -MigratableAttrs $migratableAttrs -SystemFieldsToSkip $skipFields `
            -LookupIdMap $script:LookupIdMap -EntityLogicalName $EntityLogicalName -SourceConn $SourceConn -TargetConn $TargetConn -Config $Config `
            -EntitiesInScope $EntitiesInScope -CommonEntities $CommonEntities -TargetMeta $TargetMeta -AutoMigrateMissingLookups $true -MigrateSingleRecordIfMissing $MigrateSingleRecordIfMissing
        if (-not $targetAttrs -or $targetAttrs.Count -eq 0) { return $null }

        $fieldsCreate = @{}
        $keysToSend = @($targetAttrs.Keys | Select-Object -Unique)
        if ($targetAttrSet -and $targetAttrSet.Count -gt 0) { $keysToSend = @($keysToSend | Where-Object { $targetAttrSet.ContainsKey($_) }) }
        foreach ($k in $keysToSend) {
            $safe = Get-SerializableValue -Value $targetAttrs[$k]
            if ($null -ne $safe) { $fieldsCreate[$k] = $safe }
        }
        if ($fieldsCreate.Count -eq 0) { return $null }

        $newId = & $script:NewCrmRecordCmd -conn $TargetConn -EntityLogicalName $EntityLogicalName -Fields $fieldsCreate -ErrorAction Stop
        if ($null -ne $newId -and $recHash['id']) {
            Initialize-LookupMap -EntityLogicalName $EntityLogicalName -SourceId $recHash['id'] -TargetId $newId
        }
        if ($LogInfo) { & $LogInfo ("  [Auto-pull] Dociagnieto rekord ze zrodla: $EntityLogicalName ($SourceId) dla relacji.") }
        return $newId
    } catch {
        if ($LogInfo) { & $LogInfo ("  [Auto-pull] Nie udalo sie dociagnac rekordu $EntityLogicalName ($SourceId): " + $_.Exception.Message) }
        return $null
    } finally {
        $script:AutoMigratePullStack.Remove($stackKey) | Out-Null
    }
}

function Get-TargetRecordIdByKey {
    param($Conn, [string] $EntityLogicalName, [string] $KeyAttribute, $KeyValue)
    if ([string]::IsNullOrWhiteSpace($KeyAttribute)) { return $null }
    if ($null -eq $KeyValue -or [string]::IsNullOrWhiteSpace([string]$KeyValue)) { return $null }
    $pkAttr = Get-PrimaryKeyAttributeName -EntityLogicalName $EntityLogicalName
    $valStr = [string]$KeyValue
    $valStrEsc = $valStr -replace "&", "&amp;" -replace "<", "&lt;" -replace ">", "&gt;" -replace '"', "&quot;"
    $fetch = @"
<fetch version="1.0" mapping="logical" top="1" no-lock="true">
  <entity name="$EntityLogicalName">
    <attribute name="$pkAttr" />
    <filter><condition attribute="$KeyAttribute" operator="eq" value="$valStrEsc" /></filter>
  </entity>
</fetch>
"@
    try {
        $result = Get-CrmRecordsByFetch -conn $Conn -Fetch $fetch -ErrorAction Stop
        $recs = $result.CrmRecords
        if ($recs -and $recs.Count -gt 0) {
            $r = $recs[0]
            $id = $null
            if ($r.PSObject.Properties[$pkAttr]) { $id = $r.$pkAttr }
            if (-not $id -and $r.PSObject.Properties['returnProperty_Id']) { $id = $r.returnProperty_Id }
            if (-not $id -and $r.PSObject.Properties['id']) { $id = $r.id }
            if ($id) { return if ($id -is [guid]) { $id } else { [guid]$id } }
        }
    } catch { }
    if ($EntityLogicalName -eq 'systemuser' -and $KeyAttribute -eq 'fullname' -and $valStr -match '^\s*(.+?)\s+(.+)\s*$') {
        $first = $Matches[1].Trim()
        $last = $Matches[2].Trim()
        $firstEsc = $first -replace "&", "&amp;" -replace "<", "&lt;" -replace ">", "&gt;" -replace '"', "&quot;"
        $lastEsc = $last -replace "&", "&amp;" -replace "<", "&lt;" -replace ">", "&gt;" -replace '"', "&quot;"
        $fetch2 = @"
<fetch version="1.0" mapping="logical" top="1" no-lock="true">
  <entity name="systemuser">
    <attribute name="systemuserid" />
    <filter type="and">
      <condition attribute="firstname" operator="eq" value="$firstEsc" />
      <condition attribute="lastname" operator="eq" value="$lastEsc" />
    </filter>
  </entity>
</fetch>
"@
        try {
            $result2 = Get-CrmRecordsByFetch -conn $Conn -Fetch $fetch2 -ErrorAction Stop
            $recs2 = $result2.CrmRecords
            if ($recs2 -and $recs2.Count -gt 0) {
                $r2 = $recs2[0]
                if ($r2.PSObject.Properties['systemuserid']) { return [guid]$r2.systemuserid }
                if ($r2.PSObject.Properties['returnProperty_Id']) { return [guid]$r2.returnProperty_Id }
                if ($r2.PSObject.Properties['id']) { return [guid]$r2.id }
            }
        } catch { }
    }
    return $null
}

function Copy-EntityRecords {
    <#
    .SYNOPSIS
        Kopiuje rekordy jednej encji ze źródła do celu z paging i retry.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $SourceConn,
        [Parameter(Mandatory = $true)]
        $TargetConn,
        [Parameter(Mandatory = $true)]
        [string] $EntityLogicalName,
        [Parameter(Mandatory = $true)]
        [hashtable] $EntityMeta,
        [Parameter(Mandatory = $true)]
        [hashtable] $Config,
        [Parameter(Mandatory = $false)]
        [string[]] $TargetAttributeNames = @(),
        [Parameter(Mandatory = $false)]
        [scriptblock] $LogInfo,
        [Parameter(Mandatory = $false)]
        [string[]] $EntitiesInScope = @(),
        [Parameter(Mandatory = $false)]
        [hashtable] $CommonEntities = $null,
        [Parameter(Mandatory = $false)]
        [hashtable] $TargetMeta = $null,
        [Parameter(Mandatory = $false)]
        [bool] $AutoMigrateMissingLookups = $false
    )
    if ($AutoMigrateMissingLookups) {
        $script:AutoMigratePullStack.Clear()
    }
    $migrateSingleRecordIfMissing = $null
    if ($AutoMigrateMissingLookups -and $EntitiesInScope -and $EntitiesInScope.Count -gt 0 -and $CommonEntities) {
        $migrateSingleRecordIfMissing = {
            param($ent, $srcId, $self, $forceRecreate = $false)
            Migrate-SingleRecordFromSource -SourceConn $SourceConn -TargetConn $TargetConn -EntityLogicalName $ent -SourceId $srcId -Config $Config -CommonEntities $CommonEntities -TargetMeta $TargetMeta -EntitiesInScope $EntitiesInScope -LogInfo $LogInfo -MigrateSingleRecordIfMissing $self -ForceRecreate $forceRecreate
        }
    }
    $targetAttrSet = $null
    if ($TargetAttributeNames -and $TargetAttributeNames.Count -gt 0) {
        $targetAttrSet = @{}
        foreach ($a in $TargetAttributeNames) { $targetAttrSet[[string]$a.ToLowerInvariant()] = $true }
    }
    if (($null -eq $targetAttrSet -or $targetAttrSet.Count -eq 0) -and $TargetMeta -and $TargetMeta.ContainsKey($EntityLogicalName) -and $TargetMeta[$EntityLogicalName].Attributes) {
        $targetAttrSet = @{}
        foreach ($a in $TargetMeta[$EntityLogicalName].Attributes) {
            $targetAttrSet[[string]$a.LogicalName.ToLowerInvariant()] = $true
        }
    }
    $pageSize = $Config.PageSize
    $batchSize = [Math]::Min($Config.BatchSize, 100)
    $maxRetry = $Config.MaxRetryCount
    $retryDelay = $Config.RetryDelaySeconds
    $skipFields = $Config.SystemFieldsToSkip
    $mode = if ($Config.MigrationMode) { $Config.MigrationMode } else { 'Upsert' }
    $matchBy = if ($Config.MatchBy) { $Config.MatchBy } else { 'Id' }
    $matchKeys = if ($Config.EntityMatchKey) { $Config.EntityMatchKey } else { @{} }
    $matchKey = $matchKeys[$EntityLogicalName]
    if ($matchBy -eq 'Custom' -and $Config.CustomMatchAttribute) {
        $matchKey = $Config.CustomMatchAttribute
    }
    $attrNames = @($EntityMeta.Attributes | ForEach-Object { $_.LogicalName })
    if ($matchKey -and $matchKey -notin $attrNames) {
        $matchKey = $null
    }
    $idMapPath = $Config.IdMapPath
    if (-not $idMapPath -and $Config.LogFolder) {
        $idMapPath = Join-Path $Config.LogFolder 'IdMap_latest.json'
    }
    $loadedIdMap = @{}
    if ($matchBy -eq 'Id' -or $matchBy -eq 'IdThenName') {
        $loadedIdMap = Import-MigrationIdMap -FilePath $idMapPath
    }

    $pkAttr = Get-PrimaryKeyAttributeName -EntityLogicalName $EntityLogicalName
    $migratableAttrs = Get-MigratableAttributes -EntityMeta $EntityMeta -SystemFieldsToSkip $skipFields
    # Tylko atrybuty istniejace w schemacie celu – zrodlo moze miec wiecej pol (msdyn_*, msgdpr_* itd.)
    if ($TargetMeta -and $TargetMeta.ContainsKey($EntityLogicalName) -and $TargetMeta[$EntityLogicalName].Attributes) {
        $targetAttrNames = @{}
        foreach ($a in $TargetMeta[$EntityLogicalName].Attributes) { $targetAttrNames[[string]$a.LogicalName.ToLowerInvariant()] = $true }
        $migratableAttrs = @($migratableAttrs | Where-Object { $targetAttrNames.ContainsKey($_.LogicalName.ToLowerInvariant()) })
    }
    $baseFields = @($migratableAttrs | ForEach-Object { $_.LogicalName } | Select-Object -Unique)
    $hasCreatedOn = @($EntityMeta.Attributes | Where-Object { $_.LogicalName -eq 'createdon' }).Count -gt 0
    $fieldNames = if ($hasCreatedOn) { @('createdon') + $baseFields } else { @($baseFields) }
    if ($fieldNames -notcontains $pkAttr) { $fieldNames = @($pkAttr) + @($fieldNames | Where-Object { $_ -ne $pkAttr }) }
    if (($mode -eq 'Update' -or $mode -eq 'Upsert' -or $mode -eq 'Create') -and $matchKey -and $fieldNames -notcontains $matchKey) {
        $fieldNames = @($matchKey) + $fieldNames
    }

    $records = Get-SourceRecordsPaged -Conn $SourceConn -EntityLogicalName $EntityLogicalName -Fields $fieldNames -PageSize $pageSize -Logger $LogInfo
    $total = [int]@($records).Count
    if ($total -eq 0) {
        if ($LogInfo) { & $LogInfo '  Brak rekordow do migracji.' }
        return @{ Created = 0; Updated = 0; Skipped = 0; Failed = 0; Total = 0 }
    }
    # W Dataverse nie mozna tworzyc rekordow activitypointer (Create nie jest obsługiwane); aktywności tworzy się przez email, task, appointment itd.
    if ($EntityLogicalName -eq 'activitypointer') {
        if ($LogInfo) { & $LogInfo ('  Pomijam migracje activitypointer - API nie obsluguje Create dla tej encji (aktywnosci migruj przez email/task/appointment itd.). Rekordow w zrodle: ' + $total) }
        return @{ Created = 0; Updated = 0; Skipped = $total; Failed = 0; Total = $total }
    }
    if ($LogInfo) { & $LogInfo ('  Rekordow do migracji: ' + $total + ' (tryb: ' + $mode + ')') }

    # Gdy brak atrybutow z metadanych – odkryj pola z pelnego pierwszego rekordu ze zrodla (Get-CrmRecord -Fields '*') i pobierz wszystkie rekordy z tymi polami
    if ($migratableAttrs.Count -eq 0 -and $total -gt 0) {
        $firstRec = $records[0]
        $firstId = $null
        if ($firstRec.PSObject.Properties[$pkAttr]) { $firstId = $firstRec.$pkAttr }
        if ($null -eq $firstId -and $firstRec.PSObject.Properties['returnProperty_' + $pkAttr]) { $firstId = $firstRec.('returnProperty_' + $pkAttr) }
        if ($null -eq $firstId -and $firstRec.PSObject.Properties['Id']) { $firstId = $firstRec.Id }
        if ($null -ne $firstId) {
            try {
                $getCrmRec = Get-Command -Name Get-CrmRecord -ErrorAction SilentlyContinue
                if ($getCrmRec) {
                    $fullRec = & $getCrmRec -conn $SourceConn -EntityLogicalName $EntityLogicalName -Id $firstId -Fields '*' -ErrorAction Stop
                    $discoveredKeys = [System.Collections.ArrayList]::new()
                    $exclude = @('id', 'createdon', 'overriddencreatedon') + @($skipFields)
                    $sourceKeys = $null
                    if ($fullRec.PSObject.Properties['original'] -and $null -ne $fullRec.original -and ($fullRec.original -is [hashtable] -or $fullRec.original -is [System.Collections.IDictionary])) {
                        $sourceKeys = @($fullRec.original.Keys | ForEach-Object {
                            $k = [string]$_
                            if ($k -match '^(.+)_Property$') { $Matches[1].ToLowerInvariant() } else { $k.ToLowerInvariant() }
                        } | Select-Object -Unique)
                    } elseif ($fullRec.Attributes -and ($fullRec -is [Microsoft.Xrm.Sdk.Entity] -or $null -ne $fullRec.Attributes)) {
                        $sourceKeys = @($fullRec.Attributes | ForEach-Object {
                            $keyRaw = if ($null -ne $_.Key) { $_.Key } elseif ($_.PSObject.Properties['Key']) { $_.Key } else { $null }
                            if ($keyRaw) { [string]$keyRaw.ToLowerInvariant() }
                        } | Where-Object { $_ } | Select-Object -Unique)
                    }
                    if ($sourceKeys) {
                        foreach ($key in $sourceKeys) {
                            if ([string]::IsNullOrWhiteSpace($key) -or $key -in $exclude -or $key -like '*_property') { continue }
                            if ($null -ne $targetAttrSet -and $targetAttrSet.Count -gt 0 -and -not $targetAttrSet.ContainsKey($key)) { continue }
                            if ($discoveredKeys -notcontains $key) { [void]$discoveredKeys.Add($key) }
                        }
                    } else {
                        $metaExclude = @('original', 'logicalname', 'entityreference')
                        $prefix = 'returnproperty_'
                        foreach ($p in $fullRec.PSObject.Properties) {
                            $name = $p.Name
                            if ([string]::IsNullOrWhiteSpace($name)) { continue }
                            $key = if ($name.Length -gt $prefix.Length -and $name.Substring(0, $prefix.Length).ToLowerInvariant() -eq $prefix) {
                                $name.Substring($prefix.Length).ToLowerInvariant()
                            } else {
                                $name.ToLowerInvariant()
                            }
                            if ($key -in $exclude -or $key -in $metaExclude -or $key -like '*_property') { continue }
                            if ($null -ne $targetAttrSet -and $targetAttrSet.Count -gt 0 -and -not $targetAttrSet.ContainsKey($key)) { continue }
                            if ($discoveredKeys -notcontains $key) { [void]$discoveredKeys.Add($key) }
                        }
                    }
                    if ($discoveredKeys.Count -gt 0) {
                        $migratableAttrs = @($discoveredKeys | ForEach-Object { @{ LogicalName = $_; Type = 'String' } })
                        $baseFields = @($discoveredKeys)
                        $fieldNames = @($pkAttr) + @('createdon') + @($baseFields | Where-Object { $_ -ne $pkAttr -and $_ -ne 'createdon' })
                        if ($matchKey -and $fieldNames -notcontains $matchKey) { $fieldNames = @($matchKey) + $fieldNames }
                        $records = Get-SourceRecordsPaged -Conn $SourceConn -EntityLogicalName $EntityLogicalName -Fields $fieldNames -PageSize $pageSize -Logger $LogInfo
                        $total = [int]@($records).Count
                        if ($LogInfo) { & $LogInfo ('  [INFO] Brak metadanych – odkryto ' + $migratableAttrs.Count + ' pol z pierwszego rekordu; pobrano ponownie ' + $total + ' rekordow.') }
                    }
                }
            } catch {
                if ($LogInfo) { & $LogInfo ('  [INFO] Nie udalo sie odkryc pol z Get-CrmRecord: ' + $_.Exception.Message) }
            }
        }
    }

    $created = 0
    $updated = 0
    $failed = 0
    $skipped = 0
    $idx = 0
    $entityAttrExclude = @{}  # atrybuty ktorych nie ma w celu (po bledzie "doesn't contain attribute")
    $entityAttrRequiresOptionSet = @{}  # atrybuty ktore w celu wymagaja OptionSetValue
    $entityAttrRequiresBoolean = @{}     # atrybuty ktore w celu wymagaja Boolean
    $entityAttrRequiresLookup = @{}      # atrybuty Lookup/Customer/Owner -> docelowa encja (np. 'account' lub 'account,contact')

    # Wypelnij typy z metadanych celu, zeby konwertowac wartosci przed pierwszym Create/Update (bez czekania na blad)
    if ($TargetMeta -and $TargetMeta.ContainsKey($EntityLogicalName) -and $TargetMeta[$EntityLogicalName].Attributes) {
        foreach ($a in $TargetMeta[$EntityLogicalName].Attributes) {
            $ln = [string]$a.LogicalName.ToLowerInvariant()
            $typeStr = [string]$a.Type
            if ($typeStr -eq 'OptionSet' -or $typeStr -eq 'Picklist' -or $typeStr -eq 'State' -or $typeStr -eq 'Status') {
                $entityAttrRequiresOptionSet[$ln] = $true
            } elseif ($typeStr -eq 'Boolean' -or $typeStr -eq 'Bit') {
                $entityAttrRequiresBoolean[$ln] = $true
            } elseif ($typeStr -eq 'Lookup' -or $typeStr -eq 'Customer' -or $typeStr -eq 'Owner') {
                $targets = $a.Target
                if ([string]::IsNullOrWhiteSpace($targets)) {
                    if ($typeStr -eq 'Owner') { $targets = 'systemuser,team' } else { $targets = 'account' }
                }
                $entityAttrRequiresLookup[$ln] = $targets
            }
        }
    }

    # Wymuszenie tablicy – niektore zwroty API moga enumerowac sie do zera w foreach
    $recordsToProcess = @($records)
    $maxRecords = 0
    if ($Config.MaxRecordsPerEntity -and [int]$Config.MaxRecordsPerEntity -gt 0) {
        $maxRecords = [int]$Config.MaxRecordsPerEntity
        $recordsToProcess = @($recordsToProcess | Select-Object -First $maxRecords)
    }
    $totalToProcess = $recordsToProcess.Count
    if ($totalToProcess -ne $total -and $LogInfo -and $maxRecords -eq 0) {
        & $LogInfo ('  [WARN] Liczba rekordow do petli (' + $totalToProcess + ') rozni sie od total (' + $total + ').')
    }
    if ($totalToProcess -eq 0) {
        if ($LogInfo) { & $LogInfo '  [WARN] Lista rekordow pusta – brak iteracji.' }
        return @{ Created = 0; Updated = 0; Skipped = 0; Failed = 0; Total = $total }
    }
    if ($LogInfo) {
        if ($maxRecords -gt 0) { & $LogInfo ('  Rozpoczynam petle migracji: ' + $totalToProcess + ' rekordow (limit MaxRecordsPerEntity=' + $maxRecords + ').') }
        else { & $LogInfo ('  Rozpoczynam petle migracji: ' + $totalToProcess + ' rekordow.') }
    }

    foreach ($rec in $recordsToProcess) {
        $idx++
        $pct = if ($total -gt 0) { [int](($idx / $total) * 100) } else { 0 }
        Write-Progress -Activity "Migracja encji: $EntityLogicalName" -Status "Rekord $idx z $total | utworzono: $created, zaktualizowano: $updated, bledy: $failed" -PercentComplete $pct
        $recHash = @{}
        if ($rec -is [Microsoft.Xrm.Sdk.Entity]) {
            foreach ($a in $rec.Attributes) {
                $recHash[$a.Key] = $a.Value
            }
            if ($rec.Id -ne [guid]::Empty) { $recHash['id'] = $rec.Id }
        } elseif ($rec -is [Hashtable]) {
            $recHash = $rec
        } else {
            $recHash = @{}
            if ($rec.PSObject.Properties['Attributes'] -and $null -ne $rec.Attributes) {
                try {
                    foreach ($a in $rec.Attributes) {
                        $keyRaw = if ($null -ne $a.Key) { $a.Key } elseif ($a.PSObject.Properties['Key']) { $a.Key } elseif ($a.PSObject.Properties['Name']) { $a.Name } else { $null }
                        $val = if ($null -ne $a.Value) { $a.Value } elseif ($a.PSObject.Properties['Value']) { $a.Value } else { $null }
                        $k = if ($keyRaw) { ([string]$keyRaw).ToLowerInvariant() } else { $null }
                        if ($k) { $recHash[$k] = $val }
                    }
                } catch {
                    $attrColl = $rec.Attributes
                    if ($attrColl -is [System.Collections.IDictionary]) {
                        foreach ($k in $attrColl.Keys) { $recHash[[string]$k.ToLowerInvariant()] = $attrColl[$k] }
                    }
                }
            }
            $prefix = 'returnproperty_'
            $rec.PSObject.Properties | ForEach-Object {
                $name = $_.Name
                $val = $_.Value
                $key = if ($name.Length -gt $prefix.Length -and $name.Substring(0, $prefix.Length).ToLowerInvariant() -eq $prefix) {
                    $name.Substring($prefix.Length).ToLowerInvariant()
                } else {
                    $name.ToLowerInvariant()
                }
                if ($key -and $key -ne 'attributes') { $recHash[$key] = $val }
            }
            if (-not $recHash.ContainsKey('id') -and $rec.PSObject.Properties['id']) { $recHash['id'] = $rec.id }
            if (-not $recHash.ContainsKey('id') -and $rec.PSObject.Properties['returnProperty_Id']) { $recHash['id'] = $rec.returnProperty_Id }
            if (-not $recHash.ContainsKey('id') -and $recHash.ContainsKey($pkAttr)) { $recHash['id'] = $recHash[$pkAttr] }
            foreach ($rk in @($recHash.Keys)) {
                if ($rk -match '^[a-z0-9]+_(.+)$') {
                    $short = $Matches[1]
                    if (-not $recHash.ContainsKey($short)) { $recHash[$short] = $recHash[$rk] }
                }
            }
        }
        # Gdy metadane nie zwrocily atrybutow (MigratableAttrs: 0) – tylko pola istniejace w schemacie celu (targetAttrSet)
        if ($idx -eq 1 -and $migratableAttrs.Count -eq 0) {
            if ($targetAttrSet -and $targetAttrSet.Count -gt 0) {
                $exclude = @('id', 'createdon', 'overriddencreatedon') + @($skipFields)
                $migratableAttrs = @( $recHash.Keys | Where-Object {
                    $k = $_
                    $k -ne 'id' -and $k -notlike '*_property' -and $k -notin $exclude -and $targetAttrSet.ContainsKey($k)
                } | ForEach-Object { @{ LogicalName = $_; Type = 'String' } } )
                if ($LogInfo -and $migratableAttrs.Count -gt 0) {
                    & $LogInfo ('  [INFO] Brak atrybutow z metadanych – uzyto ' + $migratableAttrs.Count + ' pol istniejacych w schemacie celu.')
                }
            } elseif ($LogInfo) {
                & $LogInfo ('  [INFO] Brak atrybutow z metadanych i brak listy atrybutow celu – przenoszone beda tylko overriddencreatedon i pola z EntityMeta.')
            }
        }
        $targetAttrs, $createdOn = Convert-RecordToTargetAttributes -SourceRecord $recHash -MigratableAttrs $migratableAttrs -SystemFieldsToSkip $skipFields -LookupIdMap $script:LookupIdMap -EntityLogicalName $EntityLogicalName -SourceConn $SourceConn -TargetConn $TargetConn -Config $Config -EntitiesInScope $EntitiesInScope -CommonEntities $CommonEntities -TargetMeta $TargetMeta -AutoMigrateMissingLookups $AutoMigrateMissingLookups -MigrateSingleRecordIfMissing $migrateSingleRecordIfMissing
        if ($idx -eq 1 -and $LogInfo) {
            $migNames = @($migratableAttrs | ForEach-Object { $_.LogicalName } | Select-Object -First 12) -join ', '
            & $LogInfo ('  [DEBUG rekord 1] MigratableAttrs: ' + $migratableAttrs.Count + '; recHash.Keys: ' + $recHash.Keys.Count + '; targetAttrs.Keys: ' + $targetAttrs.Keys.Count + '; przyklady atr: ' + $migNames)
        }
        if ($targetAttrs.Count -eq 0) { $skipped++; continue }

        $doUpdate = $false
        $existingId = $null
        $srcId = $recHash['id']
        $srcIdStr = if ($srcId) { [string]$srcId } else { '' }
        if (($matchBy -eq 'Id' -or $matchBy -eq 'IdThenName') -and $srcIdStr -and $loadedIdMap[$EntityLogicalName] -and $loadedIdMap[$EntityLogicalName][$srcIdStr]) {
            $existingId = $loadedIdMap[$EntityLogicalName][$srcIdStr]
        }
        # Upsert: jesli nie znaleziono w mapie, sprawdz czy rekord o tym samym PK istnieje na celu (np. poprzednia migracja z overriddencreatedon)
        if ($null -eq $existingId -and $mode -eq 'Upsert' -and $srcIdStr) {
            $g = [guid]::Empty
            if ([guid]::TryParse($srcIdStr, [ref]$g)) {
                $existingId = Get-TargetRecordIdByKey -Conn $TargetConn -EntityLogicalName $EntityLogicalName -KeyAttribute $pkAttr -KeyValue $g
            }
        }
        if ($null -eq $existingId -and ($matchBy -eq 'IdThenName' -or $matchBy -eq 'Name' -or $matchBy -eq 'Custom') -and $matchKey -and $recHash.ContainsKey($matchKey)) {
            if ($idx -eq 1 -and $LogInfo) { & $LogInfo ('  Rekord 1: sprawdzanie czy istnieje w celu (dopasowanie po ' + $matchKey + ')...') }
            $existingId = Get-TargetRecordIdByKey -Conn $TargetConn -EntityLogicalName $EntityLogicalName -KeyAttribute $matchKey -KeyValue $recHash[$matchKey]
        }
        if ($mode -eq 'Update' -or $mode -eq 'Upsert') {
            $doUpdate = ($null -ne $existingId)
            if ($mode -eq 'Update' -and -not $doUpdate) { $skipped++; continue }
        }
        if ($mode -eq 'Create' -and $null -ne $existingId) {
            $skipped++
            continue
        }

        if ($idx -eq 1 -and $LogInfo) {
            $action = if ($doUpdate) { 'Update (istniejacy rekord w celu)' } else { 'Create (nowy rekord)' }
            & $LogInfo ('  Rekord 1: ' + $action + ' – wysylam do celu (moze potrwac)...')
        }

        try {
            if ($doUpdate) {
                $fieldsSet = @{}
                $keysToSend = @($targetAttrs.Keys | Select-Object -Unique)
                if ($null -ne $targetAttrSet) { $keysToSend = @($keysToSend | Where-Object { $targetAttrSet.ContainsKey($_) }) }
                $keysToSend = @($keysToSend | Where-Object { -not $entityAttrExclude.ContainsKey($_) })
                foreach ($k in $keysToSend) {
                    if ($k -eq 'overriddencreatedon') { continue }
                    $safe = Get-SerializableValue -Value $targetAttrs[$k]
                    if ($null -ne $safe -and $safe -is [string]) { $safe = Normalize-ValueByAttributeName -AttributeName $k -Value $safe }
                    if ($null -ne $safe -and $entityAttrRequiresLookup.ContainsKey($k) -and $safe -is [string]) {
                        $guidVal = [guid]::Empty
                        if ([guid]::TryParse([string]$safe.Trim(), [ref]$guidVal)) {
                            $targetEntitiesStr = [string]$entityAttrRequiresLookup[$k]
                            $targetEntities = if ($targetEntitiesStr -match ',') { @($targetEntitiesStr.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }) } else { @($targetEntitiesStr.Trim()) }
                            foreach ($lookupEnt in $targetEntities) {
                                $mapped = Get-MappedLookupId -EntityLogicalName $lookupEnt -SourceId $guidVal
                                if ($null -ne $mapped) { $safe = [Microsoft.Xrm.Sdk.EntityReference]::new($lookupEnt, $mapped); break }
                            }
                        }
                    }
                    if ($null -ne $safe) {
                        if ($entityAttrRequiresOptionSet.ContainsKey($k)) {
                            if ($safe -is [int]) { $safe = [Microsoft.Xrm.Sdk.OptionSetValue]::new($safe) }
                            elseif ($safe -is [string]) { $i = 0; if ([int]::TryParse([string]$safe.Trim(), [ref]$i)) { $safe = [Microsoft.Xrm.Sdk.OptionSetValue]::new($i) } }
                        }
                        if ($entityAttrRequiresBoolean.ContainsKey($k) -and $null -ne $safe -and $safe -isnot [bool]) {
                            $s = [string]$safe.Trim().ToLowerInvariant()
                            $safe = ($s -in 'true','1','yes','t','y')
                        }
                        $fieldsSet[$k] = $safe
                    }
                }
                if ($fieldsSet.Count -gt 0) {
                    $attemptFields = $fieldsSet.Clone()
                    $maxRemove = 50
                    $updateDone = $false
                    $lastUpdateError = $null
                    for ($r = 0; $r -le $maxRemove -and -not $updateDone; $r++) {
                        try {
                            Set-CrmRecord -conn $TargetConn -EntityLogicalName $EntityLogicalName -Id $existingId -Fields $attemptFields -ErrorAction Stop
                            $updated++
                            $updateDone = $true
                            if ($idx -eq 1 -and $LogInfo) { & $LogInfo ('  Rekord 1 zaktualizowany.') }
                        } catch {
                            $errMsg = $_.Exception.Message
                            if ($_.Exception.InnerException) { $errMsg = $errMsg + ' ' + $_.Exception.InnerException.Message }
                            $lastUpdateError = $errMsg
                            if ($LogInfo -and ($idx -le 3 -or $r -eq 0)) {
                                $shortErr = if ($errMsg.Length -gt 200) { $errMsg.Substring(0, 200) + '...' } else { $errMsg }
                                & $LogInfo ('  [REKORD ' + $idx + '] Update blad (proba ' + ($r + 1) + '): ' + $shortErr)
                            }
                            if ($errMsg -match "doesn't contain attribute" -and $errMsg -match "Name\s*=\s*'([^']+)'") {
                                $bad = $Matches[1].ToLowerInvariant()
                                $entityAttrExclude[$bad] = $true
                                $next = @{}
                                foreach ($k in $attemptFields.Keys) { if ($k -ne $bad) { $next[$k] = $attemptFields[$k] } }
                                $attemptFields = $next
                                if ($attemptFields.Count -eq 0) { throw }
                            } elseif ($errMsg -match "Incorrect attribute value type System\.String" -or ($errMsg -match "Error converting attribute value to Property" -and $errMsg -match "OptionSetValue")) {
                                $bad = $null
                                if ($errMsg -match "Parameter name:\s*([\w\.]+)") { $bad = $Matches[1].ToLowerInvariant() }
                                elseif ($errMsg -match "Attribute\s*\[\s*(\w+)\s*\]") { $bad = $Matches[1].ToLowerInvariant() }
                                elseif ($errMsg -match "attribute\s+'([^']+)'") { $bad = $Matches[1].ToLowerInvariant() }
                                elseif ($errMsg -match "Attribute\s+'([^']+)'") { $bad = $Matches[1].ToLowerInvariant() }
                                elseif ($errMsg -match "for\s+attribute\s+'([^']+)'") { $bad = $Matches[1].ToLowerInvariant() }
                                if (-not $bad) {
                                    foreach ($k in $attemptFields.Keys) {
                                        if ($attemptFields[$k] -is [string] -and -not $entityAttrRequiresOptionSet.ContainsKey($k) -and -not $entityAttrRequiresBoolean.ContainsKey($k)) { $bad = $k; break }
                                    }
                                }
                                if ($bad -and $attemptFields.ContainsKey($bad)) {
                                    $entityAttrRequiresOptionSet[$bad] = $true
                                    $cur = $attemptFields[$bad]
                                    $converted = $false
                                    if ($cur -is [Microsoft.Xrm.Sdk.OptionSetValue]) {
                                        $attemptFields[$bad] = [Microsoft.Xrm.Sdk.OptionSetValue]::new($cur.Value)
                                        $converted = $true
                                    } elseif ($cur -is [int]) {
                                        $attemptFields[$bad] = [Microsoft.Xrm.Sdk.OptionSetValue]::new($cur)
                                        $converted = $true
                                    } else {
                                        $i = 0
                                        if ($null -ne $cur -and [int]::TryParse([string]$cur.Trim(), [ref]$i)) {
                                            $attemptFields[$bad] = [Microsoft.Xrm.Sdk.OptionSetValue]::new($i)
                                            $converted = $true
                                        }
                                    }
                                    if (-not $converted) {
                                        $entityAttrRequiresBoolean[$bad] = $true
                                        $s = [string]$cur
                                        if ($null -ne $s) { $s = $s.Trim().ToLowerInvariant() }
                                        if ($s -in 'true','1','yes','t','y') { $attemptFields[$bad] = $true; $converted = $true }
                                        elseif ($s -in 'false','0','no','f','n','') { $attemptFields[$bad] = $false; $converted = $true }
                                    }
                                    if (-not $converted) {
                                        $next = @{}
                                        foreach ($k in $attemptFields.Keys) { if ($k -ne $bad) { $next[$k] = $attemptFields[$k] } }
                                        $attemptFields = $next
                                        if ($attemptFields.Count -eq 0) { throw }
                                    }
                                } else { throw }
                            } elseif ($errMsg -match "Invalid value type for attribute:\s*(\w+).*Expected Type:\s*owner") {
                                $bad = $Matches[1].ToLowerInvariant()
                                $next = @{}
                                foreach ($k in $attemptFields.Keys) { if ($k -ne $bad) { $next[$k] = $attemptFields[$k] } }
                                $attemptFields = $next
                                if ($attemptFields.Count -eq 0) { throw }
                            } elseif ($errMsg -match "Entity\s+'(\w+)'\s+With Id\s+=\s+([\w\-]+)\s+Does Not Exist") {
                                $missingEntity = $Matches[1]
                                $missingIdStr = $Matches[2]
                                $pullAttr = $null
                                $sourceIdForPull = $null
                                foreach ($k in $attemptFields.Keys) {
                                    $v = $attemptFields[$k]
                                    if ($v -is [Microsoft.Xrm.Sdk.EntityReference] -and [string]$v.LogicalName -eq $missingEntity -and ([string]$v.Id -eq $missingIdStr -or [string]$v.Id -replace '-','' -eq $missingIdStr -replace '-','')) {
                                        $pullAttr = $k
                                        $sv = $null
                                        if ($recHash.ContainsKey($k)) { $sv = $recHash[$k] }
                                        elseif ($recHash.ContainsKey($k + '_property')) { $sv = $recHash[$k + '_property'] }
                                        if ($null -ne $sv) {
                                            if ($sv -is [Microsoft.Xrm.Sdk.EntityReference]) { $sourceIdForPull = $sv.Id }
                                            elseif ($sv -is [hashtable] -and $sv.id) { $sourceIdForPull = [guid]$sv.id }
                                            elseif ($sv -is [guid]) { $sourceIdForPull = $sv }
                                            elseif ($sv.PSObject.Properties['Id']) { $sourceIdForPull = [guid]$sv.Id }
                                            elseif ($sv.PSObject.Properties['Value'] -and $sv.Value.PSObject.Properties['Id']) { $sourceIdForPull = [guid]$sv.Value.Id }
                                        }
                                        break
                                    }
                                }
                                if ($pullAttr -and $null -ne $sourceIdForPull -and $AutoMigrateMissingLookups -and $migrateSingleRecordIfMissing -and $EntitiesInScope -and $missingEntity -in $EntitiesInScope) {
                                    try {
                                        $newTargetId = & $migrateSingleRecordIfMissing $missingEntity $sourceIdForPull $migrateSingleRecordIfMissing $true
                                        if ($null -ne $newTargetId -and $newTargetId -is [guid]) {
                                            $attemptFields[$pullAttr] = [Microsoft.Xrm.Sdk.EntityReference]::new($missingEntity, $newTargetId)
                                            if ($LogInfo) { & $LogInfo ('  [REKORD ' + $idx + '] Dociagnieto brakujacy rekord ' + $missingEntity + ' ze zrodla – ponawiam Update.') }
                                            continue
                                        }
                                    } catch { }
                                }
                                if ($pullAttr -and $null -eq $sourceIdForPull -and $LogInfo) { & $LogInfo ('  [REKORD ' + $idx + '] Brak source Id w recHash dla atrybutu ' + $pullAttr + ' – nie mozna dociagnac rekordu ' + $missingEntity + '.') }
                                $next = @{}
                                foreach ($k in $attemptFields.Keys) {
                                    $v = $attemptFields[$k]
                                    $remove = $false
                                    if ($v -is [Microsoft.Xrm.Sdk.EntityReference]) {
                                        if ([string]$v.LogicalName -eq $missingEntity) { $remove = $true }
                                        else {
                                        $vid = [string]$v.Id
                                        if ($vid -eq $missingIdStr -or $vid.ToLowerInvariant() -eq $missingIdStr.ToLowerInvariant()) { $remove = $true }
                                    }
                                    } else {
                                        try {
                                            if ($null -ne $v -and $v.PSObject.Properties['Id'] -and $v.PSObject.Properties['LogicalName']) {
                                                $vid = [string]$v.Id
                                                if ([string]$v.LogicalName -eq $missingEntity -or $vid.ToLowerInvariant() -eq $missingIdStr.ToLowerInvariant()) { $remove = $true }
                                            }
                                        } catch { }
                                    }
                                    if ($remove) {
                                        $entityAttrExclude[$k] = $true
                                        continue
                                    }
                                    $next[$k] = $v
                                }
                                $attemptFields = $next
                                if ($attemptFields.Count -eq 0) { throw }
                            } elseif (($errMsg -match "Type Mismatch" -and $errMsg -match "System\.Boolean" -and $errMsg -match "System\.String" -and $errMsg -match "Attribute:\s*\w+\.(\w+)\s+is") -or ($errMsg -match "Error converting attribute value to Property" -and $errMsg -match "Attribute type\s*\[\s*bit\s*\]" -and $errMsg -match "value of type\s*\[\s*System\.String\s*\]" -and $errMsg -match "Attribute\s*\[\s*(\w+)\s*\]")) {
                                $bad = $Matches[1].ToLowerInvariant()
                                $entityAttrRequiresBoolean[$bad] = $true
                                $cur = $attemptFields[$bad]
                                if ($cur -is [bool]) {
                                    $attemptFields[$bad] = [bool]$cur
                                } else {
                                    $s = [string]$cur
                                    if ($null -ne $s) { $s = $s.Trim().ToLowerInvariant() }
                                    if ($s -in 'true','1','yes','t','y') { $attemptFields[$bad] = $true }
                                    elseif ($s -in 'false','0','no','f','n','') { $attemptFields[$bad] = $false }
                                    else {
                                        $next = @{}
                                        foreach ($k in $attemptFields.Keys) { if ($k -ne $bad) { $next[$k] = $attemptFields[$k] } }
                                        $attemptFields = $next
                                        if ($attemptFields.Count -eq 0) { throw }
                                    }
                                }
                            } else {
                                Invoke-Retry -Action {
                                    Set-CrmRecord -conn $TargetConn -EntityLogicalName $EntityLogicalName -Id $existingId -Fields $attemptFields -ErrorAction Stop
                                } -MaxRetries $maxRetry -DelaySeconds $retryDelay -Logger $LogInfo | Out-Null
                                $updated++
                                $updateDone = $true
                            }
                        }
                    }
                    if (-not $updateDone) {
                        $failed++
                        if ($LogInfo) {
                            $shortErr = if ($lastUpdateError -and $lastUpdateError.Length -gt 250) { $lastUpdateError.Substring(0, 250) + '...' } elseif ($lastUpdateError) { $lastUpdateError } else { '(brak komunikatu)' }
                            & $LogInfo ('  [REKORD ' + $idx + '] Update NIE POWIODL SIE po wszystkich probach – blad: ' + $shortErr)
                        }
                    }
                } else {
                    $skipped++
                    if ($LogInfo) { & $LogInfo ('  Pominieto aktualizacje rekordu ' + $idx + ' (brak pol do ustawienia, rekord w zrodle mogl nie zawierac atrybutow)') }
                }
                if ($recHash['id']) {
                    Initialize-LookupMap -EntityLogicalName $EntityLogicalName -SourceId $recHash['id'] -TargetId $existingId
                }
            } else {
                $fieldsCreate = @{}
                $keysToSend = @($targetAttrs.Keys | Select-Object -Unique)
                if ($null -ne $targetAttrSet) { $keysToSend = @($keysToSend | Where-Object { $targetAttrSet.ContainsKey($_) }) }
                $keysToSend = @($keysToSend | Where-Object { -not $entityAttrExclude.ContainsKey($_) })
                foreach ($k in $keysToSend) {
                    $safe = Get-SerializableValue -Value $targetAttrs[$k]
                    if ($null -ne $safe -and $safe -is [string]) { $safe = Normalize-ValueByAttributeName -AttributeName $k -Value $safe }
                    if ($null -ne $safe -and $entityAttrRequiresLookup.ContainsKey($k) -and $safe -is [string]) {
                        $guidVal = [guid]::Empty
                        if ([guid]::TryParse([string]$safe.Trim(), [ref]$guidVal)) {
                            $targetEntitiesStr = [string]$entityAttrRequiresLookup[$k]
                            $targetEntities = if ($targetEntitiesStr -match ',') { @($targetEntitiesStr.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ }) } else { @($targetEntitiesStr.Trim()) }
                            foreach ($lookupEnt in $targetEntities) {
                                $mapped = Get-MappedLookupId -EntityLogicalName $lookupEnt -SourceId $guidVal
                                if ($null -ne $mapped) { $safe = [Microsoft.Xrm.Sdk.EntityReference]::new($lookupEnt, $mapped); break }
                            }
                        }
                    }
                    if ($null -ne $safe) {
                        if ($entityAttrRequiresOptionSet.ContainsKey($k)) {
                            if ($safe -is [int]) { $safe = [Microsoft.Xrm.Sdk.OptionSetValue]::new($safe) }
                            elseif ($safe -is [string]) { $i = 0; if ([int]::TryParse([string]$safe.Trim(), [ref]$i)) { $safe = [Microsoft.Xrm.Sdk.OptionSetValue]::new($i) } }
                        }
                        if ($entityAttrRequiresBoolean.ContainsKey($k) -and $null -ne $safe -and $safe -isnot [bool]) {
                            $s = [string]$safe.Trim().ToLowerInvariant()
                            $safe = ($s -in 'true','1','yes','t','y')
                        }
                        $fieldsCreate[$k] = $safe
                    }
                }
                $attemptFields = $fieldsCreate.Clone()
                $maxRemove = 50
                $newId = $null
                $createDone = $false
                $lastCreateError = $null
                for ($r = 0; $r -le $maxRemove -and -not $createDone; $r++) {
                    try {
                        $newId = & $script:NewCrmRecordCmd -conn $TargetConn -EntityLogicalName $EntityLogicalName -Fields $attemptFields -ErrorAction Stop
                        $createDone = $true
                    } catch {
                        $errMsg = $_.Exception.Message
                        if ($_.Exception.InnerException) { $errMsg = $errMsg + ' ' + $_.Exception.InnerException.Message }
                        $lastCreateError = $errMsg
                        if ($LogInfo -and ($idx -le 3 -or $r -eq 0)) {
                            $shortErr = if ($errMsg.Length -gt 200) { $errMsg.Substring(0, 200) + '...' } else { $errMsg }
                            & $LogInfo ('  [REKORD ' + $idx + '] Create blad (proba ' + ($r + 1) + '): ' + $shortErr)
                        }
                        if ($errMsg -match "doesn't contain attribute" -and $errMsg -match "Name\s*=\s*'([^']+)'") {
                            $bad = $Matches[1].ToLowerInvariant()
                            $entityAttrExclude[$bad] = $true
                            $next = @{}
                            foreach ($k in $attemptFields.Keys) { if ($k -ne $bad) { $next[$k] = $attemptFields[$k] } }
                            $attemptFields = $next
                            if ($attemptFields.Count -eq 0) { throw }
                        } elseif ($errMsg -match "Incorrect attribute value type System\.String" -or $errMsg -match "Incorrect attribute value type Microsoft\.Xrm\.Sdk\.OptionSetValue" -or $errMsg -match "Invalid Attribute Value Type for" -or ($errMsg -match "Error converting attribute value to Property" -and $errMsg -match "OptionSetValue")) {
                            $bad = $null
                            if ($errMsg -match "Parameter name:\s*([\w\.]+)") { $bad = $Matches[1].ToLowerInvariant() }
                            elseif ($errMsg -match "Attribute\s*\[\s*(\w+)\s*\]") { $bad = $Matches[1].ToLowerInvariant() }
                            elseif ($errMsg -match "attribute\s+'([^']+)'") { $bad = $Matches[1].ToLowerInvariant() }
                            elseif ($errMsg -match "Attribute\s+'([^']+)'") { $bad = $Matches[1].ToLowerInvariant() }
                            elseif ($errMsg -match "for\s+attribute\s+'([^']+)'") { $bad = $Matches[1].ToLowerInvariant() }
                            elseif ($errMsg -match "Invalid Attribute Value Type for\s+(\w+)") { $bad = $Matches[1].ToLowerInvariant() }
                            if (-not $bad) {
                                foreach ($k in $attemptFields.Keys) {
                                    if ($attemptFields[$k] -is [string] -and -not $entityAttrRequiresOptionSet.ContainsKey($k) -and -not $entityAttrRequiresBoolean.ContainsKey($k)) { $bad = $k; break }
                                }
                            }
                            if ($bad -and $attemptFields.ContainsKey($bad)) {
                                $cur = $attemptFields[$bad]
                                if ($errMsg -match "Expected:\s*Double|Expected:\s*Decimal") {
                                    $d = [decimal]::Zero
                                    if ($null -ne $cur -and [decimal]::TryParse([string]$cur.Trim().Replace(',', '.'), [ref]$d)) {
                                        $attemptFields[$bad] = $d
                                    } else {
                                        $next = @{}
                                        foreach ($k in $attemptFields.Keys) { if ($k -ne $bad) { $next[$k] = $attemptFields[$k] } }
                                        $attemptFields = $next
                                        if ($attemptFields.Count -eq 0) { throw }
                                    }
                                } else {
                                    $entityAttrRequiresOptionSet[$bad] = $true
                                    $converted = $false
                                    if ($cur -is [Microsoft.Xrm.Sdk.OptionSetValue]) {
                                        $attemptFields[$bad] = [Microsoft.Xrm.Sdk.OptionSetValue]::new($cur.Value)
                                        $converted = $true
                                    } elseif ($cur -is [int]) {
                                        $attemptFields[$bad] = [Microsoft.Xrm.Sdk.OptionSetValue]::new($cur)
                                        $converted = $true
                                    } else {
                                        $i = 0
                                        if ($null -ne $cur -and [int]::TryParse([string]$cur.Trim(), [ref]$i)) {
                                            $attemptFields[$bad] = [Microsoft.Xrm.Sdk.OptionSetValue]::new($i)
                                            $converted = $true
                                        }
                                    }
                                    if (-not $converted) {
                                        $entityAttrRequiresBoolean[$bad] = $true
                                        $s = [string]$cur
                                        if ($null -ne $s) { $s = $s.Trim().ToLowerInvariant() }
                                        if ($s -in 'true','1','yes','t','y') { $attemptFields[$bad] = $true; $converted = $true }
                                        elseif ($s -in 'false','0','no','f','n','') { $attemptFields[$bad] = $false; $converted = $true }
                                    }
                                    if (-not $converted) {
                                        $next = @{}
                                        foreach ($k in $attemptFields.Keys) { if ($k -ne $bad) { $next[$k] = $attemptFields[$k] } }
                                        $attemptFields = $next
                                        if ($attemptFields.Count -eq 0) { throw }
                                    }
                                }
                            } else { throw }
                        } elseif ($errMsg -match "Invalid value type for attribute:\s*(\w+).*Expected Type:\s*lookup.*Actual Type:\s*System\.String") {
                            $bad = $Matches[1].ToLowerInvariant()
                            if ($attemptFields.ContainsKey($bad)) {
                                $cur = $attemptFields[$bad]
                                $guidVal = [guid]::Empty
                                if ($cur -is [guid]) { $guidVal = $cur }
                                elseif ($null -ne $cur -and [guid]::TryParse([string]$cur.Trim(), [ref]$guidVal)) { }
                                if ($guidVal -ne [guid]::Empty) {
                                    $lookupEntity = if ($bad -eq 'administratorid' -or $bad -eq 'ownerid') { 'systemuser' } elseif ($bad -eq 'transactioncurrencyid') { 'transactioncurrency' } elseif ($bad -eq 'originatingleadid') { 'lead' } elseif ($bad -eq 'opportunityid') { 'opportunity' } elseif ($bad -eq 'parentaccountid') { 'account' } elseif ($bad -eq 'parentcontactid') { 'contact' } else { $bad -replace 'id$','' }
                                    $attemptFields[$bad] = [Microsoft.Xrm.Sdk.EntityReference]::new($lookupEntity, $guidVal)
                                } else {
                                    $next = @{}
                                    foreach ($k in $attemptFields.Keys) { if ($k -ne $bad) { $next[$k] = $attemptFields[$k] } }
                                    $attemptFields = $next
                                    if ($attemptFields.Count -eq 0) { throw }
                                }
                            } else { throw }
                        } elseif ($errMsg -match "Invalid value type for attribute:\s*(\w+).*Expected Type:\s*owner") {
                            $bad = $Matches[1].ToLowerInvariant()
                            $next = @{}
                            foreach ($k in $attemptFields.Keys) { if ($k -ne $bad) { $next[$k] = $attemptFields[$k] } }
                            $attemptFields = $next
                            if ($attemptFields.Count -eq 0) { throw }
                        } elseif ($errMsg -match "Entity\s+'(\w+)'\s+With Id\s+=\s+([\w\-]+)\s+Does Not Exist") {
                            $missingEntity = $Matches[1]
                            $missingIdStr = $Matches[2]
                            $pullAttr = $null
                            $sourceIdForPull = $null
                            foreach ($k in $attemptFields.Keys) {
                                $v = $attemptFields[$k]
                                if ($v -is [Microsoft.Xrm.Sdk.EntityReference] -and [string]$v.LogicalName -eq $missingEntity -and ([string]$v.Id -eq $missingIdStr -or [string]$v.Id -replace '-','' -eq $missingIdStr -replace '-','')) {
                                    $pullAttr = $k
                                    $sv = $null
                                    if ($recHash.ContainsKey($k)) { $sv = $recHash[$k] }
                                    elseif ($recHash.ContainsKey($k + '_property')) { $sv = $recHash[$k + '_property'] }
                                    if ($null -ne $sv) {
                                        if ($sv -is [Microsoft.Xrm.Sdk.EntityReference]) { $sourceIdForPull = $sv.Id }
                                        elseif ($sv -is [hashtable] -and $sv.id) { $sourceIdForPull = [guid]$sv.id }
                                        elseif ($sv -is [guid]) { $sourceIdForPull = $sv }
                                        elseif ($sv.PSObject.Properties['Id']) { $sourceIdForPull = [guid]$sv.Id }
                                        elseif ($sv.PSObject.Properties['Value'] -and $sv.Value.PSObject.Properties['Id']) { $sourceIdForPull = [guid]$sv.Value.Id }
                                    }
                                    break
                                }
                            }
                            if ($pullAttr -and $null -ne $sourceIdForPull -and $AutoMigrateMissingLookups -and $migrateSingleRecordIfMissing -and $EntitiesInScope -and $missingEntity -in $EntitiesInScope) {
                                try {
                                    $newTargetId = & $migrateSingleRecordIfMissing $missingEntity $sourceIdForPull $migrateSingleRecordIfMissing $true
                                    if ($null -ne $newTargetId -and $newTargetId -is [guid]) {
                                        $attemptFields[$pullAttr] = [Microsoft.Xrm.Sdk.EntityReference]::new($missingEntity, $newTargetId)
                                        if ($LogInfo) { & $LogInfo ('  [REKORD ' + $idx + '] Dociagnieto brakujacy rekord ' + $missingEntity + ' ze zrodla – ponawiam Create.') }
                                        continue
                                    }
                                } catch { }
                            }
                            if ($pullAttr -and $null -eq $sourceIdForPull -and $LogInfo) { & $LogInfo ('  [REKORD ' + $idx + '] Brak source Id w recHash dla atrybutu ' + $pullAttr + ' – nie mozna dociagnac rekordu ' + $missingEntity + '.') }
                            $next = @{}
                            foreach ($k in $attemptFields.Keys) {
                                $v = $attemptFields[$k]
                                $remove = $false
                                if ($v -is [Microsoft.Xrm.Sdk.EntityReference]) {
                                    if ([string]$v.LogicalName -eq $missingEntity) { $remove = $true }
                                    else {
                                        $vid = [string]$v.Id
                                        if ($vid -eq $missingIdStr -or $vid.ToLowerInvariant() -eq $missingIdStr.ToLowerInvariant()) { $remove = $true }
                                    }
                                } else {
                                    try {
                                        if ($null -ne $v -and $v.PSObject.Properties['Id'] -and $v.PSObject.Properties['LogicalName']) {
                                            $vid = [string]$v.Id
                                            if ([string]$v.LogicalName -eq $missingEntity -or $vid.ToLowerInvariant() -eq $missingIdStr.ToLowerInvariant()) { $remove = $true }
                                        }
                                    } catch { }
                                }
                                if ($remove) {
                                    $entityAttrExclude[$k] = $true
                                    continue
                                }
                                $next[$k] = $v
                            }
                            $attemptFields = $next
                            if ($attemptFields.Count -eq 0) { throw }
                        } elseif (($errMsg -match "Type Mismatch" -and $errMsg -match "System\.Boolean" -and $errMsg -match "System\.String" -and $errMsg -match "Attribute:\s*\w+\.(\w+)\s+is") -or ($errMsg -match "Error converting attribute value to Property" -and $errMsg -match "Attribute type\s*\[\s*bit\s*\]" -and $errMsg -match "value of type\s*\[\s*System\.String\s*\]" -and $errMsg -match "Attribute\s*\[\s*(\w+)\s*\]")) {
                            $bad = $Matches[1].ToLowerInvariant()
                            $entityAttrRequiresBoolean[$bad] = $true
                            $cur = $attemptFields[$bad]
                            if ($cur -is [bool]) {
                                $attemptFields[$bad] = [bool]$cur
                            } else {
                                $s = [string]$cur
                                if ($null -ne $s) { $s = $s.Trim().ToLowerInvariant() }
                                if ($s -in 'true','1','yes','t','y') { $attemptFields[$bad] = $true }
                                elseif ($s -in 'false','0','no','f','n','') { $attemptFields[$bad] = $false }
                                else {
                                    $next = @{}
                                    foreach ($k in $attemptFields.Keys) { if ($k -ne $bad) { $next[$k] = $attemptFields[$k] } }
                                    $attemptFields = $next
                                    if ($attemptFields.Count -eq 0) { throw }
                                }
                            }
                        } else {
                            $newId = Invoke-Retry -Action {
                                & $script:NewCrmRecordCmd -conn $TargetConn -EntityLogicalName $EntityLogicalName -Fields $attemptFields -ErrorAction Stop
                            } -MaxRetries $maxRetry -DelaySeconds $retryDelay -Logger $LogInfo
                            $createDone = $true
                        }
                    }
                }
                if (-not $createDone -and $null -eq $newId) {
                    $failed++
                    if ($LogInfo) {
                        $shortErr = if ($lastCreateError -and $lastCreateError.Length -gt 250) { $lastCreateError.Substring(0, 250) + '...' } elseif ($lastCreateError) { $lastCreateError } else { '(brak komunikatu)' }
                        & $LogInfo ('  [REKORD ' + $idx + '] Create NIE POWIODL SIE po wszystkich probach – blad: ' + $shortErr)
                    }
                }
                if ($null -ne $newId -and $recHash['id']) {
                    Initialize-LookupMap -EntityLogicalName $EntityLogicalName -SourceId $recHash['id'] -TargetId $newId
                }
                if ($null -ne $newId) {
                    $created++
                    if ($idx -eq 1 -and $LogInfo) { & $LogInfo ('  Rekord 1 utworzony.') }
                }
            }
        } catch {
            $msg = $_.Exception.Message
            if ($msg -match 'Value cannot be null|ArgumentNullException|cannot be NULL|RequiredFieldValidator') {
                $skipped++
                if ($LogInfo) { & $LogInfo ('  Pominieto rekord ' + $idx + ' (wymagane pole null): ' + $msg) }
            } else {
                $failed++
                if ($LogInfo) { & $LogInfo ('  Blad rekord ' + $idx + ' : ' + $_) }
            }
        }
        if ($LogInfo -and ($idx % 10 -eq 0)) { & $LogInfo ('  Postep: ' + $idx + '/' + $total + ' | utworzono: ' + $created + ', zaktualizowano: ' + $updated + ', pominieto: ' + $skipped + ', bledy: ' + $failed) }
    }
    Write-Progress -Activity "Migracja encji: $EntityLogicalName" -Completed

    return @{ Created = $created; Updated = $updated; Skipped = $skipped; Failed = $failed; Total = $total }
}

# Usuwanie wszystkich rekordow encji w celu (do czyszczenia po zlej migracji)
function Clear-TargetEntityRecords {
    param(
        [Parameter(Mandatory = $true)]
        $Conn,
        [Parameter(Mandatory = $true)]
        [string] $EntityLogicalName,
        [Parameter(Mandatory = $false)]
        [scriptblock] $Logger
    )
    $pkAttr = Get-PrimaryKeyAttributeName -EntityLogicalName $EntityLogicalName
    $pageSize = 5000
    $allIds = [System.Collections.ArrayList]::new()
    $page = 1
    $cookie = $null
    do {
        if ($page -eq 1) {
            $q = [char]34
            $fetch = '<fetch version=' + $q + '1.0' + $q + ' mapping=' + $q + 'logical' + $q + ' count=' + $q + $pageSize + $q + ' page=' + $q + '1' + $q + ' no-lock=' + $q + 'true' + $q + '><entity name=' + $q + $EntityLogicalName + $q + '><attribute name=' + $q + $pkAttr + $q + ' /></entity></fetch>'
        } else {
            $encCookie = [System.Web.HttpUtility]::UrlEncode($cookie)
            $fetch = '<fetch version=' + $q + '1.0' + $q + ' mapping=' + $q + 'logical' + $q + ' count=' + $q + $pageSize + $q + ' page=' + $q + $page + $q + ' paging-cookie=' + $q + $encCookie + $q + ' no-lock=' + $q + 'true' + $q + '><entity name=' + $q + $EntityLogicalName + $q + '><attribute name=' + $q + $pkAttr + $q + ' /></entity></fetch>'
        }
        try {
            $result = Get-CrmRecordsByFetch -conn $Conn -Fetch $fetch -ErrorAction Stop
        } catch {
            if ($Logger) { & $Logger ('  Blad pobierania ' + $EntityLogicalName + ' : ' + $_) }
            return 0
        }
        $records = $result.CrmRecords
        if (-not $records -or $records.Count -eq 0) { break }
        foreach ($r in $records) {
            $id = $null
            if ($r.PSObject.Properties[$pkAttr]) { $id = $r.$pkAttr }
            if (-not $id -and $r.PSObject.Properties['returnProperty_Id']) { $id = $r.returnProperty_Id }
            if (-not $id -and $r.PSObject.Properties['id']) { $id = $r.id }
            if ($id) {
                $guid = if ($id -is [guid]) { $id } else { [guid]$id }
                [void]$allIds.Add($guid)
            }
        }
        $cookie = $result.PagingCookie
        $more = $result.NextPage -eq $true
        $page++
    } while ($more -and $records.Count -gt 0)
    $total = $allIds.Count
    if ($total -eq 0) {
        if ($Logger) { & $Logger ('  Brak rekordow w ' + $EntityLogicalName + '.') }
        return 0
    }
    $removeCmd = Get-Command -Name Remove-CrmRecord -ErrorAction SilentlyContinue
    if (-not $removeCmd) {
        if ($Logger) { & $Logger '  Modul nie udostepnia Remove-CrmRecord - nie mozna usunac.' }
        return 0
    }
    $deleted = 0
    foreach ($rid in $allIds) {
        try {
            Remove-CrmRecord -conn $Conn -EntityLogicalName $EntityLogicalName -Id $rid -ErrorAction Stop
            $deleted++
        } catch {
            if ($Logger) { & $Logger ('  Blad usuniecia ' + $EntityLogicalName + ' ' + $rid + ' : ' + $_) }
        }
    }
    if ($Logger) { & $Logger ('  Usunieto ' + $deleted + ' / ' + $total + ' rekordow z ' + $EntityLogicalName + '.') }
    return $deleted
}


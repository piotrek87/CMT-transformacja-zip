# Moduł metadanych encji Dataverse
# Wykrywa encje wspólne i zwraca listę atrybutów do migracji

function Get-EntityMetadataFromEnv {
    <#
    .SYNOPSIS
        Pobiera metadane encji ze środowiska (lista encji i atrybutów).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Connection,
        [Parameter(Mandatory = $false)]
        [string[]] $EntityLogicalNames = $null
    )

    if ($null -eq $Connection) { throw "Get-EntityMetadataFromEnv: Connection nie moze byc null." }
    $entities = @{}
    # Pobierz wszystkie metadane encji (Get-CrmEntityAllMetadata lub iteracja po znanych encjach)
    $allMeta = $null
    try {
        $raw = Get-CrmEntityAllMetadata -conn $Connection -ErrorAction Stop
        if ($null -eq $raw) { $allMeta = @() }
        elseif ($raw -is [Array]) { $allMeta = $raw }
        elseif ($raw -and $raw.PSObject.Properties.Name -contains 'EntityMetadata') { $allMeta = @($raw.EntityMetadata) }
        else { $allMeta = @($raw) }
    } catch {
        # Fallback: pobierz metadane pojedynczo dla typowych encji biznesowych
        $defaultEntities = @('account','contact','lead','opportunity','systemuser','team','activitypointer','email','task','appointment')
        foreach ($logicalName in $defaultEntities) {
            if ($EntityLogicalNames -and $EntityLogicalNames.Count -gt 0 -and $logicalName -notin $EntityLogicalNames) { continue }
            try {
                $meta = Get-CrmEntityMetadata -Conn $Connection -EntityLogicalName $logicalName -ErrorAction Stop
                if ($meta -and $meta.Attributes) {
                    $attrs = @()
                    foreach ($attr in $meta.Attributes) {
                        if ($attr.IsLogical -eq $false -and ($attr.IsValidForCreate -eq $true -or $attr.IsValidForUpdate -eq $true)) {
                            $attrs += @{ LogicalName = $attr.LogicalName; Type = $attr.AttributeType; IsLookup = ($attr.AttributeType -eq 'Lookup'); Target = ($attr.Targets -join ',') }
                        }
                    }
                    if ($attrs.Count -eq 0) {
                        foreach ($attr in $meta.Attributes) {
                            if ($attr.IsLogical -eq $false) {
                                $attrs += @{ LogicalName = $attr.LogicalName; Type = $attr.AttributeType; IsLookup = ($attr.AttributeType -eq 'Lookup'); Target = ($attr.Targets -join ',') }
                            }
                        }
                    }
                    $entities[$logicalName] = @{ LogicalName = $logicalName; Attributes = $attrs }
                }
            } catch { }
        }
        return $entities
    }
    foreach ($entityMeta in $allMeta) {
        if ($null -eq $entityMeta) { continue }
        $logicalName = $entityMeta.LogicalName
        if ($null -eq $logicalName) { continue }
        if ($EntityLogicalNames -and $EntityLogicalNames.Count -gt 0 -and $logicalName -notin $EntityLogicalNames) { continue }
        $attrs = @()
        $attrList = $entityMeta.Attributes
        if ($null -eq $attrList) { $attrList = @() }
        foreach ($attr in $attrList) {
            if ($attr.IsLogical -eq $false -and ($attr.IsValidForCreate -eq $true -or $attr.IsValidForUpdate -eq $true)) {
                $attrs += @{ LogicalName = $attr.LogicalName; Type = $attr.AttributeType; IsLookup = ($attr.AttributeType -eq 'Lookup'); Target = ($attr.Targets -join ',') }
            }
        }
        if ($attrs.Count -eq 0) {
            foreach ($attr in $attrList) {
                if ($attr.IsLogical -eq $false) {
                    $attrs += @{ LogicalName = $attr.LogicalName; Type = $attr.AttributeType; IsLookup = ($attr.AttributeType -eq 'Lookup'); Target = ($attr.Targets -join ',') }
                }
            }
        }
        $entities[$logicalName] = @{ LogicalName = $logicalName; Attributes = $attrs }
    }
    # Gdy Get-CrmEntityAllMetadata zwrocil encje z pusta lista atrybutow – pobierz metadane pojedynczo (maja pelny Type ze zrodla)
    foreach ($logicalName in @($entities.Keys)) {
        $ent = $entities[$logicalName]
        if (-not $ent.Attributes -or $ent.Attributes.Count -eq 0) {
            try {
                $meta = Get-CrmEntityMetadata -Conn $Connection -EntityLogicalName $logicalName -ErrorAction Stop
                if ($meta -and $meta.Attributes) {
                    $attrs = @()
                    foreach ($attr in $meta.Attributes) {
                        if ($attr.IsLogical -eq $false -and ($attr.IsValidForCreate -eq $true -or $attr.IsValidForUpdate -eq $true)) {
                            $attrs += @{ LogicalName = $attr.LogicalName; Type = $attr.AttributeType; IsLookup = ($attr.AttributeType -eq 'Lookup'); Target = ($attr.Targets -join ',') }
                        }
                    }
                    if ($attrs.Count -eq 0) {
                        foreach ($attr in $meta.Attributes) {
                            if ($attr.IsLogical -eq $false) {
                                $attrs += @{ LogicalName = $attr.LogicalName; Type = $attr.AttributeType; IsLookup = ($attr.AttributeType -eq 'Lookup'); Target = ($attr.Targets -join ',') }
                            }
                        }
                    }
                    $entities[$logicalName] = @{ LogicalName = $logicalName; Attributes = $attrs }
                }
            } catch { }
        }
    }
    return $entities
}

function Get-CommonEntities {
    <#
    .SYNOPSIS
        Zwraca encje istniejące w obu środowiskach oraz wspólne atrybuty.
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
        # Schemat celu = zrodlo prawdy: przenosimy tylko atrybuty istniejace w celu (z typami z celu).
        # Dzieki temu nigdy nie wysylamy pol, ktorych nie ma w celu (np. roznice rozwiązań źródło/cel).
        $commonAttrs = @($tgtAttrs)
        $common[$name] = @{
            LogicalName = $name
            Attributes  = @($commonAttrs)
        }
    }
    return $common
}

function Get-EntitiesToMigrateOrdered {
    param(
        [hashtable] $CommonEntities,
        [hashtable] $Config,
        [string] $BpfSuffix = 'process'
    )

    $bpfEntities = @()
    $mainEntities = @()
    foreach ($name in $CommonEntities.Keys) {
        if ($name -match "$BpfSuffix`$") {
            $bpfEntities += $name
        } else {
            $mainEntities += $name
        }
    }

    $priority = $Config.EntityOrderPriority
    $ordered = @()
    foreach ($p in $priority) {
        $match = $mainEntities | Where-Object { $_ -eq $p }
        if ($match) { $ordered += $match }
    }
    foreach ($m in $mainEntities) {
        if ($m -notin $ordered) { $ordered += $m }
    }
    foreach ($b in $bpfEntities) {
        $ordered += $b
    }
    return $ordered
}

function Get-EntityRecordCounts {
    <#
    .SYNOPSIS
        Pobiera liczbe rekordow per encja ze zrodla (FetchXML aggregate count).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Connection,
        [Parameter(Mandatory = $true)]
        [string[]] $EntityLogicalNames,
        [Parameter(Mandatory = $false)]
        [scriptblock] $Logger
    )
    $counts = @{}
    $activityEntities = @{ 'activitypointer' = $true; 'email' = $true; 'task' = $true; 'appointment' = $true; 'phonecall' = $true; 'letter' = $true; 'fax' = $true; 'campaignresponse' = $true; 'campaignactivity' = $true }
    $pkSuffix = 'id'
    $aliasName = 'cnt'
    foreach ($ename in $EntityLogicalNames) {
        $pkAttr = if ($activityEntities.ContainsKey($ename)) { 'activityid' } else { $ename + $pkSuffix }
        $fetch = @"
<fetch version="1.0" mapping="logical" aggregate="true">
  <entity name="$ename">
    <attribute name="$pkAttr" aggregate="count" alias="$aliasName" />
  </entity>
</fetch>
"@
        try {
            $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetch -ErrorAction Stop
            $total = $null
            if ($result.CrmRecords -and @($result.CrmRecords).Count -gt 0) {
                $first = $result.CrmRecords[0]
                if ($first.PSObject.Properties.Name -contains $aliasName) { $total = $first.$aliasName }
            }
            if ($null -ne $total -and $total -ge 0) {
                $counts[$ename] = [int]$total
            } else {
                $counts[$ename] = 0
            }
        } catch {
            try {
                $fetch2 = @"
<fetch version="1.0" mapping="logical" aggregate="true">
  <entity name="$ename">
    <attribute name="createdon" aggregate="count" alias="$aliasName" />
  </entity>
</fetch>
"@
                $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetch2 -ErrorAction Stop
                $total = $null
                if ($result.CrmRecords -and @($result.CrmRecords).Count -gt 0 -and $result.CrmRecords[0].PSObject.Properties.Name -contains $aliasName) {
                    $total = $result.CrmRecords[0].$aliasName
                }
                $counts[$ename] = if ($null -ne $total -and $total -ge 0) { [int]$total } else { 0 }
            } catch {
                if ($Logger) { & $Logger "  Liczenie $ename : $_" }
                $counts[$ename] = -1
            }
        }
        if ($Logger -and $counts[$ename] -ge 0) { & $Logger "  $ename : $($counts[$ename]) rekordow" }
    }
    return $counts
}


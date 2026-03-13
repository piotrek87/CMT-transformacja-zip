# Kolejność migracji encji na podstawie zależności (lookup)
# systemuser -> account -> contact -> opportunity -> activities, BPF na końcu

function Get-EntityDependencyGraph {
    <#
    .SYNOPSIS
        Zwraca graf zaleznosci: hashtable encja -> tablica encji, od ktorych zalezy (lookup).
        Uzyte do rozszerzenia EntityExcludeFilter o encje zalezne od wykluczonych.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable] $CommonEntities,
        [Parameter(Mandatory = $true)]
        [hashtable] $Config,
        [Parameter(Mandatory = $false)]
        [string] $BpfSuffix = 'process'
    )
    if ($null -eq $CommonEntities) { $CommonEntities = @{} }
    if ($null -eq $Config) { $Config = @{ EntityDependencyOverrides = @{} } }
    $overrides = if ($Config.EntityDependencyOverrides) { $Config.EntityDependencyOverrides } else { @{} }
    $entityList = @($CommonEntities.Keys)
    $mainEntities = @($entityList | Where-Object { $_ -notmatch "$BpfSuffix`$" })
    $dependsOn = @{}
    foreach ($ename in $mainEntities) {
        $dependsOn[$ename] = @()
        $attrs = $CommonEntities[$ename].Attributes
        if (-not $attrs) { continue }
        foreach ($a in $attrs) {
            if (-not $a.IsLookup -or [string]::IsNullOrWhiteSpace($a.Target)) { continue }
            $targets = $a.Target -split '\s*,\s*'
            foreach ($t in $targets) {
                if ($t -and $t -in $mainEntities -and $t -ne $ename -and $t -notin $dependsOn[$ename]) {
                    $dependsOn[$ename] += $t
                }
            }
        }
        if ($overrides.ContainsKey($ename)) {
            foreach ($t in $overrides[$ename]) {
                if ($t -and $t -in $mainEntities -and $t -ne $ename -and $t -notin $dependsOn[$ename]) {
                    $dependsOn[$ename] += $t
                }
            }
        }
    }
    return $dependsOn
}

function Get-MigrationOrderByDependencies {
    <#
    .SYNOPSIS
        Ustala kolejność migracji encji tak, aby encje referencowane (lookup) były migrowane wcześniej.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable] $CommonEntities,
        [Parameter(Mandatory = $true)]
        [hashtable] $Config,
        [Parameter(Mandatory = $false)]
        [string] $BpfSuffix = 'process'
    )

    if ($null -eq $CommonEntities) { $CommonEntities = @{} }
    if ($null -eq $Config) { $Config = @{ EntityOrderPriority = @(); BpfEntitySuffix = 'process'; EntityDependencyOverrides = @{} } }
    $entityList = @($CommonEntities.Keys)
    $priorityOrder = $Config.EntityOrderPriority
    if ($null -eq $priorityOrder) { $priorityOrder = @() }
    $overrides = if ($Config.EntityDependencyOverrides) { $Config.EntityDependencyOverrides } else { @{} }
    $bpfEntities = @($entityList | Where-Object { $_ -match "$BpfSuffix`$" })
    $mainEntities = @($entityList | Where-Object { $_ -notmatch "$BpfSuffix`$" })

    # Buduj graf zależności: dla każdej encji lista encji od których zależy (target lookup)
    $dependsOn = @{}
    foreach ($ename in $mainEntities) {
        $dependsOn[$ename] = @()
        $attrs = $CommonEntities[$ename].Attributes
        foreach ($a in $attrs) {
            if (-not $a.IsLookup -or [string]::IsNullOrWhiteSpace($a.Target)) { continue }
            $targets = $a.Target -split '\s*,\s*'
            foreach ($t in $targets) {
                if ($t -and $t -in $mainEntities -and $t -ne $ename -and $t -notin $dependsOn[$ename]) {
                    $dependsOn[$ename] += $t
                }
            }
        }
        if ($overrides.ContainsKey($ename)) {
            foreach ($t in $overrides[$ename]) {
                if ($t -and $t -in $mainEntities -and $t -ne $ename -and $t -notin $dependsOn[$ename]) {
                    $dependsOn[$ename] += $t
                }
            }
        }
    }

    # Sortowanie topologiczne (encje bez zależności pierwsze)
    $ordered = @()
    $remaining = [System.Collections.ArrayList]::new($mainEntities)
    $maxIter = $remaining.Count * 2
    $iter = 0
    while ($remaining.Count -gt 0 -and $iter -lt $maxIter) {
        $iter++
        $ready = @()
        foreach ($e in $remaining) {
            $deps = $dependsOn[$e]
            $allMet = ($deps | Where-Object { $_ -notin $ordered }).Count -eq 0
            if ($allMet) { $ready += $e }
        }
        if ($ready.Count -eq 0) {
            # Cykle lub brak zależności - dodaj pozostałe w kolejności priority
            foreach ($p in $priorityOrder) {
                if ($p -in $remaining) { $ready += $p; break }
            }
            if ($ready.Count -eq 0) { $ready = @($remaining | Select-Object -First 1) }
        }
        if ($ready.Count -gt 1 -and $priorityOrder -and $priorityOrder.Count -gt 0) {
            $ready = @($ready | Sort-Object { $pi = [array]::IndexOf($priorityOrder, $_); if ($pi -lt 0) { 9999 } else { $pi } })
        }
        foreach ($r in $ready) {
            $idx = $remaining.IndexOf($r)
            if ($idx -ge 0) { $remaining.RemoveAt($idx) }
            $ordered += $r
        }
    }
    # Na końcu encje BPF
    foreach ($b in $bpfEntities) {
        $ordered += $b
    }
    return $ordered
}

function Get-EntitiesWithRecordsAndDependencies {
    <#
    .SYNOPSIS
        Zwraca liste encji do migracji: tylko te, ktore maja rekordy LUB sa potrzebne do relacji (lookup) dla encji z rekordami. Kolejnosc zachowana wedlug zaleznosci.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable] $CommonEntities,
        [Parameter(Mandatory = $true)]
        [hashtable] $RecordCounts,
        [Parameter(Mandatory = $true)]
        [hashtable] $Config,
        [Parameter(Mandatory = $false)]
        [string] $BpfSuffix = 'process'
    )
    $orderedAll = Get-MigrationOrderByDependencies -CommonEntities $CommonEntities -Config $Config -BpfSuffix $BpfSuffix
    $entityList = @($CommonEntities.Keys)
    $mainEntities = @($entityList | Where-Object { $_ -notmatch "$BpfSuffix`$" })
    $bpfEntities = @($entityList | Where-Object { $_ -match "$BpfSuffix`$" })

    # Dla kazdej encji: od ktorych zalezy (lookup targets)
    $dependsOn = @{}
    foreach ($ename in $mainEntities) {
        $dependsOn[$ename] = @()
        $attrs = $CommonEntities[$ename].Attributes
        if (-not $attrs) { continue }
        foreach ($a in $attrs) {
            if (-not $a.IsLookup -or [string]::IsNullOrWhiteSpace($a.Target)) { continue }
            $targets = $a.Target -split '\s*,\s*'
            foreach ($t in $targets) {
                if ($t -and $t -in $mainEntities -and $t -ne $ename -and $t -notin $dependsOn[$ename]) {
                    $dependsOn[$ename] += $t
                }
            }
        }
    }

    $withRecords = @()
    foreach ($e in $mainEntities) {
        $c = 0
        if ($RecordCounts.ContainsKey($e)) { $c = $RecordCounts[$e] }
        if ($c -gt 0) { $withRecords += $e }
    }
    $toInclude = @($withRecords)
    $changed = $true
    while ($changed) {
        $changed = $false
        foreach ($e in $toInclude) {
            $deps = $dependsOn[$e]
            if (-not $deps) { continue }
            foreach ($d in $deps) {
                if ($d -notin $toInclude) {
                    $toInclude += $d
                    $changed = $true
                }
            }
        }
    }
    $out = @($orderedAll | Where-Object { $_ -in $toInclude })
    return $out
}

function Get-EntitiesFromWhitelistAndDependencies {
    <#
    .SYNOPSIS
        Zwraca liste encji do migracji: tylko te z whitelisty (EntityIncludeOnly) plus encje potrzebne do relacji (lookup). Kolejnosc wedlug zaleznosci.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable] $CommonEntities,
        [Parameter(Mandatory = $true)]
        [string[]] $Whitelist,
        [Parameter(Mandatory = $true)]
        [hashtable] $Config,
        [Parameter(Mandatory = $false)]
        [string] $BpfSuffix = 'process'
    )
    if (-not $Whitelist -or $Whitelist.Count -eq 0) { return @() }
    $orderedAll = Get-MigrationOrderByDependencies -CommonEntities $CommonEntities -Config $Config -BpfSuffix $BpfSuffix
    $entityList = @($CommonEntities.Keys)
    $mainEntities = @($entityList | Where-Object { $_ -notmatch "$BpfSuffix`$" })
    $bpfEntities = @($entityList | Where-Object { $_ -match "$BpfSuffix`$" })
    $dependsOn = @{}
    foreach ($ename in $mainEntities) {
        $dependsOn[$ename] = @()
        $attrs = $CommonEntities[$ename].Attributes
        if (-not $attrs) { continue }
        foreach ($a in $attrs) {
            if (-not $a.IsLookup -or [string]::IsNullOrWhiteSpace($a.Target)) { continue }
            $targets = $a.Target -split '\s*,\s*'
            foreach ($t in $targets) {
                if ($t -and $t -in $mainEntities -and $t -ne $ename -and $t -notin $dependsOn[$ename]) {
                    $dependsOn[$ename] += $t
                }
            }
        }
    }
    $toInclude = @($Whitelist | Where-Object { $_ -in $mainEntities -or $_ -in $bpfEntities })
    $changed = $true
    while ($changed) {
        $changed = $false
        foreach ($e in $toInclude) {
            $deps = $dependsOn[$e]
            if (-not $deps) { continue }
            foreach ($d in $deps) {
                if ($d -notin $toInclude) {
                    $toInclude += $d
                    $changed = $true
                }
            }
        }
    }
    $out = @($orderedAll | Where-Object { $_ -in $toInclude })
    return $out
}


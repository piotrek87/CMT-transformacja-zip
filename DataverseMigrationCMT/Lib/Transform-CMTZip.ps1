# Przetwarza zip CMT: podmiana ownerow (IdMap), zachowanie dat, usuniecie pol nieistniejacych w celu.
# Uzycie: .\Transform-CMTZip.ps1 -InputZipPath C:\...\CMT_Export.zip [-OutputZipPath ...] [-IdMapPath ...] [-TargetConnectionString ...] [-StripFieldsNotInTarget]
# Albo: -ConfigPath ...\CMTConfig.ps1 (wtedy cel i StripFieldsNotInTarget z configu - zalecane przy uruchomieniu z menu).
# Zasada: w stringach tylko ASCII - cudzyslowy " i ', myslnik - (nie en-dash). Zapobiega bledom parsera.

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $InputZipPath,
    [Parameter(Mandatory = $false)]
    [string] $OutputZipPath,
    [Parameter(Mandatory = $false)]
    [string] $IdMapPath,
    [Parameter(Mandatory = $false)]
    [string] $TargetConnectionString,
    [Parameter(Mandatory = $false)]
    [switch] $StripFieldsNotInTarget,
    [Parameter(Mandatory = $false)]
    [string] $ConfigPath
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
if (-not (Test-Path $InputZipPath -PathType Leaf)) {
    throw "Plik zip nie istnieje: $InputZipPath"
}
$script:TransformStartTime = Get-Date
Write-Host "Transformacja start: $($script:TransformStartTime.ToString('dd.MM.yyyy HH:mm:ss'))" -ForegroundColor Gray

if ([string]::IsNullOrWhiteSpace($OutputZipPath)) {
    $dir = [System.IO.Path]::GetDirectoryName($InputZipPath)
    $name = [System.IO.Path]::GetFileNameWithoutExtension($InputZipPath)
    $OutputZipPath = Join-Path $dir ($name + '_ForTarget.zip')
}

# Config: polaczenie do celu + opcjonalnie walidacja option setow (Report/Clear/Replace)
$script:OptionSetValidationAction = 'Report'
$script:OptionSetFallbackValues = @{}
if (-not [string]::IsNullOrWhiteSpace($ConfigPath) -and (Test-Path $ConfigPath -PathType Leaf)) {
    try {
        $cfg = & $ConfigPath
        if ($cfg -and $cfg.TargetConnectionString) {
            $TargetConnectionString = $cfg.TargetConnectionString
            $StripFieldsNotInTarget = $true
            $script:StripFieldsNotInTarget = $true
        }
        if ($cfg -and $cfg.SourceConnectionString) {
            $script:SourceConnectionString = $cfg.SourceConnectionString
        }
        if ($cfg -and $null -ne $cfg.OptionSetValidationAction) {
            $script:OptionSetValidationAction = [string]$cfg.OptionSetValidationAction
        }
        if ($cfg -and $cfg.OptionSetFallbackValues -and $cfg.OptionSetFallbackValues -is [hashtable]) {
            $script:OptionSetFallbackValues = $cfg.OptionSetFallbackValues
        }
        if ($cfg -and $cfg.LookupFieldsToStripFromImport -and $cfg.LookupFieldsToStripFromImport -is [array]) {
            $script:LookupFieldsToStripFromImport = @($cfg.LookupFieldsToStripFromImport)
        }
        if ($cfg -and $cfg.EntitiesToExcludeFromImport -and $cfg.EntitiesToExcludeFromImport -is [array]) {
            $script:EntitiesToExcludeFromImport = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($e in $cfg.EntitiesToExcludeFromImport) { if (-not [string]::IsNullOrWhiteSpace($e)) { [void]$script:EntitiesToExcludeFromImport.Add($e.Trim()) } }
        }
    } catch { Write-Host "Nie udalo sie zaladowac configu: $($_.Message)" -ForegroundColor Yellow }
}
if (-not $script:EntitiesToExcludeFromImport -or $script:EntitiesToExcludeFromImport.Count -eq 0) {
    $script:EntitiesToExcludeFromImport = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
}
if (-not $script:LookupFieldsToStripFromImport -or $script:LookupFieldsToStripFromImport.Count -eq 0) {
    $script:LookupFieldsToStripFromImport = @('msdyn_accountkpiid', 'msdyn_contactkpiid', 'transactioncurrencyid', 'originatingleadid')
}
$script:StripFieldsNotInTarget = $StripFieldsNotInTarget

# IdMap: podmiana GUID ownerow (zrodlo -> cel). Klucze musza byc czystym GUID (w zipie jest value="guid").
# Jesli JSON ma klucze w formacie CRM (np. "guid systemuserid=xxx"), wyciagamy czysty GUID.
# Normalizuje lancuch daty do ISO 8601 (Dataverse/CMT pewnie przyjmuje; format lokalny np. dd.MM.yyyy moze byc odrzucany dla czesci rekordow).
function ConvertTo-DateTimeIso {
    param([string]$DateStr)
    if ([string]::IsNullOrWhiteSpace($DateStr)) { return $DateStr }
    $s = $DateStr.Trim()
    if ($s -match '^\d{4}-\d{2}-\d{2}T') { return $s }
    try {
        $dt = [DateTime]::Parse($s, [System.Globalization.CultureInfo]::InvariantCulture)
        return $dt.ToString('o')
    } catch {
        try {
            $dt = [DateTime]::Parse($s, [System.Globalization.CultureInfo]::GetCultureInfo('pl-PL'))
            return $dt.ToString('o')
        } catch { return $s }
    }
}
function Get-PureGuidFromIdMapKey {
    param([string]$Key)
    if ([string]::IsNullOrWhiteSpace($Key)) { return $null }
    $s = $Key.Trim().ToLowerInvariant()
    if ($s -match '^[0-9a-f\-]{36}$') { return $s }
    if ($s -match '=([0-9a-f\-]{36})$') { return $Matches[1] }
    $s = $s -replace '^\{|\}$', ''
    if ($s -match '^[0-9a-f\-]{36}$') { return $s }
    return $s
}
$guidMap = @{}
if (-not [string]::IsNullOrWhiteSpace($IdMapPath) -and (Test-Path $IdMapPath -PathType Leaf)) {
    $json = [System.IO.File]::ReadAllText($IdMapPath, [System.Text.Encoding]::UTF8)
    $guidMapRaw = $json | ConvertFrom-Json
    if ($guidMapRaw -is [PSCustomObject]) {
        $guidMapRaw.PSObject.Properties | ForEach-Object {
            $val = $_.Value
            if ($val -match '=([0-9a-fA-F\-]{36})$') { $val = $Matches[1] }
            elseif ($val -match '^\{?([0-9a-fA-F\-]{36})\}?$') { $val = $Matches[1] }
            $pureKey = Get-PureGuidFromIdMapKey -Key $_.Name
            if (-not [string]::IsNullOrWhiteSpace($pureKey)) { $guidMap[$pureKey] = $val }
        }
    }
    Write-Host "Zaladowano IdMap: $($guidMap.Count) mapowan (ownerzy)."
} else {
    Write-Host "Brak IdMap - bez podmiany ownerow. Uzyj opcji 2 (User Map) w menu."
}

# IdMap po imieniu i nazwisku (gdy CMT eksportuje ownerid jako tekst, nie GUID)
$displayNameToGuid = @{}
$byDisplayNamePath = $null
if (-not [string]::IsNullOrWhiteSpace($IdMapPath) -and (Test-Path $IdMapPath -PathType Leaf)) {
    $idMapDir = [System.IO.Path]::GetDirectoryName($IdMapPath)
    $byDisplayNamePath = Join-Path $idMapDir 'CMT_IdMap_ByDisplayName.json'
}
if ($null -ne $byDisplayNamePath -and (Test-Path $byDisplayNamePath -PathType Leaf)) {
    $jsonDn = [System.IO.File]::ReadAllText($byDisplayNamePath, [System.Text.Encoding]::UTF8)
    $displayNameToGuid = $jsonDn | ConvertFrom-Json
    if ($displayNameToGuid -is [PSCustomObject]) {
        $h = @{}
        $displayNameToGuid.PSObject.Properties | ForEach-Object { $h[$_.Name.Trim().ToLowerInvariant()] = $_.Value }
        $displayNameToGuid = $h
    }
    Write-Host "Zaladowano IdMap po imieniu i nazwisku: $($displayNameToGuid.Count) mapowan (owner jako tekst -> GUID celu)."
} elseif ($guidMap.Count -gt 0) {
    Write-Host "Brak pliku CMT_IdMap_ByDisplayName.json - uruchom opcje 2 (User Map), zeby generowac mapowanie imie i nazwisko -> GUID celu." -ForegroundColor Yellow
}

# Pola dozwolone w celu (entity -> HashSet atrybutow) - do usuniecia atrybutow nie z celu
$script:targetAllowedAttrs = @{}
$script:connTarget = $null

# Klucz cache metadanych (zgodny z Export-TargetMetadataCache) – po URL celu
function Get-TargetMetadataCacheKeyFromConnStr {
    param([string]$ConnectionString)
    if ([string]::IsNullOrWhiteSpace($ConnectionString)) { return '' }
    $url = ''
    foreach ($part in $ConnectionString.Split(';')) {
        $s = $part.Trim()
        if ($s.StartsWith('Url=', [StringComparison]::OrdinalIgnoreCase)) {
            $val = $s.Substring(4).Trim()
            if ($val.Length -ge 2 -and $val.StartsWith('"') -and $val.EndsWith('"')) { $val = $val.Substring(1, $val.Length - 2).Replace('""', '"') }
            $url = $val.Trim().ToLowerInvariant()
            break
        }
    }
    if ([string]::IsNullOrWhiteSpace($url)) { return 'default' }
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($url)
        $hash = [System.Security.Cryptography.SHA256]::Create().ComputeHash($bytes)
        $hex = [BitConverter]::ToString($hash) -replace '-', ''
        return $hex.Substring(0, [Math]::Min(16, $hex.Length)).ToLowerInvariant()
    } catch {
        return ('u' + [Math]::Abs($url.GetHashCode()).ToString('x8'))
    }
}

$script:entityFiltersAttributes = $null
try { $script:entityFiltersAttributes = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::Attributes } catch { }

$doOptionSetValidation = $script:OptionSetValidationAction -match '^(Report|Clear|Replace|Interactive)$'
$script:doOptionSetValidation = $doOptionSetValidation
$script:overrideCount = 0
$script:ownerReplaceCount = 0
$script:DuplicateRecordsRemovedCount = 0
$script:LookupFieldsStrippedCount = 0
$script:DuplicateOverrideRemovedCount = 0
$script:OptionSetIssues = [System.Collections.ArrayList]::new()
$script:TargetOptionSetAllowed = @{}
$script:OptionSetInteractiveChoice = @{}
$script:SourceConnectionString = $null
$script:connSource = $null
$script:SourceOptionSetCache = @{}
$script:EntitiesFoundInZip = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$script:MetadataCacheLoaded = $false

# Zaladuj cache metadanych celu (po inicjalizacji TargetOptionSetAllowed) – opcja 5 zapisuje, tu odczyt
$outputDirForCache = [System.IO.Path]::GetDirectoryName($OutputZipPath)
if ([string]::IsNullOrWhiteSpace($outputDirForCache)) { $outputDirForCache = (Get-Location).Path }
if (-not [string]::IsNullOrWhiteSpace($TargetConnectionString)) {
    $cacheKey = Get-TargetMetadataCacheKeyFromConnStr -ConnectionString $TargetConnectionString
    if (-not [string]::IsNullOrWhiteSpace($cacheKey)) {
        $metadataCachePath = Join-Path $outputDirForCache ("TargetMetadata_$cacheKey.json")
        if (Test-Path $metadataCachePath -PathType Leaf) {
            try {
                $cacheJson = [System.IO.File]::ReadAllText($metadataCachePath, [System.Text.Encoding]::UTF8)
                $cacheObj = $cacheJson | ConvertFrom-Json
                if ($cacheObj -and $cacheObj.EntityAttributes) {
                    foreach ($p in $cacheObj.EntityAttributes.PSObject.Properties) {
                        $entKey = $p.Name.Trim().ToLowerInvariant()
                        $arr = @($p.Value)
                        $hs = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                        foreach ($a in $arr) { if (-not [string]::IsNullOrWhiteSpace($a)) { [void]$hs.Add([string]$a) } }
                        $script:targetAllowedAttrs[$entKey] = $hs
                    }
                }
                if ($cacheObj -and $cacheObj.OptionSets) {
                    foreach ($entProp in $cacheObj.OptionSets.PSObject.Properties) {
                        $entKey = $entProp.Name.Trim().ToLowerInvariant()
                        $script:TargetOptionSetAllowed[$entKey] = @{}
                        foreach ($attrProp in $entProp.Value.PSObject.Properties) {
                            $o = $attrProp.Value
                            $allowedSet = [System.Collections.Generic.HashSet[int]]::new()
                            $optionsList = [System.Collections.ArrayList]::new()
                            foreach ($v in @($o.Allowed)) { [void]$allowedSet.Add([int]$v) }
                            foreach ($opt in @($o.Options)) {
                                $val = [int]$opt.Value
                                $lbl = if ($opt.Label) { [string]$opt.Label } else { "($val)" }
                                [void]$optionsList.Add([PSCustomObject]@{ Value = $val; Label = $lbl })
                            }
                            if ($allowedSet.Count -gt 0) {
                                $script:TargetOptionSetAllowed[$entKey][$attrProp.Name] = @{ AllowedSet = $allowedSet; Options = $optionsList }
                            }
                        }
                    }
                }
                $entCount = if ($cacheObj.EntityAttributes) { @($cacheObj.EntityAttributes.PSObject.Properties).Count } else { 0 }
                $cacheDateStr = ''
                if ($cacheObj.FetchedAt) {
                    try {
                        $dt = [DateTime]::Parse($cacheObj.FetchedAt)
                        $cacheDateStr = ' Ostatni pobor cache: ' + $dt.ToString('dd.MM.yyyy HH:mm')
                    } catch { $cacheDateStr = ' Ostatni pobor cache: ' + [string]$cacheObj.FetchedAt }
                } else {
                    $fi = Get-Item -LiteralPath $metadataCachePath -ErrorAction SilentlyContinue
                    if ($fi) { $cacheDateStr = ' Plik cache z: ' + $fi.LastWriteTime.ToString('dd.MM.yyyy HH:mm') }
                }
                Write-Host "Zaladowano cache metadanych celu: $metadataCachePath (encji: $entCount). Strip i option sety z cache (bez polaczenia z CRM).$cacheDateStr" -ForegroundColor Green
                $script:MetadataCacheLoaded = $true
            } catch {
                Write-Host "Nie udalo sie zaladowac cache metadanych: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }
    }
}
if ($script:MetadataCacheLoaded) {
    Write-Host "Transformacja uzywa cache metadanych - bez polaczenia z CRM (zadnych zapytan API do celu)." -ForegroundColor Green
} else {
    if (($script:StripFieldsNotInTarget -or $doOptionSetValidation) -and -not [string]::IsNullOrWhiteSpace($TargetConnectionString)) {
        if (Get-Module -ListAvailable -Name 'Microsoft.Xrm.Data.PowerShell') {
            Import-Module Microsoft.Xrm.Data.PowerShell -Force -ErrorAction SilentlyContinue
            try {
                $script:connTarget = Get-CrmConnection -ConnectionString $TargetConnectionString -ErrorAction Stop
                if ($script:StripFieldsNotInTarget) { Write-Host "Polaczenie z celem OK - beda usuniete pola nieistniejace w celu." -ForegroundColor Gray }
                if ($doOptionSetValidation) { Write-Host "Polaczenie z celem OK - walidacja option setow (akcja: $($script:OptionSetValidationAction))." -ForegroundColor Gray }
            } catch {
                Write-Host "Brak polaczenia z celem - pomijam strip pol i walidacje option setow. $($_.Exception.Message)" -ForegroundColor Yellow
                $script:connTarget = $null
                $doOptionSetValidation = $false
            }
        } else {
            Write-Host "Strip/walidacja option setow wymaga Microsoft.Xrm.Data.PowerShell - pomijam." -ForegroundColor Yellow
            $doOptionSetValidation = $false
        }
    }
    if ($script:connTarget) {
        Write-Host "Wskazowka: Aby przyspieszyc nastepna transformacje, uruchom w menu opcje 5 (Pobierz metadane celu) - unikniesz polaczen z CRM." -ForegroundColor Cyan
    }
}
if ($doOptionSetValidation -and ($script:connTarget -or $script:MetadataCacheLoaded)) {
    Write-Host "Walidacja option setow: wlaczona (akcja: $($script:OptionSetValidationAction))." -ForegroundColor Gray
}

function Get-TargetOptionSetAllowedForEntity {
    param([object]$Conn, [string]$EntityLogicalName)
    if ([string]::IsNullOrWhiteSpace($EntityLogicalName)) { return @{} }
    if ($script:TargetOptionSetAllowed.ContainsKey($EntityLogicalName)) {
        return $script:TargetOptionSetAllowed[$EntityLogicalName]
    }
    if (-not $Conn) { return @{} }
    $result = @{}
    try {
        if ($script:entityFiltersAttributes) {
            $meta = Get-CrmEntityMetadata -Conn $Conn -EntityLogicalName $EntityLogicalName -EntityFilters $script:entityFiltersAttributes -ErrorAction Stop
        } else {
            $meta = Get-CrmEntityMetadata -Conn $Conn -EntityLogicalName $EntityLogicalName -ErrorAction Stop
        }
        if (-not $meta -or -not $meta.Attributes) { $script:TargetOptionSetAllowed[$EntityLogicalName] = $result; return $result }
        foreach ($attr in $meta.Attributes) {
            $ln = $attr.LogicalName
            $type = $attr.AttributeType
            if (-not $ln) { continue }
            $isPicklist = ($type -eq 'Picklist' -or $type -eq 'OptionSet' -or $type -eq 14 -or $type -eq 15)
            if (-not $isPicklist) { continue }
            try {
                $req = [Microsoft.Xrm.Sdk.Messages.RetrieveAttributeRequest]::new()
                $req.EntityLogicalName = $EntityLogicalName
                $req.LogicalName = $ln
                $req.RetrieveAsIfPublished = $true
                $resp = $Conn.Execute($req)
                $optionSet = $resp.AttributeMetadata.OptionSet
                if (-not $optionSet -or -not $optionSet.Options) { continue }
                $allowed = [System.Collections.Generic.HashSet[int]]::new()
                $optionsWithLabels = [System.Collections.ArrayList]::new()
                foreach ($opt in $optionSet.Options) {
                    if ($null -eq $opt.Value) { continue }
                    $v = [int]$opt.Value
                    [void]$allowed.Add($v)
                    $label = ''
                    try {
                        if ($opt.Label -and $opt.Label.UserLocalizedLabel -and $opt.Label.UserLocalizedLabel.Label) {
                            $label = [string]$opt.Label.UserLocalizedLabel.Label
                        }
                        if ([string]::IsNullOrWhiteSpace($label) -and $opt.Label -and $opt.Label.LocalizedLabels -and $opt.Label.LocalizedLabels.Count -gt 0) {
                            $label = [string]$opt.Label.LocalizedLabels[0].Label
                        }
                    } catch { }
                    if ([string]::IsNullOrWhiteSpace($label)) { $label = "($v)" }
                    [void]$optionsWithLabels.Add([PSCustomObject]@{ Value = $v; Label = $label })
                }
                if ($allowed.Count -gt 0) {
                    $result[$ln] = @{ AllowedSet = $allowed; Options = $optionsWithLabels }
                }
            } catch {
                Write-Host "  (pomijam option set $EntityLogicalName.$ln : $($_.Message))" -ForegroundColor DarkGray
            }
        }
    } catch {
        Write-Host "  Brak metadanych encji $EntityLogicalName : $($_.Message)" -ForegroundColor Yellow
    }
    $script:TargetOptionSetAllowed[$EntityLogicalName] = $result
    return $result
}

function Get-SourceOptionSetOptions {
    param([object]$Conn, [string]$EntityLogicalName, [string]$AttributeLogicalName)
    if (-not $Conn -or [string]::IsNullOrWhiteSpace($EntityLogicalName) -or [string]::IsNullOrWhiteSpace($AttributeLogicalName)) { return $null }
    $cacheKey = $EntityLogicalName + '|' + $AttributeLogicalName
    if ($script:SourceOptionSetCache -and $script:SourceOptionSetCache.ContainsKey($cacheKey)) {
        return $script:SourceOptionSetCache[$cacheKey]
    }
    $optionsList = [System.Collections.ArrayList]::new()
    try {
        $req = [Microsoft.Xrm.Sdk.Messages.RetrieveAttributeRequest]::new()
        $req.EntityLogicalName = $EntityLogicalName
        $req.LogicalName = $AttributeLogicalName
        $req.RetrieveAsIfPublished = $true
        $resp = $Conn.Execute($req)
        $optionSet = $resp.AttributeMetadata.OptionSet
        if (-not $optionSet -or -not $optionSet.Options) { $script:SourceOptionSetCache[$cacheKey] = $null; return $null }
        foreach ($opt in $optionSet.Options) {
            if ($null -eq $opt.Value) { continue }
            $v = [int]$opt.Value
            $label = ''
            try {
                if ($opt.Label -and $opt.Label.UserLocalizedLabel -and $opt.Label.UserLocalizedLabel.Label) {
                    $label = [string]$opt.Label.UserLocalizedLabel.Label
                }
                if ([string]::IsNullOrWhiteSpace($label) -and $opt.Label -and $opt.Label.LocalizedLabels -and $opt.Label.LocalizedLabels.Count -gt 0) {
                    $label = [string]$opt.Label.LocalizedLabels[0].Label
                }
            } catch { }
            if ([string]::IsNullOrWhiteSpace($label)) { $label = "($v)" }
            [void]$optionsList.Add([PSCustomObject]@{ Value = $v; Label = $label })
        }
    } catch {
        Write-Host "  (zrodlo: brak opcji dla $EntityLogicalName.$AttributeLogicalName : $($_.Message))" -ForegroundColor DarkGray
        $script:SourceOptionSetCache[$cacheKey] = $null
        return $null
    }
    $result = @{ Options = $optionsList }
    if (-not $script:SourceOptionSetCache) { $script:SourceOptionSetCache = @{} }
    $script:SourceOptionSetCache[$cacheKey] = $result
    return $result
}

function Get-OptionSetUserChoice {
    param([string]$Key, [object]$AttrInfo, [object]$AllowedSet, [string]$EntName, [string]$Cname, [int]$ValInt, [object]$SourceOptions = $null)
    if ($script:OptionSetInteractiveChoice.ContainsKey($Key)) { return $script:OptionSetInteractiveChoice[$Key] }
    Write-Host ""
    Write-Host "  [Option set] Encja: $EntName, pole: $Cname - wartosc w zipie: $ValInt nie istnieje w celu." -ForegroundColor Yellow
    $sourceValues = $null
    if ($SourceOptions -and $SourceOptions.Options -and $SourceOptions.Options.Count -gt 0) {
        Write-Host "  Opcje w ZRODLE (numer=nazwa):" -ForegroundColor Cyan
        $sorted = $SourceOptions.Options | Sort-Object -Property Value
        foreach ($o in $sorted) { Write-Host "    $($o.Value)=$($o.Label)" -ForegroundColor Gray }
        $sourceValues = [System.Collections.Generic.HashSet[int]]::new()
        foreach ($o in $SourceOptions.Options) { [void]$sourceValues.Add([int]$o.Value) }
        $prompt = '  Na jaka opcje ze zrodla zamienic? Wpisz numer (Enter=pomin, C=wyczysc pole): '
    } else {
        Write-Host "  Dozwolone w CELU (numer=nazwa):" -ForegroundColor Cyan
        $sorted = $AttrInfo.Options | Sort-Object -Property Value
        foreach ($o in $sorted) { Write-Host "    $($o.Value)=$($o.Label)" -ForegroundColor Gray }
        Write-Host "  (Jesli w zrodle masz opcje Inne - w Config ustaw Zrodlo i Cel w Polaczenia.txt, zeby zobaczyc opcje ze zrodla.)" -ForegroundColor DarkGray
        $prompt = '  Wpisz numer do podstawienia (Enter=pomin, C=wyczysc pole): '
    }
    $userInput = (Read-Host $prompt).Trim()
    if ([string]::IsNullOrWhiteSpace($userInput)) { $script:OptionSetInteractiveChoice[$Key] = 'Skip'; return 'Skip' }
    if ($userInput -eq 'C' -or $userInput -eq 'c') { $script:OptionSetInteractiveChoice[$Key] = 'Clear'; return 'Clear' }
    $num = $null
    if ([int]::TryParse($userInput, [ref]$num)) {
        if ($null -ne $sourceValues -and $sourceValues.Contains($num)) {
            $script:OptionSetInteractiveChoice[$Key] = $num; return $num
        }
        if ($null -eq $sourceValues -and $AllowedSet -and $AllowedSet.Contains($num)) {
            $script:OptionSetInteractiveChoice[$Key] = $num; return $num
        }
    }
    Write-Host "  Nieprawidlowy numer - pomijam (zostaw bez zmiany)." -ForegroundColor DarkYellow
    $script:OptionSetInteractiveChoice[$Key] = 'Skip'
    return 'Skip'
}

function Invoke-OptionSetValidationForField {
    param($Ent, $Rec, $Child, $Cname, $ConnTarget)
    $changed = $false
    $entName = $Ent.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($entName)) { $entName = $Ent.GetAttribute('Name') }
    if ([string]::IsNullOrWhiteSpace($entName)) { return $false }
    $allowedMap = Get-TargetOptionSetAllowedForEntity -Conn $ConnTarget -EntityLogicalName $entName
    if (-not $allowedMap -or -not $allowedMap.ContainsKey($Cname)) { return $false }
    $attrInfo = $allowedMap[$Cname]
    $allowedSet = $attrInfo.AllowedSet
    $rawVal = $Child.GetAttribute('value'); if ([string]::IsNullOrWhiteSpace($rawVal)) { $rawVal = $Child.GetAttribute('Value') }; if ([string]::IsNullOrWhiteSpace($rawVal)) { $rawVal = $Child.InnerText }
    if ([string]::IsNullOrWhiteSpace($rawVal)) { return $false }
    $valInt = $null
    if (-not [int]::TryParse($rawVal.Trim(), [ref]$valInt)) { return $false }
    if (-not $allowedSet -or $allowedSet.Contains($valInt)) { return $false }
    $recId = $Rec.GetAttribute('id'); if ([string]::IsNullOrWhiteSpace($recId)) { $recId = $Rec.GetAttribute('Id') }
    $recName = ''
    $nameNode = $Rec.SelectSingleNode(".//*[@name='name' or @Name='name']")
    if ($nameNode) { $recName = $nameNode.InnerText.Trim(); if ([string]::IsNullOrWhiteSpace($recName)) { $recName = $nameNode.GetAttribute('value') } }
    $allowedValuesStr = ($allowedSet | ForEach-Object { $_ } | Sort-Object) -join ','
    [void]$script:OptionSetIssues.Add([PSCustomObject]@{ Entity = $entName; RecordId = $recId; RecordName = $recName; Field = $Cname; ValueInZip = $valInt; AllowedValues = $allowedValuesStr })
    $replacement = $null
    if ($script:OptionSetValidationAction -eq 'Interactive') {
        $key = $entName + '|' + $Cname + '|' + [string]$valInt
        $replacement = Get-OptionSetUserChoice -Key $key -AttrInfo $attrInfo -AllowedSet $allowedSet -EntName $entName -Cname $Cname -ValInt $valInt
    }
    if ($script:OptionSetValidationAction -eq 'Clear') { $replacement = 'Clear' }
    if ($script:OptionSetValidationAction -eq 'Replace' -and $script:OptionSetFallbackValues -and $script:OptionSetFallbackValues.ContainsKey($Cname)) { $replacement = $script:OptionSetFallbackValues[$Cname] }
    if ($replacement -eq 'Clear') {
        $Child.InnerText = ''; try { $Child.SetAttribute('value', '') } catch { }; try { $Child.SetAttribute('Value', '') } catch { }; $changed = $true
    } elseif ($replacement -is [int] -or ($replacement -ne 'Skip' -and $null -ne $replacement)) {
        $fallbackVal = [string]$replacement
        $Child.InnerText = $fallbackVal; try { $Child.SetAttribute('value', $fallbackVal) } catch { }; try { $Child.SetAttribute('Value', $fallbackVal) } catch { }; $changed = $true
    }
    return $changed
}

function Convert-CMTXmlContent {
    param([string]$Content, [System.IO.FileInfo]$File)
    $content = $Content
    $changed = $false
    try {
        Write-Host "    Laduje XML (DOM)..." -ForegroundColor DarkGray
        $doc = [xml]$content
        $entityNodes = @()
        foreach ($n in $doc.SelectNodes('//*[local-name()="entity" and (@name or @Name)]')) { $entityNodes += $n }
        foreach ($n in $doc.SelectNodes('//*[local-name()="Entity" and (@name or @Name)]')) { $entityNodes += $n }
        if ($entityNodes.Count -eq 0 -and $doc.DocumentElement) {
            foreach ($child in $doc.DocumentElement.ChildNodes) {
                if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                $ln = $child.LocalName
                if (($ln -eq 'entity' -or $ln -eq 'Entity') -and ($child.GetAttribute('name') -or $child.GetAttribute('Name'))) { $entityNodes += $child }
            }
        }
        Write-Host "    Encje: $($entityNodes.Count), przetwarzam rekordy..." -ForegroundColor DarkGray
        $stripLookupSet = $null
        if ($script:LookupFieldsToStripFromImport -and $script:LookupFieldsToStripFromImport.Count -gt 0) {
            $stripLookupSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($fn in $script:LookupFieldsToStripFromImport) { [void]$stripLookupSet.Add([string]$fn) }
        }
        foreach ($ent in $entityNodes) {
            $en = $ent.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($en)) { $en = $ent.GetAttribute('Name') }
            if (-not [string]::IsNullOrWhiteSpace($en)) { [void]$script:EntitiesFoundInZip.Add($en) }
            $recordNodes = @()
            foreach ($r in $ent.SelectNodes('.//*[local-name()="record"]')) { $recordNodes += $r }
            foreach ($r in $ent.SelectNodes('.//*[local-name()="Record"]')) { $recordNodes += $r }
            # Deduplikacja: CMT rzuca "Element o tym samym kluczu zostal juz dodany" przy duplikatach PK (account, contact, itd.)
            $pkName = $ent.GetAttribute('primaryidfield'); if ([string]::IsNullOrWhiteSpace($pkName)) { $pkName = $ent.GetAttribute('PrimaryIdField') }; if ([string]::IsNullOrWhiteSpace($pkName)) { $pkName = $en + 'id' }
            $seenPk = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            $duplicatesToRemove = [System.Collections.ArrayList]::new()
            foreach ($rec in $recordNodes) {
                $pkVal = $rec.GetAttribute('id'); if ([string]::IsNullOrWhiteSpace($pkVal)) { $pkVal = $rec.GetAttribute('Id') }
                if ([string]::IsNullOrWhiteSpace($pkVal)) {
                    foreach ($n in $rec.ChildNodes) {
                        if ($n.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                        $fn = $n.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($fn)) { $fn = $n.GetAttribute('Name') }
                        if ($fn -and ($fn -eq $pkName -or $fn -eq ($en + 'id'))) {
                            $pkVal = $n.InnerText.Trim(); if ([string]::IsNullOrWhiteSpace($pkVal)) { $pkVal = $n.GetAttribute('value'); if ([string]::IsNullOrWhiteSpace($pkVal)) { $pkVal = $n.GetAttribute('Value') } }
                            break
                        }
                    }
                }
                if ([string]::IsNullOrWhiteSpace($pkVal)) {
                    foreach ($n in $rec.SelectNodes('.//*[@name or @Name]')) {
                        $fn = $n.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($fn)) { $fn = $n.GetAttribute('Name') }
                        if ($fn -and ($fn -eq $pkName -or $fn -eq ($en + 'id'))) {
                            $pkVal = $n.InnerText.Trim(); if ([string]::IsNullOrWhiteSpace($pkVal)) { $pkVal = $n.GetAttribute('value'); if ([string]::IsNullOrWhiteSpace($pkVal)) { $pkVal = $n.GetAttribute('Value') } }
                            break
                        }
                    }
                }
                if (-not [string]::IsNullOrWhiteSpace($pkVal)) {
                    $pkNorm = $pkVal.TrimStart('{').TrimEnd('}').Trim()
                    if ($seenPk.Contains($pkNorm)) { [void]$duplicatesToRemove.Add($rec) } else { [void]$seenPk.Add($pkNorm) }
                }
            }
            foreach ($dup in $duplicatesToRemove) {
                if ($dup.ParentNode) { [void]$dup.ParentNode.RemoveChild($dup); $changed = $true; $script:DuplicateRecordsRemovedCount++ }
            }
            foreach ($rec in $recordNodes) {
                if ($duplicatesToRemove -contains $rec) { continue }
                # Zduplikowane overriddencreatedon (bug eksportu) - zostaw tylko pierwsze
                $overrideDupNodes = @($rec.SelectNodes('.//*[@name or @Name]') | Where-Object {
                    $fn = $_.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($fn)) { $fn = $_.GetAttribute('Name') }
                    $fn -eq 'overriddencreatedon'
                })
                for ($i = 1; $i -lt $overrideDupNodes.Count; $i++) {
                    if ($overrideDupNodes[$i].ParentNode) { [void]$overrideDupNodes[$i].ParentNode.RemoveChild($overrideDupNodes[$i]); $changed = $true; $script:DuplicateOverrideRemovedCount++ }
                }
                # Lookupy do encji nieobecnych w pakiecie (msdyn_contactkpiid, msdyn_accountkpiid itd.) - usun z rekordu
                if ($stripLookupSet -and $stripLookupSet.Count -gt 0) {
                    $toStrip = @($rec.SelectNodes('.//*[@name or @Name]') | Where-Object {
                        $fn = $_.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($fn)) { $fn = $_.GetAttribute('Name') }; if ([string]::IsNullOrWhiteSpace($fn)) { $fn = $_.LocalName }
                        -not [string]::IsNullOrWhiteSpace($fn) -and $stripLookupSet.Contains($fn)
                    })
                    foreach ($n in $toStrip) { if ($n.ParentNode) { [void]$n.ParentNode.RemoveChild($n); $changed = $true; $script:LookupFieldsStrippedCount++ } }
                }
                $createdonVal = $null
                $modifiedonVal = $null
                $overrideNode = $null
                $fieldLikeParent = $null
                $useValueAttr = $false
                $allFieldNodes = @($rec.SelectNodes('.//*[@name or @Name]'))
                if ($allFieldNodes.Count -eq 0) { $allFieldNodes = @($rec.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element }) }
                foreach ($child in $allFieldNodes) {
                    if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                    $cname = $child.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($cname)) { $cname = $child.GetAttribute('Name') }
                    if ([string]::IsNullOrWhiteSpace($cname)) { $cname = $child.LocalName }
                    if ($null -eq $fieldLikeParent -and $child.LocalName -match '^(field|attribute|Field|Attribute)$') { $fieldLikeParent = $child }
                    if ($cname -eq 'createdon') {
                        $createdonVal = $child.InnerText.Trim()
                        if ([string]::IsNullOrWhiteSpace($createdonVal)) {
                            $createdonVal = $child.GetAttribute('value')
                            if ([string]::IsNullOrWhiteSpace($createdonVal)) { $createdonVal = $child.GetAttribute('Value') }
                            if (-not [string]::IsNullOrWhiteSpace($createdonVal)) { $useValueAttr = $true }
                        }
                    }
                    if ($cname -eq 'modifiedon') {
                        $modifiedonVal = $child.InnerText.Trim()
                        if ([string]::IsNullOrWhiteSpace($modifiedonVal)) { $modifiedonVal = $child.GetAttribute('value'); if ([string]::IsNullOrWhiteSpace($modifiedonVal)) { $modifiedonVal = $child.GetAttribute('Value') } }
                    }
                    if ($cname -eq 'overriddencreatedon') { $overrideNode = $child }
                    if ($cname -eq 'ownerid' -or $cname -eq 'owner') {
                        $ownerVal = $child.InnerText.Trim()
                        if ([string]::IsNullOrWhiteSpace($ownerVal)) { $ownerVal = $child.GetAttribute('value'); if ([string]::IsNullOrWhiteSpace($ownerVal)) { $ownerVal = $child.GetAttribute('Value') } }
                        if (-not [string]::IsNullOrWhiteSpace($ownerVal)) {
                            $ownerGuid = $ownerVal -replace '^\{|\}$', ''
                            $ownerKey = $ownerGuid.Trim().ToLowerInvariant()
                            $tgtOwner = $null
                            if ($ownerGuid -match '^[0-9a-fA-F\-]{36}$' -and $guidMap.ContainsKey($ownerKey)) {
                                $tgtOwner = $guidMap[$ownerKey]
                            } elseif ($displayNameToGuid.Count -gt 0) {
                                $dnKeyNorm = [regex]::Replace($ownerVal.Trim(), '\s+', ' ').ToLowerInvariant()
                                if ($displayNameToGuid.ContainsKey($dnKeyNorm)) { $tgtOwner = $displayNameToGuid[$dnKeyNorm] }
                                if (-not $tgtOwner) {
                                    $dnKeyNorm = $ownerVal.Trim().ToLowerInvariant()
                                    if ($displayNameToGuid.ContainsKey($dnKeyNorm)) { $tgtOwner = $displayNameToGuid[$dnKeyNorm] }
                                }
                                if (-not $tgtOwner -and $dnKeyNorm -match '^(.+)\s+(.+)$') {
                                    $reversed = ($Matches[2] + ' ' + $Matches[1]).Trim()
                                    if ($displayNameToGuid.ContainsKey($reversed)) { $tgtOwner = $displayNameToGuid[$reversed] }
                                }
                            }
                            if (-not [string]::IsNullOrWhiteSpace($tgtOwner)) {
                                $child.InnerText = $tgtOwner
                                try { $child.SetAttribute('value', $tgtOwner) } catch { }
                                try { $child.SetAttribute('Value', $tgtOwner) } catch { }
                                $changed = $true
                                $script:ownerReplaceCount++
                            }
                        }
                    }
                    if ($script:doOptionSetValidation -and ($script:connTarget -or $script:MetadataCacheLoaded) -and -not [string]::IsNullOrWhiteSpace($cname)) {
                        if (Invoke-OptionSetValidationForField -Ent $ent -Rec $rec -Child $child -Cname $cname -ConnTarget $script:connTarget) { $changed = $true }
                    }
                }
                # Dla annotation: jesli w zipie nie ma createdon, uzyj modifiedon jako daty (oryginalna data utworzenia)
                if ([string]::IsNullOrWhiteSpace($createdonVal) -and -not [string]::IsNullOrWhiteSpace($modifiedonVal) -and [string]::Equals($en, 'annotation', [StringComparison]::OrdinalIgnoreCase)) {
                    $createdonVal = $modifiedonVal
                }
                # Overriddencreatedon tylko tam gdzie cel ma to pole (np. salesliteratureitem w niektorych orgach go nie ma - CMT: Missing Fields)
                $enKey = $en.Trim().ToLowerInvariant()
                $entityHasOverrideInTarget = $false
                if ($script:targetAllowedAttrs -and $script:targetAllowedAttrs.ContainsKey($enKey)) {
                    $entityHasOverrideInTarget = $script:targetAllowedAttrs[$enKey].Contains('overriddencreatedon')
                }
                if ($entityHasOverrideInTarget -and -not [string]::IsNullOrWhiteSpace($createdonVal)) {
                    $overrideDateIso = ConvertTo-DateTimeIso $createdonVal
                    if ($overrideNode) {
                        $overrideNode.InnerText = $overrideDateIso
                        if ($useValueAttr) { $overrideNode.SetAttribute('value', $overrideDateIso) }
                    } else {
                        if ($fieldLikeParent -and $fieldLikeParent.LocalName -match '^(field|Field)$') {
                            $ns = $fieldLikeParent.NamespaceURI
                            if (-not [string]::IsNullOrWhiteSpace($ns)) {
                                $newEl = $doc.CreateElement($fieldLikeParent.LocalName, $ns)
                            } else {
                                $newEl = $doc.CreateElement($fieldLikeParent.LocalName)
                            }
                            $newEl.SetAttribute('name', 'overriddencreatedon')
                            $newEl.InnerText = $overrideDateIso
                            if ($useValueAttr) { $newEl.SetAttribute('value', $overrideDateIso) }
                            $parent = if ($fieldLikeParent.ParentNode -and $fieldLikeParent.ParentNode -ne $rec) { $fieldLikeParent.ParentNode } else { $rec }
                            [void]$parent.AppendChild($newEl)
                        } else {
                            $newEl = $doc.CreateElement('overriddencreatedon')
                            $newEl.InnerText = $overrideDateIso
                            [void]$rec.AppendChild($newEl)
                        }
                    }
                    $changed = $true
                    $script:overrideCount++
                } elseif (-not $entityHasOverrideInTarget -and $overrideNode -and $overrideNode.ParentNode) {
                    [void]$overrideNode.ParentNode.RemoveChild($overrideNode)
                    $changed = $true
                }
            }
        }
        $doStrip = $script:StripFieldsNotInTarget -and ($script:connTarget -or ($script:targetAllowedAttrs -and $script:targetAllowedAttrs.Count -gt 0))
        if ($doStrip) {
            foreach ($ent in $entityNodes) {
                $entName = $ent.GetAttribute('name')
                if ([string]::IsNullOrWhiteSpace($entName)) { continue }
                if (-not $script:targetAllowedAttrs.ContainsKey($entName)) {
                    if (-not $script:connTarget) { continue }
                    try {
                        if ($script:entityFiltersAttributes -ne $null) {
                            $meta = Get-CrmEntityMetadata -Conn $script:connTarget -EntityLogicalName $entName -EntityFilters $script:entityFiltersAttributes -ErrorAction Stop
                        } else {
                            $meta = Get-CrmEntityMetadata -Conn $script:connTarget -EntityLogicalName $entName -ErrorAction Stop
                        }
                        if ($meta -and $meta.Attributes) {
                            $allowed = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                            foreach ($a in $meta.Attributes) { [void]$allowed.Add($a.LogicalName) }
                            $script:targetAllowedAttrs[$entName] = $allowed
                            Write-Host "  Cel: encja $entName - atrybutow dozwolonych: $($allowed.Count)" -ForegroundColor Gray
                        } else {
                            Write-Host "  Cel: encja $entName - brak atrybutow w metadanych (pomijam strip)." -ForegroundColor Yellow
                        }
                    } catch {
                        Write-Host "  Cel: encja $entName - blad metadanych: $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }
                $allowedSet = $script:targetAllowedAttrs[$entName]
                if (-not $allowedSet) { continue }
                $recordNodes = $ent.SelectNodes('.//*[local-name()="record"]') + $ent.SelectNodes('.//*[local-name()="Record"]')
                foreach ($rec in $recordNodes) {
                    $toRemove = @()
                    foreach ($child in $rec.ChildNodes) {
                        if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                        $name = $child.GetAttribute('name')
                        if ([string]::IsNullOrWhiteSpace($name)) { $name = $child.GetAttribute('Name') }
                        if ([string]::IsNullOrWhiteSpace($name)) { $name = $child.LocalName }
                        if ([string]::IsNullOrWhiteSpace($name)) { continue }
                        if ($allowedSet.Contains($name) -eq $false) { $toRemove += $child }
                    }
                    foreach ($n in $toRemove) { [void]$rec.RemoveChild($n); $changed = $true }
                }
                $fieldsParent = $ent.SelectSingleNode('.//*[local-name()="fields"]')
                if (-not $fieldsParent) { $fieldsParent = $ent.SelectSingleNode('.//*[local-name()="Fields"]') }
                if ($fieldsParent) {
                    $toRemoveFields = @()
                    $hasOverride = $false
                    foreach ($fieldNode in $fieldsParent.ChildNodes) {
                        if ($fieldNode.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                        $fname = $fieldNode.GetAttribute('name')
                        if ([string]::IsNullOrWhiteSpace($fname)) { $fname = $fieldNode.GetAttribute('Name') }
                        if ([string]::IsNullOrWhiteSpace($fname)) { $fname = $fieldNode.LocalName }
                        if ($fname -eq 'overriddencreatedon') { $hasOverride = $true }
                        if (-not [string]::IsNullOrWhiteSpace($fname) -and $allowedSet.Contains($fname) -eq $false) { $toRemoveFields += $fieldNode }
                    }
                    foreach ($n in $toRemoveFields) { [void]$fieldsParent.RemoveChild($n); $changed = $true }
                    if (-not $hasOverride -and $allowedSet.Contains('overriddencreatedon')) {
                        $newField = $doc.CreateElement('field')
                        $newField.SetAttribute('displayname', 'overriddencreatedon')
                        $newField.SetAttribute('name', 'overriddencreatedon')
                        $newField.SetAttribute('type', 'datetime')
                        [void]$fieldsParent.AppendChild($newField)
                        $changed = $true
                    }
                }
            }
        }
        # Pomijaj encje bez rekordow tylko w pliku danych (NIE w data_schema.xml - tam sa definicje, nie rekordy)
        # Oraz encje z listy EntitiesToExcludeFromImport (np. salesliteratureitem gdy cel ma inna wersje)
        $isDataFile = $File -and [string]::Equals($File.Name, 'data.xml', [StringComparison]::OrdinalIgnoreCase)
        if ($isDataFile) {
            $entitiesToRemove = [System.Collections.ArrayList]::new()
            foreach ($ent in $entityNodes) {
                $en = $ent.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($en)) { $en = $ent.GetAttribute('Name') }
                $recs = @($ent.SelectNodes('.//*[local-name()="record"]')) + @($ent.SelectNodes('.//*[local-name()="Record"]'))
                $remove = ($recs.Count -eq 0) -or ($script:EntitiesToExcludeFromImport -and $script:EntitiesToExcludeFromImport.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($en) -and $script:EntitiesToExcludeFromImport.Contains($en))
                if ($remove) { [void]$entitiesToRemove.Add([PSCustomObject]@{ Node = $ent; EntityName = $en }) }
            }
            foreach ($item in $entitiesToRemove) {
                if ($item.Node.ParentNode) {
                    [void]$item.Node.ParentNode.RemoveChild($item.Node)
                    if (-not [string]::IsNullOrWhiteSpace($item.EntityName)) { [void]$script:EntitiesFoundInZip.Remove($item.EntityName) }
                    $changed = $true
                }
            }
            if ($entitiesToRemove.Count -gt 0) {
                $removedNames = @($entitiesToRemove | ForEach-Object { $_.EntityName } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                Write-Host "    Pominieto encje bez rekordow / z listy wykluczen: $($removedNames -join ', ')." -ForegroundColor DarkGray
            }
        }
        # W data_schema.xml usun definicje encji z listy wykluczen (zeby CMT nie walidowal Missing Fields)
        $isSchemaFile = $File -and [string]::Equals($File.Name, 'data_schema.xml', [StringComparison]::OrdinalIgnoreCase)
        if ($isSchemaFile -and $script:EntitiesToExcludeFromImport -and $script:EntitiesToExcludeFromImport.Count -gt 0) {
            $schemaEntitiesToRemove = [System.Collections.ArrayList]::new()
            foreach ($ent in $entityNodes) {
                $en = $ent.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($en)) { $en = $ent.GetAttribute('Name') }
                if (-not [string]::IsNullOrWhiteSpace($en) -and $script:EntitiesToExcludeFromImport.Contains($en)) {
                    [void]$schemaEntitiesToRemove.Add([PSCustomObject]@{ Node = $ent; EntityName = $en })
                }
            }
            foreach ($item in $schemaEntitiesToRemove) {
                if ($item.Node.ParentNode) {
                    [void]$item.Node.ParentNode.RemoveChild($item.Node)
                    $changed = $true
                }
            }
            if ($schemaEntitiesToRemove.Count -gt 0) {
                $removedSchemaNames = @($schemaEntitiesToRemove | ForEach-Object { $_.EntityName })
                Write-Host "    Z schematu usunieto encje z listy wykluczen: $($removedSchemaNames -join ', ')." -ForegroundColor DarkGray
            }
        }
        if ($changed) { $content = $doc.OuterXml }
        if ($displayNameToGuid.Count -gt 0) {
            $contentBeforeDn = $content
            foreach ($dn in $displayNameToGuid.Keys) {
                $tgt = $displayNameToGuid[$dn]
                $esc = [regex]::Escape($dn)
                $content = $content -creplace ('(?i)(<[^>]*name=["'']ownerid["''][^>]*>)' + $esc + '(</[^>]+>)'), ('$1' + $tgt + '$2')
                $content = $content -creplace ('(?i)(<[^>]*name=["'']ownerid["''][^>]*value=["''])' + $esc + '(["''])', ('$1' + $tgt + '$2'))
            }
            if ($content -ne $contentBeforeDn) { $changed = $true }
        }
    } catch {
        Write-Host "  Ostrzezenie: nie udalo sie przetworzyc XML $($File.Name): $($_.Exception.Message)" -ForegroundColor Yellow
    }
    return @{ Content = $content; Changed = $changed }
}

function Invoke-TransformCMTFiles {
    param([System.IO.FileInfo[]]$Files)
    $total = $Files.Count
    $current = 0
    foreach ($f in $Files) {
        $current++
        if ($total -gt 3 -and ($current -eq 1 -or $current -eq $total -or ($current % 10 -eq 0))) {
            Write-Host "  Przetwarzam pliki... $current / $total" -ForegroundColor DarkGray
        }
        $content = $null
        try {
            $content = [System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)
        } catch { continue }
        $changed = $false
        foreach ($src in $guidMap.Keys) {
            $tgt = $guidMap[$src]
            if ($content.IndexOf($src, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $content = $content -ireplace [regex]::Escape($src), $tgt
                $changed = $true
            }
            $srcBraced = '{' + $src + '}'
            if ($content.IndexOf($srcBraced, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $content = $content -ireplace ('\{' + [regex]::Escape($src) + '\}'), ('{' + $tgt + '}')
                $changed = $true
            }
        }
        if ($f.Extension -eq '.xml') {
            $sizeMb = [Math]::Round($content.Length / 1MB, 1).ToString([System.Globalization.CultureInfo]::InvariantCulture)
            Write-Host "  Plik $current/$total : $($f.Name) ($sizeMb MB) - parsowanie i transformacja (moze potrwac)..." -ForegroundColor Gray
            $res = Convert-CMTXmlContent -Content $content -File $f
            $content = $res.Content
            $changed = $res.Changed
        }
        if ($changed) {
            [System.IO.File]::WriteAllText($f.FullName, $content, [System.Text.UTF8Encoding]::new($false))
        }
    }
}

$tempDir = Join-Path $env:TEMP ([System.Guid]::NewGuid().ToString('N'))
try {
    [System.IO.Compression.ZipFile]::ExtractToDirectory($InputZipPath, $tempDir)
    $files = Get-ChildItem -Path $tempDir -Recurse -File
    if ($guidMap.Count -gt 0) {
        $foundGuids = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($f in $files) {
            $text = [System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)
            foreach ($src in $guidMap.Keys) {
                if (-not $foundGuids.Contains($src) -and $text.IndexOf($src, [StringComparison]::OrdinalIgnoreCase) -ge 0) { [void]$foundGuids.Add($src) }
            }
            if ($foundGuids.Count -eq $guidMap.Count) { break }
        }
        $found = $foundGuids.Count
        if ($found -eq 0) {
            Write-Host 'W zipie nie ma GUIDow z IdMap (ownerid moze byc jako imie i nazwisko). Uzycie mapowania po display name (CMT_IdMap_ByDisplayName.json).' -ForegroundColor Gray
        } else {
            Write-Host ('W zipie znaleziono ' + $found + ' z ' + $guidMap.Count + ' GUIDow z IdMap - podmiana ownerow bedzie wykonana.') -ForegroundColor Gray
        }
    }
    # Jedna transformacja; przy Interactive pytanie o option set przy pierwszym wystapieniu kazdego problemu (bez wstepnego skanowania)
    if ($doOptionSetValidation -and ($script:connTarget -or $script:MetadataCacheLoaded)) {
        Write-Host "Walidacja option setow: wlaczona (akcja: $($script:OptionSetValidationAction)). Pytania w trakcie transformacji przy pierwszym wystapieniu." -ForegroundColor Gray
    }
    [void]$script:EntitiesFoundInZip.Clear()
    Invoke-TransformCMTFiles -Files $files
    $entitiesList = @($script:EntitiesFoundInZip | Sort-Object)
    if ($entitiesList.Count -gt 0) {
        Write-Host ('W zipie przetworzono encje: ' + ($entitiesList -join ', ') + '.') -ForegroundColor Gray
        if ($script:MetadataCacheLoaded -and $script:targetAllowedAttrs -and $script:targetAllowedAttrs.Count -gt 0) {
            $missingInCache = [System.Collections.ArrayList]::new()
            foreach ($e in $entitiesList) {
                $ek = $e.Trim().ToLowerInvariant()
                if (-not $script:targetAllowedAttrs.ContainsKey($ek)) { [void]$missingInCache.Add($e) }
            }
            if ($missingInCache.Count -eq 0) {
                Write-Host "Cache: wszystkie encje z zaznaczonej paczki sa w cache." -ForegroundColor Green
            } else {
                Write-Host ("Cache: brakuje " + $missingInCache.Count + " encji: " + ($missingInCache -join ', ')) -ForegroundColor Yellow
                Write-Host "  Uruchom opcje 5 (Pobierz metadane celu) z tym zipem, zeby uzupelnic cache." -ForegroundColor Gray
            }
        }
        $lastZipEntitiesPath = Join-Path $outputDirForCache 'CMT_LastZipEntities.json'
        try {
            $lastRunPayload = @{ ZipPath = $InputZipPath; ZipName = [System.IO.Path]::GetFileName($InputZipPath); Entities = @($entitiesList); CheckedAt = (Get-Date).ToString('o') }
            $lastRunPayload | ConvertTo-Json -Compress:$false | Set-Content -Path $lastZipEntitiesPath -Encoding UTF8 -ErrorAction Stop
        } catch { }
        Write-Host 'Uwaga: W tym zipie sa tylko te encje. Jesli przy imporcie bedzie brakowac rekordow lub calych encji - dodaj do folderu Input zip z CMT zawierajacy brakujace encje i uruchom opcje 3 dla tego zipa.' -ForegroundColor Yellow
    }
    if ($doOptionSetValidation -and $script:OptionSetIssues -and $script:OptionSetIssues.Count -gt 0) {
        $reportDir = Split-Path $OutputZipPath -Parent
        if ([string]::IsNullOrWhiteSpace($reportDir)) { $reportDir = (Get-Location).Path }
        $reportName = "CMT_OptionSetValidation_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv"
        $reportPath = Join-Path $reportDir $reportName
        $script:OptionSetIssues | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
        Write-Host ('Walidacja option setow: znaleziono ' + $script:OptionSetIssues.Count + ' niepasujacych wartosci. Raport: ' + $reportPath) -ForegroundColor Yellow
        if ($script:OptionSetValidationAction -eq 'Report') {
            Write-Host '  Aby wybrac zamienniki, ustaw w Config OptionSetValidationAction=Interactive i uruchom transformacje ponownie.' -ForegroundColor Gray
        }
    }

    if (Test-Path $OutputZipPath -PathType Leaf) { Remove-Item $OutputZipPath -Force }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $OutputZipPath)
    $elapsed = (Get-Date) - $script:TransformStartTime
    $minStr = [Math]::Round($elapsed.TotalMinutes, 1).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    Write-Host "Zakonczono: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss') (trwalo: $minStr min)" -ForegroundColor Gray
    Write-Host ('Zapisano: ' + $OutputZipPath)
    if ($script:overrideCount -gt 0) { Write-Host ('Ustawiono overriddencreatedon w ' + $script:overrideCount + ' rekordach (oryginalna data utworzenia).') -ForegroundColor Gray }
    if ($script:DuplicateOverrideRemovedCount -gt 0) { Write-Host ('Usunieto ' + $script:DuplicateOverrideRemovedCount + ' zduplikowanych pol overriddencreatedon w rekordach (bug eksportu).') -ForegroundColor Gray }
    if ($script:LookupFieldsStrippedCount -gt 0) { Write-Host ('Usunieto ' + $script:LookupFieldsStrippedCount + ' pol lookup (msdyn_contactkpiid, msdyn_accountkpiid, transactioncurrencyid, originatingleadid) - brak w celu.') -ForegroundColor Gray }
    if ($script:DuplicateRecordsRemovedCount -gt 0) { Write-Host ('Usunieto ' + $script:DuplicateRecordsRemovedCount + ' zduplikowanych rekordow (ten sam klucz glowny) - zapobiega bledowi CMT: Element o tym samym kluczu.') -ForegroundColor Gray }
    if ($guidMap.Count -gt 0 -or $displayNameToGuid.Count -gt 0) {
        Write-Host 'Podmiana ownerow: ' -NoNewline -ForegroundColor Gray
        if ($guidMap.Count -gt 0) { Write-Host ('IdMap GUID ' + $guidMap.Count + '; ') -NoNewline -ForegroundColor Gray }
        if ($displayNameToGuid.Count -gt 0) { Write-Host ('imie i nazwisko -> GUID: ' + $displayNameToGuid.Count + '; ') -NoNewline -ForegroundColor Gray }
        Write-Host ('w rekordach ustawiono ownerid: ' + $script:ownerReplaceCount + ' razy.') -ForegroundColor Gray
    }

    # Weryfikacja: czy w wyjsciowym zipie faktycznie sa zmiany i czy ownerid to GUID celu
    $verifyDir = Join-Path $env:TEMP ([System.Guid]::NewGuid().ToString('N'))
    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($OutputZipPath, $verifyDir)
        $dataFiles = @(Get-ChildItem -Path $verifyDir -Recurse -Filter 'data.xml' -File -ErrorAction SilentlyContinue)
        if ($dataFiles.Count -eq 0) { $dataFiles = @(Get-ChildItem -Path $verifyDir -Recurse -Filter '*.xml' -File -ErrorAction SilentlyContinue) }
        $totalOverrideInZip = 0
        $totalOwnerInZip = 0
        $sampleOwnerLine = $null
        $anyRecordCount = 0
        $targetGuidSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($v in $guidMap.Values) { [void]$targetGuidSet.Add([string]$v) }
        foreach ($v in $displayNameToGuid.Values) { [void]$targetGuidSet.Add([string]$v) }
        $ownerValuesInZip = [System.Collections.ArrayList]::new()
        $ownerValuesMapped = 0
        $ownerValuesNotMapped = 0
        foreach ($df in $dataFiles) {
            $xmlText = [System.IO.File]::ReadAllText($df.FullName, [System.Text.Encoding]::UTF8)
            $totalOverrideInZip += ([regex]::Matches($xmlText, 'overriddencreatedon', [System.StringComparison]::OrdinalIgnoreCase)).Count
            $totalOwnerInZip += ([regex]::Matches($xmlText, 'ownerid', [System.StringComparison]::OrdinalIgnoreCase)).Count
            if ($totalOwnerInZip -gt 0 -and $null -eq $sampleOwnerLine) {
                if ($xmlText -match '(?s)<[^>]*ownerid[^>]*>([^<]*)</[^>]+>') { $sampleOwnerLine = $Matches[0].Substring(0, [Math]::Min(120, $Matches[0].Length)) + '...' }
            }
            $sq2 = [char]39; $dq2 = [char]34; $qC2 = '[' + $dq2 + $sq2 + ']'; $ownerIdContentPattern = '<[^>]*\bname=' + $qC2 + 'ownerid' + $qC2 + '[^>]*>([^<]*)<'
            foreach ($m in [regex]::Matches($xmlText, $ownerIdContentPattern, [System.StringComparison]::OrdinalIgnoreCase)) {
                $val = $m.Groups[1].Value.Trim()
                if (-not [string]::IsNullOrWhiteSpace($val)) {
                    [void]$ownerValuesInZip.Add($val)
                    $guidNorm = ($val.TrimStart('{').TrimEnd('}').Trim())
                    if ($targetGuidSet.Contains($guidNorm)) { $ownerValuesMapped++ } else { $ownerValuesNotMapped++ }
                }
            }
            $sq = [char]39; $dq = [char]34; $cc = '[' + [char]94 + $dq + $sq + ']'; $cg = '(' + $cc + '+)'; $qClass = '[' + $dq + $sq + ']'; $ownerIdValuePattern = 'name=' + $qClass + 'ownerid' + $qClass + '[^>]*value=' + $qClass + $cg + $qClass
            foreach ($m in [regex]::Matches($xmlText, $ownerIdValuePattern, [System.StringComparison]::OrdinalIgnoreCase)) {
                $val = $m.Groups[1].Value.Trim()
                if (-not [string]::IsNullOrWhiteSpace($val)) {
                    [void]$ownerValuesInZip.Add($val)
                    $guidNorm = ($val.TrimStart('{').TrimEnd('}').Trim())
                    if ($targetGuidSet.Contains($guidNorm)) { $ownerValuesMapped++ } else { $ownerValuesNotMapped++ }
                }
            }
            $anyRecordCount += ([regex]::Matches($xmlText, '<record', [System.StringComparison]::OrdinalIgnoreCase)).Count
            $anyRecordCount += ([regex]::Matches($xmlText, '<Record', [System.StringComparison]::OrdinalIgnoreCase)).Count
        }
        Remove-Item -Path $verifyDir -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "" -ForegroundColor Gray
        Write-Host '--- Weryfikacja zipa ---' -ForegroundColor Cyan
        Write-Host ('  Plikow data: ' + $dataFiles.Count + ' | Rekordow (record): ' + $anyRecordCount + ' | overriddencreatedon: ' + $totalOverrideInZip + ' | ownerid: ' + $totalOwnerInZip) -ForegroundColor White
        if ($ownerValuesInZip.Count -gt 0 -and $targetGuidSet.Count -gt 0) {
            Write-Host ('  ownerid w zipie: ' + $ownerValuesMapped + ' z ' + $ownerValuesInZip.Count + ' to GUID celu (zmapowane OK), ' + $ownerValuesNotMapped + ' nie zmapowane.') -ForegroundColor $(if ($ownerValuesNotMapped -eq 0) { 'Green' } else { 'Yellow' })
            if ($ownerValuesNotMapped -gt 0) {
                $examples = @($ownerValuesInZip | Where-Object {
                    $g = ($_.TrimStart('{').TrimEnd('}').Trim())
                    -not $targetGuidSet.Contains($g)
                } | Select-Object -First 5 -Unique)
                $joinPipe = ' | '; Write-Host ('  Przyklady wartosci ownerid nie zmapowanych na GUID celu: ' + ($examples -join $joinPipe)) -ForegroundColor Yellow
            }
        }
        if ($dataFiles.Count -eq 0) {
            Write-Host '  UWAGA: W zipie nie znaleziono plikow data.xml / *.xml - sprawdz zawartosc zipa.' -ForegroundColor Yellow
        }
        if ($anyRecordCount -gt 0 -and $totalOverrideInZip -eq 0) {
            Write-Host '  UWAGA: W zipie sa rekordy, ale brak overriddencreatedon - sprawdz strukture XML (np. czy pole createdon istnieje).' -ForegroundColor Yellow
        }
        if ($guidMap.Count -gt 0 -and $displayNameToGuid.Count -eq 0 -and $totalOwnerInZip -gt 0 -and $ownerValuesNotMapped -gt 0) {
            Write-Host '  Wlasciciele w zipie to prawdopodobnie imiona i nazwiska (nie GUID). Uruchom opcje 2 - generuje CMT_IdMap_ByDisplayName.json.' -ForegroundColor Yellow
        }
        if ($totalOverrideInZip -gt 0 -or $totalOwnerInZip -gt 0) {
            if ($ownerValuesNotMapped -eq 0 -and $ownerValuesInZip.Count -gt 0) { Write-Host '  OK: Wszystkie ownerid to GUID celu.' -ForegroundColor Green }
            elseif ($ownerValuesInZip.Count -eq 0) { Write-Host '  OK: Zmiany widoczne w pliku.' -ForegroundColor Green }
        }
        if ($sampleOwnerLine) { Write-Host ('  Przyklad owner w XML: ' + $sampleOwnerLine) -ForegroundColor DarkGray }
    } catch {
        Write-Host ('  Weryfikacja nie powiodla sie: ' + $_.Exception.Message) -ForegroundColor Yellow
        if (Test-Path $verifyDir) { Remove-Item -Path $verifyDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
    if ($script:EntitiesFoundInZip -and $script:EntitiesFoundInZip.Count -gt 0) {
        Write-Host ""
        Write-Host 'Jesli import nie powiodl sie dla jakiejs encji (np. account): sprawdz czy ta encja jest na liscie W zipie przetworzono encje powyzej. Jesli nie ma jej na liscie - ten zip jej nie zawiera; uruchom opcje 3 dla zipa z CMT, ktory zawiera te encje.' -ForegroundColor Gray
    }
} finally {
    if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
}

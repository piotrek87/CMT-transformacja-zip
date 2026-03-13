# Przetwarza zip CMT: podmiana ownerow (IdMap), zachowanie dat, usuniecie pol nieistniejacych w celu.
# Uzycie: .\Transform-CMTZip.ps1 -InputZipPath C:\...\CMT_Export.zip [-OutputZipPath ...] [-IdMapPath ...] [-TargetConnectionString ...] [-StripFieldsNotInTarget]
# Albo: -ConfigPath ...\CMTConfig.ps1 (wtedy cel i StripFieldsNotInTarget z configu – zalecane przy uruchomieniu z menu).

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
        if ($cfg -and $null -ne $cfg.OptionSetValidationAction) {
            $script:OptionSetValidationAction = [string]$cfg.OptionSetValidationAction
        }
        if ($cfg -and $cfg.OptionSetFallbackValues -and $cfg.OptionSetFallbackValues -is [hashtable]) {
            $script:OptionSetFallbackValues = $cfg.OptionSetFallbackValues
        }
        if ($cfg -and $cfg.LookupFieldsToStripFromImport -and $cfg.LookupFieldsToStripFromImport -is [array]) {
            $script:LookupFieldsToStripFromImport = @($cfg.LookupFieldsToStripFromImport)
        }
    } catch { Write-Host "Nie udalo sie zaladowac configu: $($_.Message)" -ForegroundColor Yellow }
}
if (-not $script:LookupFieldsToStripFromImport -or $script:LookupFieldsToStripFromImport.Count -eq 0) {
    $script:LookupFieldsToStripFromImport = @('msdyn_accountkpiid', 'msdyn_contactkpiid', 'transactioncurrencyid', 'originatingleadid')
}
$script:StripFieldsNotInTarget = $StripFieldsNotInTarget

# IdMap: podmiana GUID ownerow (zrodlo -> cel). Klucze musza byc czystym GUID (w zipie jest value="guid").
# Jesli JSON ma klucze w formacie CRM (np. "guid systemuserid=xxx"), wyciagamy czysty GUID.
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
if (($StripFieldsNotInTarget -or $script:StripFieldsNotInTarget) -and -not [string]::IsNullOrWhiteSpace($TargetConnectionString)) {
    if (Get-Module -ListAvailable -Name 'Microsoft.Xrm.Data.PowerShell') {
        Import-Module Microsoft.Xrm.Data.PowerShell -Force -ErrorAction SilentlyContinue
        try {
            $script:connTarget = Get-CrmConnection -ConnectionString $TargetConnectionString -ErrorAction Stop
            Write-Host "Polaczenie z celem OK - beda usuniete pola nieistniejace w celu."
        } catch {
            Write-Host "Brak polaczenia z celem - pomijam usuwanie pol." -ForegroundColor Yellow
            Write-Host "  Blad: $($_.Exception.Message)" -ForegroundColor Gray
            if ($_.Exception.InnerException) { Write-Host "  Inner: $($_.Exception.InnerException.Message)" -ForegroundColor Gray }
        }
    } else {
        Write-Host "StripFieldsNotInTarget wymaga Microsoft.Xrm.Data.PowerShell - pomijam." -ForegroundColor Yellow
    }
}
$script:entityFiltersAttributes = $null
try { $script:entityFiltersAttributes = [Microsoft.Xrm.Sdk.Metadata.EntityFilters]::Attributes } catch { }

# Polaczenie do celu tez dla walidacji option setow (Report/Clear/Replace)
$doOptionSetValidation = $script:OptionSetValidationAction -match '^(Report|Clear|Replace)$'
$script:doOptionSetValidation = $doOptionSetValidation
if ($doOptionSetValidation -and $null -eq $script:connTarget -and -not [string]::IsNullOrWhiteSpace($TargetConnectionString)) {
    if (Get-Module -ListAvailable -Name 'Microsoft.Xrm.Data.PowerShell') {
        Import-Module Microsoft.Xrm.Data.PowerShell -Force -ErrorAction SilentlyContinue
        try {
            $script:connTarget = Get-CrmConnection -ConnectionString $TargetConnectionString -ErrorAction Stop
            Write-Host "Polaczenie z celem OK - walidacja option setow (akcja: $($script:OptionSetValidationAction))." -ForegroundColor Gray
        } catch {
            Write-Host "Brak polaczenia z celem - pomijam walidacje option setow. $($_.Message)" -ForegroundColor Yellow
            $doOptionSetValidation = $false
        }
    } else {
        Write-Host "Walidacja option setow wymaga Microsoft.Xrm.Data.PowerShell - pomijam." -ForegroundColor Yellow
        $doOptionSetValidation = $false
    }
}

$script:overrideCount = 0
$script:ownerReplaceCount = 0
$script:DuplicateRecordsRemovedCount = 0
$script:LookupFieldsStrippedCount = 0
$script:DuplicateOverrideRemovedCount = 0
$script:OptionSetIssues = [System.Collections.ArrayList]::new()
$script:TargetOptionSetAllowed = @{}
$script:OptionSetInteractiveChoice = @{}
$script:EntitiesFoundInZip = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)

function Get-TargetOptionSetAllowedForEntity {
    param([object]$Conn, [string]$EntityLogicalName)
    if (-not $Conn -or [string]::IsNullOrWhiteSpace($EntityLogicalName)) { return @{} }
    if ($script:TargetOptionSetAllowed.ContainsKey($EntityLogicalName)) {
        return $script:TargetOptionSetAllowed[$EntityLogicalName]
    }
    $result = @{}
    try {
        $meta = Get-CrmEntityMetadata -Conn $Conn -EntityLogicalName $EntityLogicalName -ErrorAction Stop
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

function Get-OptionSetUserChoice {
    param([string]$Key, [object]$AttrInfo, [object]$AllowedSet, [string]$EntName, [string]$Cname, [int]$ValInt)
    if ($script:OptionSetInteractiveChoice.ContainsKey($Key)) { return $script:OptionSetInteractiveChoice[$Key] }
    $optionsList = ($AttrInfo.Options | Sort-Object -Property Value | ForEach-Object { "$($_.Value)=$($_.Label)" }) -join ', '
    Write-Host ""
    Write-Host "  [Option set] Encja: $EntName, pole: $Cname – wartosc w zipie: $ValInt nie istnieje w celu." -ForegroundColor Yellow
    Write-Host "  Dozwolone w celu (numer=nazwa): $optionsList" -ForegroundColor Cyan
    $prompt = '  Wpisz numer do podstawienia (Enter=pomin, C=wyczysc pole): '
    $userInput = (Read-Host $prompt).Trim()
    if ([string]::IsNullOrWhiteSpace($userInput)) { $script:OptionSetInteractiveChoice[$Key] = 'Skip'; return 'Skip' }
    if ($userInput -eq 'C' -or $userInput -eq 'c') { $script:OptionSetInteractiveChoice[$Key] = 'Clear'; return 'Clear' }
    $num = $null
    if ([int]::TryParse($userInput, [ref]$num) -and $AllowedSet.Contains($num)) { $script:OptionSetInteractiveChoice[$Key] = $num; return $num }
    Write-Host "  Nieprawidlowy numer – pomijam (zostaw bez zmiany)." -ForegroundColor DarkYellow
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
                # Zduplikowane overriddencreatedon (bug eksportu) – zostaw tylko pierwsze
                $overrideDupNodes = @($rec.SelectNodes('.//*[@name or @Name]') | Where-Object {
                    $fn = $_.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($fn)) { $fn = $_.GetAttribute('Name') }
                    $fn -eq 'overriddencreatedon'
                })
                for ($i = 1; $i -lt $overrideDupNodes.Count; $i++) {
                    if ($overrideDupNodes[$i].ParentNode) { [void]$overrideDupNodes[$i].ParentNode.RemoveChild($overrideDupNodes[$i]); $changed = $true; $script:DuplicateOverrideRemovedCount++ }
                }
                # Lookupy do encji nieobecnych w pakiecie (msdyn_contactkpiid, msdyn_accountkpiid itd.) – usun z rekordu
                if ($stripLookupSet -and $stripLookupSet.Count -gt 0) {
                    $toStrip = @($rec.SelectNodes('.//*[@name or @Name]') | Where-Object {
                        $fn = $_.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($fn)) { $fn = $_.GetAttribute('Name') }; if ([string]::IsNullOrWhiteSpace($fn)) { $fn = $_.LocalName }
                        -not [string]::IsNullOrWhiteSpace($fn) -and $stripLookupSet.Contains($fn)
                    })
                    foreach ($n in $toStrip) { if ($n.ParentNode) { [void]$n.ParentNode.RemoveChild($n); $changed = $true; $script:LookupFieldsStrippedCount++ } }
                }
                $createdonVal = $null
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
                    if ($script:doOptionSetValidation -and $script:connTarget -and -not [string]::IsNullOrWhiteSpace($cname)) {
                        if (Invoke-OptionSetValidationForField -Ent $ent -Rec $rec -Child $child -Cname $cname -ConnTarget $script:connTarget) { $changed = $true }
                    }
                }
                # Ustaw overriddencreatedon raz na rekord (nie w petli po polach – wtedy dodawaloby sie N razy)
                if (-not [string]::IsNullOrWhiteSpace($createdonVal)) {
                    if ($overrideNode) {
                        $overrideNode.InnerText = $createdonVal
                        if ($useValueAttr) { $overrideNode.SetAttribute('value', $createdonVal) }
                    } else {
                        if ($fieldLikeParent -and $fieldLikeParent.LocalName -match '^(field|Field)$') {
                            $ns = $fieldLikeParent.NamespaceURI
                            if (-not [string]::IsNullOrWhiteSpace($ns)) {
                                $newEl = $doc.CreateElement($fieldLikeParent.LocalName, $ns)
                            } else {
                                $newEl = $doc.CreateElement($fieldLikeParent.LocalName)
                            }
                            $newEl.SetAttribute('name', 'overriddencreatedon')
                            $newEl.InnerText = $createdonVal
                            if ($useValueAttr) { $newEl.SetAttribute('value', $createdonVal) }
                            $parent = if ($fieldLikeParent.ParentNode -and $fieldLikeParent.ParentNode -ne $rec) { $fieldLikeParent.ParentNode } else { $rec }
                            [void]$parent.AppendChild($newEl)
                        } else {
                            $newEl = $doc.CreateElement('overriddencreatedon')
                            $newEl.InnerText = $createdonVal
                            [void]$rec.AppendChild($newEl)
                        }
                    }
                    $changed = $true
                    $script:overrideCount++
                }
            }
        }
        if ($script:StripFieldsNotInTarget -and $script:connTarget) {
            foreach ($ent in $entityNodes) {
                $entName = $ent.GetAttribute('name')
                if ([string]::IsNullOrWhiteSpace($entName)) { continue }
                if (-not $script:targetAllowedAttrs.ContainsKey($entName)) {
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
    foreach ($f in $Files) {
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
        $allContent = ($files | ForEach-Object { [System.IO.File]::ReadAllText($_.FullName, [System.Text.Encoding]::UTF8) }) -join "`n"
        $found = 0
        foreach ($src in $guidMap.Keys) {
            if ($allContent.IndexOf($src, [StringComparison]::OrdinalIgnoreCase) -ge 0) { $found++ }
        }
        if ($found -eq 0) {
            Write-Host 'W zipie nie ma GUIDow z IdMap (ownerid moze byc jako imie i nazwisko). Uzycie mapowania po display name (CMT_IdMap_ByDisplayName.json).' -ForegroundColor Gray
        } else {
            Write-Host ('W zipie znaleziono ' + $found + ' z ' + $guidMap.Count + ' GUIDow z IdMap - podmiana ownerow bedzie wykonana.') -ForegroundColor Gray
        }
    }
    [void]$script:EntitiesFoundInZip.Clear()
    Invoke-TransformCMTFiles -Files $files
    $entitiesList = @($script:EntitiesFoundInZip | Sort-Object)
    if ($entitiesList.Count -gt 0) {
        Write-Host ('W zipie przetworzono encje: ' + ($entitiesList -join ', ') + '.') -ForegroundColor Gray
        Write-Host 'Uwaga: W tym zipie sa tylko te encje. Jesli przy imporcie bedzie brakowac rekordow lub calych encji – dodaj do folderu Input zip z CMT zawierajacy brakujace encje i uruchom opcje 3 dla tego zipa.' -ForegroundColor Yellow
    }

    # Raport walidacji option setow (Report/Clear/Replace) + opcjonalnie interaktywna korekta i ponowna transformacja
    $didOptionSetCorrection = $false
    if ($doOptionSetValidation -and $script:OptionSetIssues -and $script:OptionSetIssues.Count -gt 0) {
        $reportDir = Split-Path $OutputZipPath -Parent
        if ([string]::IsNullOrWhiteSpace($reportDir)) { $reportDir = (Get-Location).Path }
        $reportName = "CMT_OptionSetValidation_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv"
        $reportPath = Join-Path $reportDir $reportName
        $script:OptionSetIssues | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
        Write-Host ""
        Write-Host ('Walidacja option setow: znaleziono ' + $script:OptionSetIssues.Count + ' niepasujacych wartosci (akcja: ' + $script:OptionSetValidationAction + '). Raport: ' + $reportPath) -ForegroundColor Yellow
        Write-Host '  Rekordy z wartoscia option set nieistniejaca w celu:' -ForegroundColor Yellow
        foreach ($issue in $script:OptionSetIssues) {
            $recInfo = if (-not [string]::IsNullOrWhiteSpace($issue.RecordName)) { $issue.RecordName } else { $issue.RecordId }
            Write-Host ('    Encja: ' + $issue.Entity + ' | Rekord: ' + $recInfo + ' | Pole: ' + $issue.Field + ' | Wartosc w zipie: ' + [string]$issue.ValueInZip + ' | Dozwolone w celu: ' + $issue.AllowedValues) -ForegroundColor Yellow
        }
        Write-Host ""
        Write-Host '--- Przejdzmy przez bledy: wybierz co zrobic (ta sama decyzja dla wszystkich rekordow z ta sama encja, pole i wartosc w zipie). ---' -ForegroundColor Cyan
        $seenKeys = @{}
        foreach ($issue in $script:OptionSetIssues) {
            $key = $issue.Entity + '|' + $issue.Field + '|' + [string]$issue.ValueInZip
            if ($seenKeys[$key]) { continue }
            $seenKeys[$key] = $true
            $attrInfo = $null
            if ($script:TargetOptionSetAllowed -and $script:TargetOptionSetAllowed.ContainsKey($issue.Entity) -and $script:TargetOptionSetAllowed[$issue.Entity].ContainsKey($issue.Field)) {
                $attrInfo = $script:TargetOptionSetAllowed[$issue.Entity][$issue.Field]
            }
            if (-not $attrInfo) {
                Write-Host ('  Pomijam (brak metadanych): encja=' + $issue.Entity + ', pole=' + $issue.Field) -ForegroundColor Gray
                continue
            }
            Get-OptionSetUserChoice -Key $key -AttrInfo $attrInfo -AllowedSet $attrInfo.AllowedSet -EntName $issue.Entity -Cname $issue.Field -ValInt $issue.ValueInZip
        }
        Write-Host ""
        Write-Host 'Ponawiam transformacje z Twoimi wyborami...' -ForegroundColor Cyan
        $savedAction = $script:OptionSetValidationAction
        $script:OptionSetValidationAction = 'Interactive'
        $script:OptionSetIssues.Clear()
        $tempDir2 = Join-Path $env:TEMP ([System.Guid]::NewGuid().ToString('N'))
        try {
            [System.IO.Compression.ZipFile]::ExtractToDirectory($InputZipPath, $tempDir2)
            $files2 = Get-ChildItem -Path $tempDir2 -Recurse -File
            Invoke-TransformCMTFiles -Files $files2
            if (Test-Path $OutputZipPath -PathType Leaf) { Remove-Item $OutputZipPath -Force }
            [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir2, $OutputZipPath)
            Write-Host ('Zapisano (po poprawkach option setow): ' + $OutputZipPath) -ForegroundColor Green
            $didOptionSetCorrection = $true
        } finally {
            if (Test-Path $tempDir2) { Remove-Item $tempDir2 -Recurse -Force -ErrorAction SilentlyContinue }
        }
        $script:OptionSetValidationAction = $savedAction
    }

    if (-not $didOptionSetCorrection) {
        if (Test-Path $OutputZipPath -PathType Leaf) { Remove-Item $OutputZipPath -Force }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $OutputZipPath)
        Write-Host ('Zapisano: ' + $OutputZipPath)
    }
    if ($script:overrideCount -gt 0) { Write-Host ('Ustawiono overriddencreatedon w ' + $script:overrideCount + ' rekordach (oryginalna data utworzenia).') -ForegroundColor Gray }
    if ($script:DuplicateOverrideRemovedCount -gt 0) { Write-Host ('Usunieto ' + $script:DuplicateOverrideRemovedCount + ' zduplikowanych pol overriddencreatedon w rekordach (bug eksportu).') -ForegroundColor Gray }
    if ($script:LookupFieldsStrippedCount -gt 0) { Write-Host ('Usunieto ' + $script:LookupFieldsStrippedCount + ' pol lookup (msdyn_contactkpiid, msdyn_accountkpiid, transactioncurrencyid, originatingleadid) – brak w celu.') -ForegroundColor Gray }
    if ($script:DuplicateRecordsRemovedCount -gt 0) { Write-Host ('Usunieto ' + $script:DuplicateRecordsRemovedCount + ' zduplikowanych rekordow (ten sam klucz glowny) – zapobiega bledowi CMT: Element o tym samym kluczu.') -ForegroundColor Gray }
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

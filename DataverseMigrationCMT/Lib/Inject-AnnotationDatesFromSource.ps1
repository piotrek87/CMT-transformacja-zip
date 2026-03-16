# Pobiera ze zrodla (Dataverse API) createdon i modifiedon dla rekordow annotation
# z zipa CMT i wstrzykuje je do data.xml w formacie ISO 8601.
# Mozesz uruchomic opcje 6 PO opcji 3 (zip *_ForTarget), wtedy poprawiasz tylko daty Uwag;
# wynik nadaje sie do importu bez ponownego uruchamiania opcji 3.
# Uzycie: .\Inject-AnnotationDatesFromSource.ps1 -InputZipPath ".\Output\CMT_Export.zip" [-OutputZipPath ...] [-ConfigPath ...]

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $InputZipPath,
    [Parameter(Mandatory = $false)]
    [string] $OutputZipPath,
    [Parameter(Mandatory = $false)]
    [string] $ConfigPath,
    [Parameter(Mandatory = $false)]
    [int] $BatchSize = 500
)

$ErrorActionPreference = 'Stop'
$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }

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
$configDir = Join-Path $scriptRoot '..\Config'
if ([string]::IsNullOrWhiteSpace($ConfigPath)) { $ConfigPath = Join-Path $configDir 'CMTConfig.ps1' }

if (-not (Test-Path $InputZipPath -PathType Leaf)) {
    throw "Brak pliku zip: $InputZipPath"
}
if (-not (Test-Path $ConfigPath -PathType Leaf)) {
    throw "Brak configu: $ConfigPath. Ustaw SourceConnectionString (Config\Polaczenia.txt)."
}

$config = & $ConfigPath
$srcConnStr = $config.SourceConnectionString
if ([string]::IsNullOrWhiteSpace($srcConnStr)) {
    throw "W configu brak SourceConnectionString (Config\Polaczenia.txt: ZrodloUrl, ZrodloLogin, ZrodloHaslo)."
}

if (-not (Get-Module -ListAvailable -Name 'Microsoft.Xrm.Data.PowerShell')) {
    throw "Wymagany modul: Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser"
}
Import-Module Microsoft.Xrm.Data.PowerShell -Force -ErrorAction Stop
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop

if ([string]::IsNullOrWhiteSpace($OutputZipPath)) {
    $dir = [System.IO.Path]::GetDirectoryName($InputZipPath)
    $name = [System.IO.Path]::GetFileNameWithoutExtension($InputZipPath)
    $OutputZipPath = Join-Path $dir ($name + '_WithAnnotationDates.zip')
}

Write-Host "Zip wejsciowy: $InputZipPath" -ForegroundColor Cyan
Write-Host "Zip wyjsciowy: $OutputZipPath" -ForegroundColor Cyan
Write-Host "Polaczenie ze zrodlem (pobieranie createdon/modifiedon dla Uwag)..." -ForegroundColor Gray

$tempDir = Join-Path $env:TEMP ([System.Guid]::NewGuid().ToString('N'))
try {
    [System.IO.Compression.ZipFile]::ExtractToDirectory($InputZipPath, $tempDir)
    $dataFile = Get-ChildItem -Path $tempDir -Recurse -Filter 'data.xml' -File -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $dataFile -or -not (Test-Path $dataFile.FullName)) {
        throw "W zipie nie znaleziono data.xml"
    }

    $dataPath = $dataFile.FullName
    $content = [System.IO.File]::ReadAllText($dataPath, [System.Text.Encoding]::UTF8)
    $doc = [xml]$content

    $entityNodes = @()
    foreach ($n in $doc.SelectNodes('//*[local-name()="entity" and (@name or @Name)]')) { $entityNodes += $n }
    foreach ($n in $doc.SelectNodes('//*[local-name()="Entity" and (@name or @Name)]')) {
        if ($entityNodes -notcontains $n) { $entityNodes += $n }
    }

    $annotationEntity = $entityNodes | Where-Object {
        $en = $_.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($en)) { $en = $_.GetAttribute('Name') }
        [string]::Equals($en, 'annotation', [StringComparison]::OrdinalIgnoreCase)
    } | Select-Object -First 1

    if (-not $annotationEntity) {
        Write-Host "W zipie brak encji annotation - pomijam." -ForegroundColor Yellow
        if (Test-Path $OutputZipPath -PathType Leaf) { Remove-Item $OutputZipPath -Force }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $OutputZipPath)
        Write-Host "Skopiowano zip do: $OutputZipPath" -ForegroundColor Green
        return
    }

    $recordNodes = @($annotationEntity.SelectNodes('.//*[local-name()="record" or local-name()="Record"]'))
    if ($recordNodes.Count -eq 0) {
        Write-Host "Brak rekordow annotation w zipie." -ForegroundColor Yellow
        if (Test-Path $OutputZipPath -PathType Leaf) { Remove-Item $OutputZipPath -Force }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $OutputZipPath)
        return
    }

    $annotationIds = [System.Collections.ArrayList]::new()
    foreach ($rec in $recordNodes) {
        $id = $rec.GetAttribute('id'); if ([string]::IsNullOrWhiteSpace($id)) { $id = $rec.GetAttribute('Id') }
        if ([string]::IsNullOrWhiteSpace($id)) {
            foreach ($child in $rec.ChildNodes) {
                if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                $fn = $child.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($fn)) { $fn = $child.GetAttribute('Name') }
                if ([string]::Equals($fn, 'annotationid', [StringComparison]::OrdinalIgnoreCase)) {
                    $id = $child.InnerText.Trim()
                    if ([string]::IsNullOrWhiteSpace($id)) { $id = $child.GetAttribute('value'); if ([string]::IsNullOrWhiteSpace($id)) { $id = $child.GetAttribute('Value') } }
                    break
                }
            }
        }
        $id = $id -replace '^\{|\}$', ''
        if (-not [string]::IsNullOrWhiteSpace($id)) { [void]$annotationIds.Add($id) }
    }

    $uniqueIds = @($annotationIds | Select-Object -Unique)
    Write-Host "Rekordow annotation w zipie: $($recordNodes.Count), unikalnych ID: $($uniqueIds.Count)" -ForegroundColor Gray

    $srcConn = Get-CrmConnection -ConnectionString $srcConnStr -ErrorAction Stop
    $datesByAnnotationId = @{}

    for ($offset = 0; $offset -lt $uniqueIds.Count; $offset += $BatchSize) {
        $batch = $uniqueIds[$offset..([Math]::Min($offset + $BatchSize - 1, $uniqueIds.Count - 1))]
        $conditions = ($batch | ForEach-Object { [char]0x3C + "condition attribute=`"annotationid`" operator=`"eq`" value=`"$_`" />" }) -join "`n      "
        $fetchXml = @"
<fetch no-lock="true" mapping="logical">
  <entity name="annotation">
    <attribute name="annotationid" />
    <attribute name="createdon" />
    <attribute name="modifiedon" />
    <filter type="or">
      $conditions
    </filter>
  </entity>
</fetch>
"@
        try {
            $result = Get-CrmRecordsByFetch -conn $srcConn -Fetch $fetchXml -ErrorAction Stop
            $records = @($result.CrmRecords)
            foreach ($r in $records) {
                $aid = $null
                if ($r.PSObject.Properties['annotationid']) { $aid = [string]$r.annotationid -replace '^\{|\}$', '' }
                if ([string]::IsNullOrWhiteSpace($aid)) { continue }
                $createdon = $null
                $modifiedon = $null
                if ($r.PSObject.Properties['createdon'] -and $null -ne $r.createdon) {
                    $createdon = $r.createdon
                    if ($createdon -is [DateTime]) { $createdon = $createdon.ToString('o') }
                }
                if ($r.PSObject.Properties['modifiedon'] -and $null -ne $r.modifiedon) {
                    $modifiedon = $r.modifiedon
                    if ($modifiedon -is [DateTime]) { $modifiedon = $modifiedon.ToString('o') }
                }
                $key = $aid.Trim().ToLowerInvariant()
                $datesByAnnotationId[$key] = @{ CreatedOn = $createdon; ModifiedOn = $modifiedon }
            }
        } catch {
            Write-Host ('Ostrzezenie: batch ' + ($offset+1) + '-' + ($offset+$batch.Count) + ': ' + $_.Message) -ForegroundColor Yellow
        }
    }

    Write-Host "Pobrano daty ze zrodla dla $($datesByAnnotationId.Count) rekordow." -ForegroundColor Gray

    $firstRecField = $recordNodes[0].SelectSingleNode('.//*[local-name()="field" or local-name()="Field"]')
    $fieldLocalName = if ($firstRecField) { $firstRecField.LocalName } else { 'field' }
    $fieldNs = if ($firstRecField -and $firstRecField.NamespaceURI) { $firstRecField.NamespaceURI } else { [string]::Empty }
    $injectedCount = 0
    $missingFromSourceCount = 0
    foreach ($rec in $recordNodes) {
        $id = $rec.GetAttribute('id'); if ([string]::IsNullOrWhiteSpace($id)) { $id = $rec.GetAttribute('Id') }
        if ([string]::IsNullOrWhiteSpace($id)) {
            foreach ($child in $rec.ChildNodes) {
                if ($child.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                $fn = $child.GetAttribute('name'); if ([string]::IsNullOrWhiteSpace($fn)) { $fn = $child.GetAttribute('Name') }
                if ([string]::Equals($fn, 'annotationid', [StringComparison]::OrdinalIgnoreCase)) {
                    $id = $child.InnerText.Trim()
                    if ([string]::IsNullOrWhiteSpace($id)) { $id = $child.GetAttribute('value') }
                    break
                }
            }
        }
        $idNorm = ($id -replace '^\{|\}$', '').Trim()
        if ([string]::IsNullOrWhiteSpace($idNorm)) { continue }
        $idKey = $idNorm.ToLowerInvariant()

        $existingCreated = $rec.SelectSingleNode('.//*[local-name()="field" or local-name()="Field"][@name="createdon" or @Name="createdon"]')
        $existingModified = $rec.SelectSingleNode('.//*[local-name()="field" or local-name()="Field"][@name="modifiedon" or @Name="modifiedon"]')
        if (-not $existingCreated) { $existingCreated = $rec.SelectSingleNode('.//*[local-name()="createdon"]') }
        if (-not $existingModified) { $existingModified = $rec.SelectSingleNode('.//*[local-name()="modifiedon"]') }

        if (-not $datesByAnnotationId.ContainsKey($idKey)) {
            $missingFromSourceCount++
            # Gdy zrodlo nie zwrocilo rekordu: uzyj dat juz obecnych w zipie (np. z eksportu CMT), zeby overriddencreatedon bylo uzupelnione
            $fromZipCreated = $null
            $fromZipModified = $null
            if ($existingCreated) {
                $fromZipCreated = $existingCreated.InnerText.Trim()
                if ([string]::IsNullOrWhiteSpace($fromZipCreated)) { $fromZipCreated = $existingCreated.GetAttribute('value'); if ([string]::IsNullOrWhiteSpace($fromZipCreated)) { $fromZipCreated = $existingCreated.GetAttribute('Value') } }
            }
            if ($existingModified) {
                $fromZipModified = $existingModified.InnerText.Trim()
                if ([string]::IsNullOrWhiteSpace($fromZipModified)) { $fromZipModified = $existingModified.GetAttribute('value'); if ([string]::IsNullOrWhiteSpace($fromZipModified)) { $fromZipModified = $existingModified.GetAttribute('Value') } }
            }
            if ([string]::IsNullOrWhiteSpace($fromZipCreated) -and [string]::IsNullOrWhiteSpace($fromZipModified)) { continue }
            $dates = @{ CreatedOn = $fromZipCreated; ModifiedOn = $fromZipModified }
        } else {
            $dates = $datesByAnnotationId[$idKey]
        }

        $existingOverride = $rec.SelectSingleNode('.//*[local-name()="field" or local-name()="Field"][@name="overriddencreatedon" or @Name="overriddencreatedon"]')
        if (-not $existingOverride) { $existingOverride = $rec.SelectSingleNode('.//*[local-name()="overriddencreatedon"]') }

        $valueForOverride = $dates.CreatedOn
        if ([string]::IsNullOrWhiteSpace($valueForOverride)) { $valueForOverride = $dates.ModifiedOn }
        $needOverride = -not [string]::IsNullOrWhiteSpace($valueForOverride)
        if ($existingCreated -and $existingModified -and $existingOverride -and -not $needOverride) { continue }

        $firstField = $rec.SelectSingleNode('.//*[local-name()="field" or local-name()="Field"]')
        $parent = if ($firstField -and $firstField.ParentNode) { $firstField.ParentNode } else { $rec }
        $needCreated = $dates.CreatedOn -and -not $existingCreated
        $needModified = $dates.ModifiedOn -and -not $existingModified

        $createField = {
            param($name, $val)
            if ([string]::IsNullOrWhiteSpace($fieldNs)) {
                $el = $doc.CreateElement($fieldLocalName)
            } else {
                $el = $doc.CreateElement($fieldLocalName, $fieldNs)
            }
            $el.SetAttribute('name', $name)
            $el.SetAttribute('value', $val)
            $el.InnerText = $val
            return $el
        }

        $createdValIso = if ($dates.CreatedOn) { ConvertTo-DateTimeIso $dates.CreatedOn } else { $null }
        $modifiedValIso = if ($dates.ModifiedOn) { ConvertTo-DateTimeIso $dates.ModifiedOn } else { $null }
        $overrideValIso = if ($valueForOverride) { ConvertTo-DateTimeIso $valueForOverride } else { $null }

        if ($needCreated -and $createdValIso) {
            $el = & $createField -name 'createdon' -val $createdValIso
            if ($firstField) { [void]$parent.InsertBefore($el, $firstField) } else { [void]$parent.AppendChild($el) }
            $injectedCount++
        }
        if ($needModified -and $modifiedValIso) {
            $el = & $createField -name 'modifiedon' -val $modifiedValIso
            if ($firstField) { [void]$parent.InsertBefore($el, $firstField) } else { [void]$parent.AppendChild($el) }
            $injectedCount++
        }
        if ($needOverride -and $overrideValIso) {
            if ($existingOverride) {
                $existingOverride.InnerText = $overrideValIso
                try { $existingOverride.SetAttribute('value', $overrideValIso) } catch { }
                try { $existingOverride.SetAttribute('Value', $overrideValIso) } catch { }
            } else {
                $el = & $createField -name 'overriddencreatedon' -val $overrideValIso
                if ($firstField) { [void]$parent.InsertBefore($el, $firstField) } else { [void]$parent.AppendChild($el) }
            }
            $injectedCount++
        }
    }

    $doc.Save($dataPath)
    Write-Host "Wstrzykieto createdon/modifiedon/overriddencreatedon w $injectedCount miejsc (rekordy annotation)." -ForegroundColor Green
    if ($missingFromSourceCount -gt 0) {
        Write-Host "UWAGA: Dla $missingFromSourceCount rekordow nie znaleziono dat w zrodle - przy imporcie otrzymaja date importu (Data utworzenia = data importu)." -ForegroundColor Yellow
    }

    if (Test-Path $OutputZipPath -PathType Leaf) { Remove-Item $OutputZipPath -Force }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $OutputZipPath)
    Write-Host "Zapisano: $OutputZipPath" -ForegroundColor Green
    Write-Host 'Daty sa w formacie ISO 8601. Jesli uruchomiles opcje 6 na zipie po opcji 3, ten zip mozesz od razu importowac (bez ponownej opcji 3).' -ForegroundColor Cyan
} finally {
    if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
}

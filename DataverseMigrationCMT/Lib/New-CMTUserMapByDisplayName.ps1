# Generuje plik User Map dla CMT (Configuration Migration Tool) na podstawie dopasowania
# użytkowników po imieniu i nazwisku (lub fullname) między źródłem a celem.
# Wymaga: Microsoft.Xrm.Data.PowerShell, połączenia w Config (Polaczenia.txt / CMTConfig.ps1).
# Użycie: .\New-CMTUserMapByDisplayName.ps1 [-ConfigPath ...] [-OutputPath ...]

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string] $ConfigPath,
    [Parameter(Mandatory = $false)]
    [string] $OutputPath,
    [Parameter(Mandatory = $false)]
    [switch] $UseDomainName
)

$ErrorActionPreference = 'Stop'
$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$configDir = Join-Path $scriptRoot '..\Config'
if (-not $ConfigPath) { $ConfigPath = Join-Path $configDir 'CMTConfig.ps1' }
if (-not (Test-Path $ConfigPath)) {
    Write-Error "Brak pliku konfiguracji: $ConfigPath"
}
$config = & $ConfigPath
$srcConnStr = $config.SourceConnectionString
$tgtConnStr = $config.TargetConnectionString
if ([string]::IsNullOrWhiteSpace($srcConnStr) -or [string]::IsNullOrWhiteSpace($tgtConnStr)) {
    Write-Error "W configu musza byc ustawione SourceConnectionString i TargetConnectionString (np. w Config\Polaczenia.txt)."
}

if (-not (Get-Module -ListAvailable -Name 'Microsoft.Xrm.Data.PowerShell')) {
    Write-Error "Wymagany modul: Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser"
}
Import-Module Microsoft.Xrm.Data.PowerShell -Force -ErrorAction Stop

if (-not $OutputPath) {
    $outDir = $config.ExportOutputDirectory
    if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = Join-Path $scriptRoot '..\Output' }
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    $OutputPath = Join-Path $outDir 'CMT_UserMap_ByDisplayName.xml'
}

function Get-UserDisplayName {
    param($Rec)
    $first = ''
    $last = ''
    if ($Rec.PSObject.Properties['firstname']) { $first = [string]$Rec.firstname }
    if ($Rec.PSObject.Properties['lastname'])  { $last  = [string]$Rec.lastname }
    $full = ($first.Trim() + ' ' + $last.Trim()).Trim()
    if ([string]::IsNullOrWhiteSpace($full) -and $Rec.PSObject.Properties['fullname']) {
        $full = [string]$Rec.fullname
    }
    return $full
}

Write-Host "Polaczenie ze zrodlem..."
$srcConn = Get-CrmConnection -ConnectionString $srcConnStr -ErrorAction Stop
Write-Host "Polaczenie z celem..."
$tgtConn = Get-CrmConnection -ConnectionString $tgtConnStr -ErrorAction Stop

# Bez atrybutu top – Get-CrmRecordsByFetch sam stosuje stronicowanie (top + page sie wykluczaja)
$fetchSource = @'
<fetch no-lock="true">
  <entity name="systemuser">
    <attribute name="systemuserid" />
    <attribute name="fullname" />
    <attribute name="firstname" />
    <attribute name="lastname" />
    <attribute name="domainname" />
    <attribute name="internalemailaddress" />
  </entity>
</fetch>
'@
$fetchTarget = $fetchSource

$srcUsers = Get-CrmRecordsByFetch -conn $srcConn -Fetch $fetchSource -ErrorAction Stop
$tgtUsers = Get-CrmRecordsByFetch -conn $tgtConn -Fetch $fetchTarget -ErrorAction Stop
$srcList = @($srcUsers.CrmRecords)
$tgtList = @($tgtUsers.CrmRecords)

# Indeks celu: DisplayName (znormalizowany) -> rekord
$tgtByDisplayName = @{}
foreach ($r in $tgtList) {
    $dn = Get-UserDisplayName -Rec $r
    if ([string]::IsNullOrWhiteSpace($dn)) {
        if ($r.PSObject.Properties['domainname']) { $dn = [string]$r.domainname }
        elseif ($r.PSObject.Properties['internalemailaddress']) { $dn = [string]$r.internalemailaddress }
    }
    if (-not [string]::IsNullOrWhiteSpace($dn)) {
        $key = $dn.Trim().ToLowerInvariant()
        if (-not $tgtByDisplayName.ContainsKey($key)) {
            $tgtByDisplayName[$key] = $r
        }
    }
}

# Mapowanie: źródło -> cel (po DisplayName)
$mappings = [System.Collections.ArrayList]::new()
foreach ($r in $srcList) {
    $srcDn = Get-UserDisplayName -Rec $r
    if ([string]::IsNullOrWhiteSpace($srcDn)) {
        if ($r.PSObject.Properties['domainname']) { $srcDn = [string]$r.domainname }
        elseif ($r.PSObject.Properties['internalemailaddress']) { $srcDn = [string]$r.internalemailaddress }
    }
    if ([string]::IsNullOrWhiteSpace($srcDn)) { continue }
    $key = $srcDn.Trim().ToLowerInvariant()
    $tgtRec = $tgtByDisplayName[$key]
    $srcId = $r.PSObject.Properties['systemuserid']; if (-not $srcId) { $srcId = $r.systemuserid }
    $srcDomain = ''; if ($r.PSObject.Properties['domainname']) { $srcDomain = [string]$r.domainname }
    $tgtDn = ''
    $tgtDomain = ''
    $tgtId = $null
    if ($tgtRec) {
        $tgtDn = Get-UserDisplayName -Rec $tgtRec
        if ($tgtRec.PSObject.Properties['domainname']) { $tgtDomain = [string]$tgtRec.domainname }
        $tgtId = $tgtRec.PSObject.Properties['systemuserid']; if (-not $tgtId) { $tgtId = $tgtRec.systemuserid }
    }
    $null = $mappings.Add([PSCustomObject]@{
        SourceDisplayName = $srcDn.Trim()
        SourceDomainName   = $srcDomain
        SourceId          = $srcId
        TargetDisplayName = $tgtDn
        TargetDomainName  = $tgtDomain
        TargetId          = $tgtId
        Matched           = ($null -ne $tgtRec)
    })
}

# Zapis XML – format czytelny dla CMT (wiele narzędzi oczekuje source -> new/target).
# CMT w GUI często generuje plik z identyfikatorami użytkowników źródła; pole "New" to użytkownik w celu.
# Używamy DomainName jako identyfikatora (bo CMT zwykle mapuje po domena\użytkownik), ewentualnie DisplayName.
$sb = [System.Text.StringBuilder]::new()
[void]$sb.AppendLine('<?xml version="1.0" encoding="utf-8"?>')
[void]$sb.AppendLine('<!-- User Map wygenerowany przez New-CMTUserMapByDisplayName.ps1 (dopasowanie po imieniu i nazwisku). -->')
[void]$sb.AppendLine('<!-- Uzyj UserMapFilePath w CMTConfig.ps1 lub -UserMapFilePath przy Import-CrmDataFile. -->')
[void]$sb.AppendLine('<UserMap>')
foreach ($m in $mappings) {
    $srcKey = if ($UseDomainName -and -not [string]::IsNullOrWhiteSpace($m.SourceDomainName)) { $m.SourceDomainName } else { $m.SourceDisplayName }
    $tgtKey = if ($UseDomainName -and -not [string]::IsNullOrWhiteSpace($m.TargetDomainName)) { $m.TargetDomainName } else { $m.TargetDisplayName }
    if ([string]::IsNullOrWhiteSpace($tgtKey)) { $tgtKey = $srcKey }
    [void]$sb.AppendLine('  <User>')
    [void]$sb.AppendLine('    <Source>' + [System.Security.SecurityElement]::Escape($srcKey) + '</Source>')
    [void]$sb.AppendLine('    <New>'    + [System.Security.SecurityElement]::Escape($tgtKey) + '</New>')
    [void]$sb.AppendLine('    <SourceDisplayName>' + [System.Security.SecurityElement]::Escape($m.SourceDisplayName) + '</SourceDisplayName>')
    [void]$sb.AppendLine('    <TargetDisplayName>' + [System.Security.SecurityElement]::Escape($m.TargetDisplayName) + '</TargetDisplayName>')
    [void]$sb.AppendLine('    <Matched>' + $(if ($m.Matched) { 'true' } else { 'false' }) + '</Matched>')
    [void]$sb.AppendLine('  </User>')
}
[void]$sb.AppendLine('</UserMap>')

$outDir = [System.IO.Path]::GetDirectoryName($OutputPath)
if (-not [string]::IsNullOrEmpty($outDir) -and -not [System.IO.Directory]::Exists($outDir)) {
    [System.IO.Directory]::CreateDirectory($outDir) | Out-Null
}
[System.IO.File]::WriteAllText($OutputPath, $sb.ToString(), [System.Text.UTF8Encoding]::new($false))

# Zapis mapowania GUID (zrodlo -> cel) dla Transform-CMTZip – podmiana ownerid w zipie CMT
# Wyciagamy czysty GUID (CRM moze zwracac "guid systemuserid=xxx" zamiast "xxx")
function Get-PureGuid {
    param([string]$Val)
    if ([string]::IsNullOrWhiteSpace($Val)) { return $null }
    $s = $Val.Trim()
    if ($s -match '^[0-9a-fA-F\-]{36}$') { return $s.ToLowerInvariant() }
    if ($s -match '=([0-9a-fA-F\-]{36})$') { return $Matches[1].ToLowerInvariant() }
    $s = $s -replace '^\{|\}$', ''
    if ($s -match '^[0-9a-fA-F\-]{36}$') { return $s.ToLowerInvariant() }
    return $s.ToLowerInvariant()
}
$idMapPath = [System.IO.Path]::Combine($outDir, 'CMT_IdMap_SystemUser.json')
$guidMap = @{}
foreach ($m in $mappings) {
    if ($m.Matched -and $m.SourceId -and $m.TargetId) {
        $srcGuid = Get-PureGuid ([string]$m.SourceId)
        $tgtGuid = Get-PureGuid ([string]$m.TargetId)
        if (-not [string]::IsNullOrWhiteSpace($srcGuid) -and -not [string]::IsNullOrWhiteSpace($tgtGuid)) {
            $guidMap[$srcGuid] = $tgtGuid
        }
    }
}
$idMapJson = $guidMap | ConvertTo-Json -Compress
[System.IO.File]::WriteAllText($idMapPath, $idMapJson, [System.Text.UTF8Encoding]::new($false))
Write-Host "IdMap (GUID): $idMapPath - uzyj w Transform-CMTZip.ps1 do podmiany ownerid w zipie."

# Mapowanie DisplayName (zrodlo) -> Target GUID – gdy CMT eksportuje wlasciciela jako imie i nazwisko zamiast GUID
$byDisplayNamePath = [System.IO.Path]::Combine($outDir, 'CMT_IdMap_ByDisplayName.json')
$displayNameToGuid = @{}
foreach ($m in $mappings) {
    if ($m.Matched -and $m.TargetId -and -not [string]::IsNullOrWhiteSpace($m.SourceDisplayName)) {
        $key = [regex]::Replace($m.SourceDisplayName.Trim(), '\s+', ' ').ToLowerInvariant()
        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $displayNameToGuid[$key] = [string]$m.TargetId
        }
    }
}
if ($displayNameToGuid.Count -gt 0) {
    $byDisplayNameJson = $displayNameToGuid | ConvertTo-Json -Compress
    [System.IO.File]::WriteAllText($byDisplayNamePath, $byDisplayNameJson, [System.Text.UTF8Encoding]::new($false))
    Write-Host "IdMap (imie i nazwisko -> GUID celu): $byDisplayNamePath - gdy w zipie ownerid to tekst (display name)."
}

$matched = ($mappings | Where-Object { $_.Matched }).Count
$unmatched = ($mappings | Where-Object { -not $_.Matched }).Count
Write-Host "Zapisano: $OutputPath"
Write-Host "Zrodlo uzytkownikow: $($srcList.Count) | Cel: $($tgtList.Count) | Zmapowani (po imieniu i nazwisku): $matched | Bez dopasowania: $unmatched"
if ($unmatched -gt 0) {
    Write-Host "Uzytkownicy bez dopasowania (bedą zmapowani na siebie lub wymagaja recznego wpisu w User Map):" -ForegroundColor Yellow
    $mappings | Where-Object { -not $_.Matched } | ForEach-Object { Write-Host "  - $($_.SourceDisplayName)" }
}
Write-Host ('Ustaw w Config\CMTConfig.ps1: UserMapFilePath = "' + $OutputPath + '" lub podaj -UserMapFilePath przy imporcie.')

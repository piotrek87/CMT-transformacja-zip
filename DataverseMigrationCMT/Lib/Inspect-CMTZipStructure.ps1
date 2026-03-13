# Tymczasowy skrypt: rozpakowuje zip i pokazuje fragment data.xml (struktura record/contactid)
param([string]$ZipPath = (Join-Path $PSScriptRoot '..\Input\contact.zip'))
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
$tempDir = Join-Path $env:TEMP ('cmt_inspect_' + [Guid]::NewGuid().ToString('N').Substring(0,8))
[System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $tempDir)
$dataFile = Get-ChildItem $tempDir -Recurse -Filter 'data.xml' -File | Select-Object -First 1
if (-not $dataFile) { $dataFile = Get-ChildItem $tempDir -Recurse -Filter '*.xml' -File | Select-Object -First 1 }
if ($dataFile) {
    $text = [System.IO.File]::ReadAllText($dataFile.FullName, [System.Text.Encoding]::UTF8)
    # Pokaz pierwsze 6000 znakow - entity contact i poczatek rekordow
    $idx = $text.IndexOf('<entity name="contact"', [StringComparison]::OrdinalIgnoreCase)
    if ($idx -lt 0) { $idx = $text.IndexOf('contact', [StringComparison]::OrdinalIgnoreCase) }
    if ($idx -lt 0) { $idx = 0 }
    $snippet = $text.Substring($idx, [Math]::Min(6500, $text.Length - $idx))
    Write-Host "=== Fragment data.xml (od entity contact) ==="
    Write-Host $snippet
} else {
    Write-Host "Brak pliku data.xml w zipie. Zawartosc:"; Get-ChildItem $tempDir -Recurse | ForEach-Object { $_.FullName }
}
Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue

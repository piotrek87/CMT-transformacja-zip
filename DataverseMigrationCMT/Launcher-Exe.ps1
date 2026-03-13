# Launcher do budowania MigracjaCMT.exe (ps2exe).
# Exe uruchamia menu ze skryptow z dysku (ten sam katalog co exe). Opcja 2 wywoluje Start-CMTMigration.ps1 z dysku,
# wiec latanie schematu CMT (activitypointer, workflow, processstage, listmember) dziala przy uruchomieniu przez exe.

$exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
$root = [System.IO.Path]::GetDirectoryName($exePath)
$menuScript = Join-Path $root 'Start-CMTMigrationMenu.ps1'
if (-not (Test-Path $menuScript)) {
    Write-Host "Brak pliku Start-CMTMigrationMenu.ps1 w: $root"
    Read-Host "Nacisnij Enter"
    exit 1
}
& $menuScript

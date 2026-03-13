# Launcher do budowania .exe (ps2exe).
# Exe uruchamia GUI w osobnym procesie PowerShell (normalny host), zeby dzialalo
# logowanie interaktywne Dynamics i modul Microsoft.Xrm.Data.PowerShell.

$exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
$root = [System.IO.Path]::GetDirectoryName($exePath)
$guiScript = Join-Path $root 'Start-MigrationGUI.ps1'
$argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden', '-File', "`"$guiScript`"")
Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -WorkingDirectory $root

<#
.SYNOPSIS
    GUI migracji Dataverse - logowanie przez okno Dynamics (dwa polaczenia: zrodlo i cel).
.DESCRIPTION
    Uruchom z konsoli: .\Start-MigrationGUI.ps1 lub dwuklik MigracjaDataverse.exe (exe uruchamia ten skrypt w PowerShell).
    Wymaga: modul Microsoft.Xrm.Data.PowerShell.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:Root = $PSScriptRoot
if (-not $script:Root -and $MyInvocation.MyCommand.Path) { $script:Root = Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $script:Root) { $script:Root = (Get-Location).Path }

# Modul Xrm wymaga hosta PowerShell >= 1.0. Gdy PSRunspace (np. z Cursora) - odpal w zwyklym PowerShellu i wyjdz.
$hostName = $Host.Name
$hostVersion = $Host.Version
$isBadHost = ($hostName -match 'PSRunspace' -or ($hostVersion -and $hostVersion.Major -lt 1))
if ($isBadHost) {
    Start-Process -FilePath "powershell.exe" -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$script:Root\Start-MigrationGUI.ps1`"") -WorkingDirectory $script:Root
    exit 0
}

$script:SourceConn = $null
$script:TargetConn = $null
$script:EntityList = @()
$script:EntityRecordCounts = @{}

# Zaladuj modul i biblioteki (w tym samym procesie - dziala gdy GUI uruchomione przez powershell.exe)
try {
    Import-Module Microsoft.Xrm.Data.PowerShell -Force -ErrorAction Stop
} catch {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    [System.Windows.Forms.MessageBox]::Show("Nie mozna zaladowac modulu Microsoft.Xrm.Data.PowerShell. Uruchom przez MigracjaDataverse.bat lub MigracjaDataverse.exe (w folderze aplikacji). Blad: $_", "Blad", "OK", "Error")
    exit 1
}
$libPath = Join-Path $script:Root 'Lib'
. (Join-Path $libPath 'Connect-Dataverse.ps1')
. (Join-Path $libPath 'Get-EntityMetadata.ps1')
. (Join-Path $libPath 'Get-MigrationOrder.ps1')
. (Join-Path $libPath 'Migrate-EntityData.ps1')
$configDir = Join-Path $script:Root 'Config'
$script:ConfigPath = Join-Path $configDir 'MigrationConfig.ps1'
if (Test-Path $script:ConfigPath) { $script:Config = . $script:ConfigPath }
else {
    $script:Config = @{
        SystemEntitiesToSkip = @('bulkdeleteoperation','asyncoperation','workflow','pluginassembly')
        EntityOrderPriority = @('systemuser','team','account','contact','lead','opportunity','activitypointer','email','task','appointment')
        BpfEntitySuffix = 'process'
    }
}
$script:GuiEntityDefaultTargetLookupStr = ''
$script:GuiEntityLookupResolveByName = @()

$form = New-Object System.Windows.Forms.Form
$form.Text = "Migracja Dataverse"
$form.Size = New-Object System.Drawing.Size(800, 820)
$form.MinimumSize = New-Object System.Drawing.Size(720, 680)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox = $true

$pad = 12
[int]$y = 12

# --- ZRODLO ---
$grpSource = New-Object System.Windows.Forms.GroupBox
$grpSource.Location = New-Object System.Drawing.Point($pad, $y)
$grpSource.Size = New-Object System.Drawing.Size(330, 120)
$grpSource.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$grpSource.Text = " ZRODLO "
$grpSource.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($grpSource)

$chkUseSavedSource = New-Object System.Windows.Forms.CheckBox
$chkUseSavedSource.Location = New-Object System.Drawing.Point(12, 22)
$chkUseSavedSource.Size = New-Object System.Drawing.Size(300, 20)
$chkUseSavedSource.Text = "Uzyj zapisanych login/haslo (plik Config\LoginHaslo.txt)"
$chkUseSavedSource.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$grpSource.Controls.Add($chkUseSavedSource)

$lblUrlSource = New-Object System.Windows.Forms.Label
$lblUrlSource.Location = New-Object System.Drawing.Point(12, 44)
$lblUrlSource.Size = New-Object System.Drawing.Size(35, 18)
$lblUrlSource.Text = "URL:"
$lblUrlSource.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$grpSource.Controls.Add($lblUrlSource)
$txtUrlSource = New-Object System.Windows.Forms.TextBox
$txtUrlSource.Location = New-Object System.Drawing.Point(50, 42)
$txtUrlSource.Size = New-Object System.Drawing.Size(268, 22)
$txtUrlSource.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$grpSource.Controls.Add($txtUrlSource)

$btnConnectSource = New-Object System.Windows.Forms.Button
$btnConnectSource.Location = New-Object System.Drawing.Point(12, 70)
$btnConnectSource.Size = New-Object System.Drawing.Size(200, 32)
$btnConnectSource.Text = "Polacz ze zrodlem..."
$btnConnectSource.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grpSource.Controls.Add($btnConnectSource)

$lblStatusSource = New-Object System.Windows.Forms.Label
$lblStatusSource.Location = New-Object System.Drawing.Point(220, 76)
$lblStatusSource.Size = New-Object System.Drawing.Size(100, 22)
$lblStatusSource.Text = "nie polaczono"
$lblStatusSource.ForeColor = [System.Drawing.Color]::Gray
$grpSource.Controls.Add($lblStatusSource)

$y = [int]$y + 128

# --- CEL ---
$grpTarget = New-Object System.Windows.Forms.GroupBox
$grpTarget.Location = New-Object System.Drawing.Point($pad, $y)
$grpTarget.Size = New-Object System.Drawing.Size(330, 120)
$grpTarget.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$grpTarget.Text = " CEL "
$grpTarget.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($grpTarget)

$chkUseSavedTarget = New-Object System.Windows.Forms.CheckBox
$chkUseSavedTarget.Location = New-Object System.Drawing.Point(12, 22)
$chkUseSavedTarget.Size = New-Object System.Drawing.Size(300, 20)
$chkUseSavedTarget.Text = "Uzyj zapisanych login/haslo (plik Config\LoginHaslo.txt)"
$chkUseSavedTarget.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$grpTarget.Controls.Add($chkUseSavedTarget)

$lblUrlTarget = New-Object System.Windows.Forms.Label
$lblUrlTarget.Location = New-Object System.Drawing.Point(12, 44)
$lblUrlTarget.Size = New-Object System.Drawing.Size(35, 18)
$lblUrlTarget.Text = "URL:"
$lblUrlTarget.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$grpTarget.Controls.Add($lblUrlTarget)
$txtUrlTarget = New-Object System.Windows.Forms.TextBox
$txtUrlTarget.Location = New-Object System.Drawing.Point(50, 42)
$txtUrlTarget.Size = New-Object System.Drawing.Size(268, 22)
$txtUrlTarget.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$grpTarget.Controls.Add($txtUrlTarget)

$btnConnectTarget = New-Object System.Windows.Forms.Button
$btnConnectTarget.Location = New-Object System.Drawing.Point(12, 70)
$btnConnectTarget.Size = New-Object System.Drawing.Size(200, 32)
$btnConnectTarget.Text = "Polacz z celem..."
$btnConnectTarget.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grpTarget.Controls.Add($btnConnectTarget)

$lblStatusTarget = New-Object System.Windows.Forms.Label
$lblStatusTarget.Location = New-Object System.Drawing.Point(220, 76)
$lblStatusTarget.Size = New-Object System.Drawing.Size(100, 22)
$lblStatusTarget.Text = "nie polaczono"
$lblStatusTarget.ForeColor = [System.Drawing.Color]::Gray
$grpTarget.Controls.Add($lblStatusTarget)

$y = [int]$y + 128

$y = [int]$y + 32

$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Location = New-Object System.Drawing.Point($pad, $y)
$btnLoad.Size = New-Object System.Drawing.Size(160, 28)
$btnLoad.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$btnLoad.Text = "Pobierz liste encji"
$btnLoad.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($btnLoad)

$btnCountRecords = New-Object System.Windows.Forms.Button
$btnCountRecords.Location = New-Object System.Drawing.Point(178, $y)
$btnCountRecords.Size = New-Object System.Drawing.Size(140, 28)
$btnCountRecords.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$btnCountRecords.Text = "Policz rekordy"
$btnCountRecords.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnCountRecords.Enabled = $false
$form.Controls.Add($btnCountRecords)
$y = [int]$y + 36

$lblEntities = New-Object System.Windows.Forms.Label
$lblEntities.Location = New-Object System.Drawing.Point($pad, $y)
$lblEntities.Size = New-Object System.Drawing.Size(500, 18)
$lblEntities.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$lblEntities.Text = "Encje do migracji (zaznacz wybrane lub zostaw wszystkie):"
$lblEntities.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($lblEntities)
$y = [int]$y + 24

$listEntities = New-Object System.Windows.Forms.CheckedListBox
$listEntities.Location = New-Object System.Drawing.Point($pad, $y)
$listEntities.Size = New-Object System.Drawing.Size(776, 200)
$listEntities.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$listEntities.CheckOnClick = $true
$listEntities.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($listEntities)
$y = [int]$y + 208

$chkOnlySelected = New-Object System.Windows.Forms.CheckBox
$chkOnlySelected.Location = New-Object System.Drawing.Point($pad, $y)
$chkOnlySelected.Size = New-Object System.Drawing.Size(280, 22)
$chkOnlySelected.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$chkOnlySelected.Text = "Migruj tylko zaznaczone encje"
$chkOnlySelected.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$chkOnlySelected.Checked = $true
$form.Controls.Add($chkOnlySelected)

$chkExcludeSelected = New-Object System.Windows.Forms.CheckBox
$chkExcludeSelected.Location = New-Object System.Drawing.Point(300, $y)
$chkExcludeSelected.Size = New-Object System.Drawing.Size(260, 22)
$chkExcludeSelected.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$chkExcludeSelected.Text = "Wyklucz zaznaczone encje"
$chkExcludeSelected.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$chkExcludeSelected.Checked = $false
$form.Controls.Add($chkExcludeSelected)
$y = [int]$y + 30

$chkOnlyWithRecords = New-Object System.Windows.Forms.CheckBox
$chkOnlyWithRecords.Location = New-Object System.Drawing.Point($pad, $y)
$chkOnlyWithRecords.Size = New-Object System.Drawing.Size(520, 22)
$chkOnlyWithRecords.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left
$chkOnlyWithRecords.Text = "Tylko encje z rekordami i ich zaleznosci (pomija puste tabele, kolejnosc wedlug relacji)"
$chkOnlyWithRecords.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$chkOnlyWithRecords.Checked = $true
$form.Controls.Add($chkOnlyWithRecords)
$y = [int]$y + 30

$lblMode = New-Object System.Windows.Forms.Label
$lblMode.Location = New-Object System.Drawing.Point($pad, $y)
$lblMode.Size = New-Object System.Drawing.Size(90, 20)
$lblMode.Text = "Tryb migracji:"
$lblMode.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($lblMode)
$comboMode = New-Object System.Windows.Forms.ComboBox
$comboMode.Location = New-Object System.Drawing.Point(102, ([int]$y - 2))
$comboMode.Size = New-Object System.Drawing.Size(220, 24)
$comboMode.DropDownStyle = "DropDownList"
[void]$comboMode.Items.Add("Create")
[void]$comboMode.Items.Add("Update")
[void]$comboMode.Items.Add("Upsert")
$comboMode.SelectedIndex = 2
$form.Controls.Add($comboMode)
$y = [int]$y + 28

$lblMatch = New-Object System.Windows.Forms.Label
$lblMatch.Location = New-Object System.Drawing.Point($pad, $y)
$lblMatch.Size = New-Object System.Drawing.Size(85, 20)
$lblMatch.Text = "Dopasowanie:"
$lblMatch.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($lblMatch)
$comboMatchBy = New-Object System.Windows.Forms.ComboBox
$comboMatchBy.Location = New-Object System.Drawing.Point(102, ([int]$y - 2))
$comboMatchBy.Size = New-Object System.Drawing.Size(200, 24)
$comboMatchBy.DropDownStyle = "DropDownList"
[void]$comboMatchBy.Items.Add("Id")
[void]$comboMatchBy.Items.Add("IdThenName")
[void]$comboMatchBy.Items.Add("Name")
[void]$comboMatchBy.Items.Add("Custom")
$comboMatchBy.SelectedIndex = 1
$form.Controls.Add($comboMatchBy)
$lblCustomAttr = New-Object System.Windows.Forms.Label
$lblCustomAttr.Location = New-Object System.Drawing.Point(310, $y)
$lblCustomAttr.Size = New-Object System.Drawing.Size(90, 20)
$lblCustomAttr.Text = "Wlasne pole:"
$lblCustomAttr.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($lblCustomAttr)
$txtCustomMatch = New-Object System.Windows.Forms.TextBox
$txtCustomMatch.Location = New-Object System.Drawing.Point(402, ([int]$y - 2))
$txtCustomMatch.Size = New-Object System.Drawing.Size(120, 22)
$txtCustomMatch.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($txtCustomMatch)
$btnAdvanced = New-Object System.Windows.Forms.Button
$btnAdvanced.Location = New-Object System.Drawing.Point(532, ([int]$y - 2))
$btnAdvanced.Size = New-Object System.Drawing.Size(160, 26)
$btnAdvanced.Text = "Ustawienia zaawansowane..."
$btnAdvanced.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($btnAdvanced)
$y = [int]$y + 28

$btnAdvanced.Add_Click({
    $adv = New-Object System.Windows.Forms.Form
    $adv.Text = "Ustawienia zaawansowane (lookupy)"
    $adv.Size = New-Object System.Drawing.Size(520, 340)
    $adv.StartPosition = "CenterParent"
    $adv.FormBorderStyle = "FixedDialog"
    $lblAdvDefault = New-Object System.Windows.Forms.Label
    $lblAdvDefault.Location = New-Object System.Drawing.Point(12, 12)
    $lblAdvDefault.Size = New-Object System.Drawing.Size(460, 32)
    $lblAdvDefault.Text = "Domyślny Id w celu (encja=guid, jedna linia na encje, np. businessunit=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)"
    $lblAdvDefault.AutoSize = $false
    $adv.Controls.Add($lblAdvDefault)
    $txtDefaultLookup = New-Object System.Windows.Forms.TextBox
    $txtDefaultLookup.Location = New-Object System.Drawing.Point(12, 48)
    $txtDefaultLookup.Size = New-Object System.Drawing.Size(478, 80)
    $txtDefaultLookup.Multiline = $true
    $txtDefaultLookup.ScrollBars = "Vertical"
    $txtDefaultLookup.Text = $script:GuiEntityDefaultTargetLookupStr
    $adv.Controls.Add($txtDefaultLookup)
    $lblAdvResolve = New-Object System.Windows.Forms.Label
    $lblAdvResolve.Location = New-Object System.Drawing.Point(12, 138)
    $lblAdvResolve.Size = New-Object System.Drawing.Size(460, 28)
    $lblAdvResolve.Text = "Dopasuj po nazwie (encje po przecinku, np. businessunit, uomschedule, uom, subject, transactioncurrency)"
    $lblAdvResolve.AutoSize = $false
    $adv.Controls.Add($lblAdvResolve)
    $txtResolveByName = New-Object System.Windows.Forms.TextBox
    $txtResolveByName.Location = New-Object System.Drawing.Point(12, 168)
    $txtResolveByName.Size = New-Object System.Drawing.Size(478, 24)
    $txtResolveByName.Text = ($script:GuiEntityLookupResolveByName -join ", ")
    $adv.Controls.Add($txtResolveByName)
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Location = New-Object System.Drawing.Point(312, 210)
    $btnOk.Size = New-Object System.Drawing.Size(86, 28)
    $btnOk.Text = "OK"
    $btnOk.DialogResult = "OK"
    $adv.Controls.Add($btnOk)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(404, 210)
    $btnCancel.Size = New-Object System.Drawing.Size(86, 28)
    $btnCancel.Text = "Anuluj"
    $btnCancel.DialogResult = "Cancel"
    $adv.Controls.Add($btnCancel)
    $adv.AcceptButton = $btnOk
    $adv.CancelButton = $btnCancel
    if ($adv.ShowDialog($form) -eq "OK") {
        $script:GuiEntityDefaultTargetLookupStr = $txtDefaultLookup.Text.Trim()
        $script:GuiEntityLookupResolveByName = @($txtResolveByName.Text -split "[,\s]+" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
})

$btnWhatIf = New-Object System.Windows.Forms.Button
$btnWhatIf.Location = New-Object System.Drawing.Point($pad, $y)
$btnWhatIf.Size = New-Object System.Drawing.Size(140, 32)
$btnWhatIf.Text = "Tylko podglad (WhatIf)"
$btnWhatIf.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($btnWhatIf)

$btnMigrate = New-Object System.Windows.Forms.Button
$btnMigrate.Location = New-Object System.Drawing.Point(150, $y)
$btnMigrate.Size = New-Object System.Drawing.Size(140, 32)
$btnMigrate.Text = "Uruchom migracje"
$btnMigrate.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnMigrate.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 204)
$btnMigrate.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnMigrate)

$btnLogs = New-Object System.Windows.Forms.Button
$btnLogs.Location = New-Object System.Drawing.Point(298, $y)
$btnLogs.Size = New-Object System.Drawing.Size(100, 32)
$btnLogs.Text = "Otworz logi"
$btnLogs.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($btnLogs)
$btnClearTarget = New-Object System.Windows.Forms.Button
$btnClearTarget.Location = New-Object System.Drawing.Point(404, $y)
$btnClearTarget.Size = New-Object System.Drawing.Size(120, 32)
$btnClearTarget.Text = "Wyczysc cel"
$btnClearTarget.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnClearTarget.BackColor = [System.Drawing.Color]::FromArgb(200, 80, 80)
$btnClearTarget.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnClearTarget)

$chkTestOneRecord = New-Object System.Windows.Forms.CheckBox
$chkTestOneRecord.Location = New-Object System.Drawing.Point -ArgumentList 532, ([int]$y + 6)
$chkTestOneRecord.Size = New-Object System.Drawing.Size(240, 24)
$chkTestOneRecord.Text = "Test: tylko 1 rekord na encje"
$chkTestOneRecord.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$chkTestOneRecord.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$form.Controls.Add($chkTestOneRecord)

$y = [int]$y + 40

$lblProgress = New-Object System.Windows.Forms.Label
$lblProgress.Location = New-Object System.Drawing.Point($pad, $y)
$lblProgress.Size = New-Object System.Drawing.Size(776, 22)
$lblProgress.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$lblProgress.Text = "Status: gotowy"
$lblProgress.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblProgress.BackColor = [System.Drawing.Color]::FromArgb(240, 248, 255)
$form.Controls.Add($lblProgress)
$y = [int]$y + 26

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point($pad, $y)
$txtLog.Size = New-Object System.Drawing.Size(776, 280)
$txtLog.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Bottom
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 8)
$txtLog.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
$form.Controls.Add($txtLog)

$script:ProgressEntity = ''
$script:ProgressEntityNum = 0
$script:ProgressEntityTotal = 0
$script:ProgressRecordCurrent = 0
$script:ProgressRecordTotal = 0

function Update-ProgressFromLogLine {
    param([string] $Line)
    if ([string]::IsNullOrWhiteSpace($Line)) { return }
    if ($Line -match 'Migracja encji\s*\((\d+)/(\d+)\)\s*:\s*(\S+)') {
        $script:ProgressEntityNum = [int]$Matches[1]
        $script:ProgressEntityTotal = [int]$Matches[2]
        $script:ProgressEntity = $Matches[3]
        $script:ProgressRecordCurrent = 0
        $script:ProgressRecordTotal = 0
    }
    if ($Line -match 'Rekordow do migracji\s*:\s*(\d+)') {
        $script:ProgressRecordTotal = [int]$Matches[1]
        $script:ProgressRecordCurrent = 0
    }
    if ($Line -match 'Postep\s*:\s*(\d+)/(\d+)') {
        $script:ProgressRecordCurrent = [int]$Matches[1]
        $script:ProgressRecordTotal = [int]$Matches[2]
    }
    if ($Line -match 'Migracja encji|Postep:|Rekordow do migracji|Zakonczono') {
        $remaining = $script:ProgressEntityTotal - $script:ProgressEntityNum
        $s = "Aktualnie: $($script:ProgressEntity) (encja $($script:ProgressEntityNum)/$($script:ProgressEntityTotal))"
        if ($script:ProgressRecordTotal -gt 0) {
            $s += " | Rekord $($script:ProgressRecordCurrent)/$($script:ProgressRecordTotal)"
        }
        $s += " | Pozostalo encji: $remaining"
        $lblProgress.Text = $s
    }
    if ($Line -match 'Migracja zakonczona|WhatIf zakonczony') {
        $lblProgress.Text = "Status: zakonczono"
    }
}

function Add-Log {
    param([string] $Message)
    Update-ProgressFromLogLine -Line $Message
    $txtLog.AppendText("$([DateTime]::Now.ToString('HH:mm:ss')) $Message`r`n")
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Read-LoginHasloFromFile {
    $path = Join-Path $script:Root 'Config\LoginHaslo.txt'
    $out = @{ Login = ''; Haslo = '' }
    if (-not (Test-Path $path)) { return $out }
    Get-Content -Path $path -Encoding UTF8 -ErrorAction SilentlyContinue | ForEach-Object {
        $line = $_.Trim()
        if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith('#')) { return }
        if ($line -match '^Login=(.*)$') { $out.Login = $Matches[1].Trim() }
        if ($line -match '^Haslo=(.*)$')  { $out.Haslo = $Matches[1].Trim() }
    }
    return $out
}

$btnConnectSource.Add_Click({
    $btnConnectSource.Enabled = $false
    try {
        if ($chkUseSavedSource.Checked) {
            $url = $txtUrlSource.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($url)) {
                [System.Windows.Forms.MessageBox]::Show("Wpisz URL srodowiska zrodla (np. https://org.crm4.dynamics.com).", "Brak URL", "OK", "Warning")
                $btnConnectSource.Enabled = $true
                return
            }
            $cred = Read-LoginHasloFromFile
            if ([string]::IsNullOrWhiteSpace($cred.Login) -or [string]::IsNullOrWhiteSpace($cred.Haslo)) {
                [System.Windows.Forms.MessageBox]::Show("W pliku Config\LoginHaslo.txt wypelnij Login= i Haslo= (format w pierwszej linii pliku).", "Brak login/haslo", "OK", "Warning")
                $btnConnectSource.Enabled = $true
                return
            }
            $connStr = New-DataverseConnectionString -Url $url -Username $cred.Login -Password $cred.Haslo
            Add-Log "Polaczenie ze zrodlem (login z pliku, URL z aplikacji)..."
            $script:SourceConn = Connect-DataverseEnvironment -ConnectionString $connStr
        } else {
            Add-Log "Otwieranie okna logowania do ZRODLA..."
            $script:SourceConn = Connect-DataverseEnvironment -Interactive
        }
        if (Test-DataverseConnection -Connection $script:SourceConn) {
            $lblStatusSource.Text = "polaczono"
            $lblStatusSource.ForeColor = [System.Drawing.Color]::Green
            Add-Log "Zrodlo: polaczono."
        } else {
            $lblStatusSource.Text = "blad testu"
            $lblStatusSource.ForeColor = [System.Drawing.Color]::Red
            $script:SourceConn = $null
        }
    } catch {
        Add-Log "Blad: $_"
        [System.Windows.Forms.MessageBox]::Show("$_", "Blad polaczenia ze zrodlem", "OK", "Error")
        $lblStatusSource.Text = "blad"
        $lblStatusSource.ForeColor = [System.Drawing.Color]::Red
        $script:SourceConn = $null
    } finally {
        $btnConnectSource.Enabled = $true
    }
})

$btnConnectTarget.Add_Click({
    $btnConnectTarget.Enabled = $false
    try {
        if ($chkUseSavedTarget.Checked) {
            $url = $txtUrlTarget.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($url)) {
                [System.Windows.Forms.MessageBox]::Show("Wpisz URL srodowiska celu (np. https://org2.crm4.dynamics.com).", "Brak URL", "OK", "Warning")
                $btnConnectTarget.Enabled = $true
                return
            }
            $cred = Read-LoginHasloFromFile
            if ([string]::IsNullOrWhiteSpace($cred.Login) -or [string]::IsNullOrWhiteSpace($cred.Haslo)) {
                [System.Windows.Forms.MessageBox]::Show("W pliku Config\LoginHaslo.txt wypelnij Login= i Haslo=.", "Brak login/haslo", "OK", "Warning")
                $btnConnectTarget.Enabled = $true
                return
            }
            $connStr = New-DataverseConnectionString -Url $url -Username $cred.Login -Password $cred.Haslo
            Add-Log "Polaczenie z celem (login z pliku, URL z aplikacji)..."
            $script:TargetConn = Connect-DataverseEnvironment -ConnectionString $connStr
        } else {
            Add-Log "Otwieranie okna logowania do CELU..."
            $script:TargetConn = Connect-DataverseEnvironment -Interactive
        }
        if (Test-DataverseConnection -Connection $script:TargetConn) {
            $lblStatusTarget.Text = "polaczono"
            $lblStatusTarget.ForeColor = [System.Drawing.Color]::Green
            Add-Log "Cel: polaczono."
        } else {
            $lblStatusTarget.Text = "blad testu"
            $lblStatusTarget.ForeColor = [System.Drawing.Color]::Red
            $script:TargetConn = $null
        }
    } catch {
        Add-Log "Blad: $_"
        [System.Windows.Forms.MessageBox]::Show("$_", "Blad polaczenia z celem", "OK", "Error")
        $lblStatusTarget.Text = "blad"
        $lblStatusTarget.ForeColor = [System.Drawing.Color]::Red
        $script:TargetConn = $null
    } finally {
        $btnConnectTarget.Enabled = $true
    }
})

$btnLoad.Add_Click({
    if (-not $script:SourceConn -or -not $script:TargetConn) {
        [System.Windows.Forms.MessageBox]::Show("Najpierw polacz ze zrodlem i z celem (przyciski u gory).", "Brak polaczen", "OK", "Warning")
        return
    }
    $txtLog.Clear()
    Add-Log "Pobieranie listy encji..."
    $btnLoad.Enabled = $false
    try {
        $sourceMeta = Get-EntityMetadataFromEnv -Connection $script:SourceConn
        $targetMeta = Get-EntityMetadataFromEnv -Connection $script:TargetConn
        $commonEntities = Get-CommonEntities -SourceMetadata $sourceMeta -TargetMetadata $targetMeta -ExcludeEntities $script:Config.SystemEntitiesToSkip
        $orderedEntities = Get-MigrationOrderByDependencies -CommonEntities $commonEntities -Config $script:Config -BpfSuffix $script:Config.BpfEntitySuffix
        if ($script:Config.EntityIncludeOnly -and $script:Config.EntityIncludeOnly.Count -gt 0) {
            $orderedEntities = Get-EntitiesFromWhitelistAndDependencies -CommonEntities $commonEntities -Whitelist $script:Config.EntityIncludeOnly -Config $script:Config -BpfSuffix $script:Config.BpfEntitySuffix
        }
        $script:EntityList = @($orderedEntities)
        $script:EntityRecordCounts = @{}
        $listEntities.Items.Clear()
        foreach ($e in $script:EntityList) {
            $bpf = if ($e -match 'process$') { " [BPF]" } else { "" }
            [void]$listEntities.Items.Add("$e$bpf", $false)
        }
        $btnCountRecords.Enabled = $true
        $numEnt = [int](@($script:EntityList).Count)
        Add-Log "Pobrano $numEnt encji. Opcjonalnie: kliknij Policz rekordy."
    } catch {
        Add-Log "Blad: $_"
        [System.Windows.Forms.MessageBox]::Show("Blad: $_", "Blad", "OK", "Error")
    } finally {
        $btnLoad.Enabled = $true
    }
})

$btnCountRecords.Add_Click({
    if (-not $script:SourceConn -or [int](@($script:EntityList).Count) -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Najpierw pobierz liste encji (polacz ze zrodlem i celem).", "Brak listy", "OK", "Warning")
        return
    }
    $btnCountRecords.Enabled = $false
    $btnLoad.Enabled = $false
    $checked = @()
    for ($i = 0; $i -lt $listEntities.Items.Count; $i++) { $checked += $listEntities.GetItemChecked($i) }
    Add-Log "Liczenie rekordow w zrodle (moze potrwac)..."
    $lblProgress.Text = "Status: liczenie rekordow..."
    try {
        $script:EntityRecordCounts = Get-EntityRecordCounts -Connection $script:SourceConn -EntityLogicalNames $script:EntityList -Logger { param($msg) Add-Log $msg }
        $listEntities.Items.Clear()
        for ($i = 0; $i -lt $script:EntityList.Count; $i++) {
            $e = $script:EntityList[$i]
            $bpf = if ($e -match 'process$') { " [BPF]" } else { "" }
            $cnt = if ($script:EntityRecordCounts.ContainsKey($e) -and $script:EntityRecordCounts[$e] -ge 0) { $script:EntityRecordCounts[$e] } else { "?" }
            $wasChecked = if ($i -lt $checked.Count) { $checked[$i] } else { $false }
            [void]$listEntities.Items.Add("$e ($cnt)$bpf", $wasChecked)
        }
        Add-Log "Wyswietlono liczby rekordow."
    } catch {
        Add-Log "Nie udalo sie policzyc rekordow: $_"
    }
    $lblProgress.Text = "Status: gotowy"
    $btnCountRecords.Enabled = $true
    $btnLoad.Enabled = $true
})

function Get-SelectedEntityFilter {
    $filter = @()
    $cnt = [int]$listEntities.Items.Count
    for ($i = 0; $i -lt $cnt; $i++) {
        if ($listEntities.GetItemChecked($i)) {
            $text = $listEntities.Items[$i] -replace ' \(\d+|\?\)$', '' -replace ' \[BPF\]$', ''
            $filter += $text
        }
    }
    return @($filter | Select-Object -Unique)
}

$chkOnlySelected.Add_CheckedChanged({
    if ($chkOnlySelected.Checked -and $chkExcludeSelected.Checked) { $chkExcludeSelected.Checked = $false }
})
$chkExcludeSelected.Add_CheckedChanged({
    if ($chkExcludeSelected.Checked -and $chkOnlySelected.Checked) { $chkOnlySelected.Checked = $false }
})

$btnWhatIf.Add_Click({
    if (-not $script:SourceConn -or -not $script:TargetConn) {
        [System.Windows.Forms.MessageBox]::Show("Najpierw polacz ze zrodlem i z celem.", "Brak polaczen", "OK", "Warning")
        return
    }
    $txtLog.Clear()
    $lblProgress.Text = "Status: WhatIf..."
    Add-Log "Uruchamiam WhatIf..."
    $btnWhatIf.Enabled = $false
    Import-Module Microsoft.Xrm.Data.PowerShell -Force -Scope Global -ErrorAction SilentlyContinue
    $filter = Get-SelectedEntityFilter
    try {
        $params = @{ SourceConn = $script:SourceConn; TargetConn = $script:TargetConn; WhatIf = $true; ConfigPath = $script:ConfigPath; MigrationMode = $comboMode.SelectedItem.ToString(); MatchBy = $comboMatchBy.SelectedItem.ToString() }
        if ($comboMatchBy.SelectedItem.ToString() -eq 'Custom') { $params['CustomMatchAttribute'] = $txtCustomMatch.Text.Trim() }
        if ($chkOnlyWithRecords.Checked) { $params['OnlyEntitiesWithRecordsAndDependencies'] = $true }
        if ($chkTestOneRecord.Checked) { $params['MaxRecordsPerEntity'] = 1 }
        if ($chkExcludeSelected.Checked -and [int](@($filter).Count) -gt 0) { $params['EntityExcludeFilter'] = $filter }
        elseif ($chkOnlySelected.Checked -and [int](@($filter).Count) -gt 0) { $params['EntityFilter'] = $filter }
        if ($script:GuiEntityDefaultTargetLookupStr) { $params['EntityDefaultTargetLookupStr'] = $script:GuiEntityDefaultTargetLookupStr }
        if ($script:GuiEntityLookupResolveByName -and $script:GuiEntityLookupResolveByName.Count -gt 0) { $params['EntityLookupResolveByName'] = $script:GuiEntityLookupResolveByName }
        & (Join-Path $script:Root 'Start-DataverseMigration.ps1') @params 2>&1 | ForEach-Object { Add-Log $_ }
        Add-Log "WhatIf zakonczony."
    } catch {
        Add-Log "Blad: $_"
        [System.Windows.Forms.MessageBox]::Show("Blad: $_", "Blad", "OK", "Error")
    } finally {
        $btnWhatIf.Enabled = $true
    }
})

$btnMigrate.Add_Click({
    if (-not $script:SourceConn -or -not $script:TargetConn) {
        [System.Windows.Forms.MessageBox]::Show("Najpierw polacz ze zrodlem i z celem.", "Brak polaczen", "OK", "Warning")
        return
    }
    $ok = [System.Windows.Forms.MessageBox]::Show("Uruchomic migracje danych do srodowiska docelowego?", "Potwierdzenie", "YesNo", "Question")
    if ($ok -ne "Yes") { return }
    $txtLog.Clear()
    $lblProgress.Text = "Status: uruchamiam migracje..."
    Add-Log "Uruchamiam migracje..."
    $btnMigrate.Enabled = $false
    Import-Module Microsoft.Xrm.Data.PowerShell -Force -Scope Global -ErrorAction SilentlyContinue
    $filter = Get-SelectedEntityFilter
    try {
        $params = @{ SourceConn = $script:SourceConn; TargetConn = $script:TargetConn; ConfigPath = $script:ConfigPath; MigrationMode = $comboMode.SelectedItem.ToString(); MatchBy = $comboMatchBy.SelectedItem.ToString() }
        if ($comboMatchBy.SelectedItem.ToString() -eq 'Custom') { $params['CustomMatchAttribute'] = $txtCustomMatch.Text.Trim() }
        if ($chkOnlyWithRecords.Checked) { $params['OnlyEntitiesWithRecordsAndDependencies'] = $true }
        if ($chkTestOneRecord.Checked) { $params['MaxRecordsPerEntity'] = 1 }
        if ($chkExcludeSelected.Checked -and [int](@($filter).Count) -gt 0) { $params['EntityExcludeFilter'] = $filter }
        elseif ($chkOnlySelected.Checked -and [int](@($filter).Count) -gt 0) { $params['EntityFilter'] = $filter }
        if ($script:GuiEntityDefaultTargetLookupStr) { $params['EntityDefaultTargetLookupStr'] = $script:GuiEntityDefaultTargetLookupStr }
        if ($script:GuiEntityLookupResolveByName -and $script:GuiEntityLookupResolveByName.Count -gt 0) { $params['EntityLookupResolveByName'] = $script:GuiEntityLookupResolveByName }
        & (Join-Path $script:Root 'Start-DataverseMigration.ps1') @params 2>&1 | ForEach-Object { Add-Log $_ }
        Add-Log "Migracja zakonczona."
        $lblProgress.Text = "Status: zakonczono"
        [System.Windows.Forms.MessageBox]::Show("Migracja zakonczona. Sprawdz log w folderze Logs.", "Info", "OK", "Information")
    } catch {
        Add-Log "Blad: $_"
        $lblProgress.Text = "Status: blad"
        [System.Windows.Forms.MessageBox]::Show("Blad: $_", "Blad", "OK", "Error")
    } finally {
        $btnMigrate.Enabled = $true
    }
})

$btnLogs.Add_Click({
    $logDir = Join-Path $script:Root 'Logs'
    if (Test-Path $logDir) { Start-Process explorer.exe -ArgumentList $logDir }
    else { [System.Windows.Forms.MessageBox]::Show('Folder Logs nie istnieje.', 'Info', 'OK', 'Information') }
})

function Get-EntitiesToClear {
    $allNames = @()
    $cnt = [int]$listEntities.Items.Count
    for ($i = 0; $i -lt $cnt; $i++) {
        $text = $listEntities.Items[$i] -replace ' \(\d+|\?\)$', '' -replace ' \[BPF\]$', ''
        $allNames += $text
    }
    $allNames = @($allNames | Select-Object -Unique)
    $filter = Get-SelectedEntityFilter
    if ($chkOnlySelected.Checked -and $filter.Count -gt 0) { return $filter }
    if ($chkExcludeSelected.Checked -and $filter.Count -gt 0) { return @($allNames | Where-Object { $_ -notin $filter }) }
    return $allNames
}

$btnClearTarget.Add_Click({
    if (-not $script:TargetConn) {
        [System.Windows.Forms.MessageBox]::Show("Najpierw polacz z celem.", "Brak polaczenia", "OK", "Warning")
        return
    }
    $toClear = Get-EntitiesToClear
    if (-not $toClear -or $toClear.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Brak encji do czyszczenia. Pobierz liste encji i zaznacz encje do wyczyszczenia (lub wylacz ""Migruj tylko zaznaczone"" / ""Wyklucz zaznaczone"" aby wybrac wszystkie).", "Info", "OK", "Information")
        return
    }
    $msg = "Czy na pewno usunac WSZYSTKIE rekordy w CELU z nastepujacych encji?`n`n" + ($toClear -join ", ") + "`n`nAby potwierdzic wpisz: TAK"
    $input = [System.Windows.Forms.MessageBox]::Show($msg, "Wyczysc cel - potwierdzenie", "OKCancel", "Warning")
    if ($input -ne "OK") { return }
    $confirmForm = New-Object System.Windows.Forms.Form
    $confirmForm.Text = "Potwierdzenie"
    $confirmForm.Size = New-Object System.Drawing.Size(360, 120)
    $confirmForm.StartPosition = "CenterParent"
    $confirmForm.FormBorderStyle = "FixedDialog"
    $lblConfirm = New-Object System.Windows.Forms.Label
    $lblConfirm.Text = "Wpisz TAK (wielkosc liter ma znaczenie) aby usunac rekordy:"
    $lblConfirm.Location = New-Object System.Drawing.Point(12, 12)
    $lblConfirm.Size = New-Object System.Drawing.Size(320, 20)
    $confirmForm.Controls.Add($lblConfirm)
    $txtConfirm = New-Object System.Windows.Forms.TextBox
    $txtConfirm.Location = New-Object System.Drawing.Point(12, 36)
    $txtConfirm.Size = New-Object System.Drawing.Size(320, 22)
    $confirmForm.Controls.Add($txtConfirm)
    $btnConfirmOk = New-Object System.Windows.Forms.Button
    $btnConfirmOk.Text = "OK"
    $btnConfirmOk.Location = New-Object System.Drawing.Point(168, 64)
    $btnConfirmOk.DialogResult = "OK"
    $confirmForm.AcceptButton = $btnConfirmOk
    $confirmForm.Controls.Add($btnConfirmOk)
    $btnConfirmCancel = New-Object System.Windows.Forms.Button
    $btnConfirmCancel.Text = "Anuluj"
    $btnConfirmCancel.Location = New-Object System.Drawing.Point(252, 64)
    $btnConfirmCancel.DialogResult = "Cancel"
    $confirmForm.CancelButton = $btnConfirmCancel
    $confirmForm.Controls.Add($btnConfirmCancel)
    if ($confirmForm.ShowDialog($form) -ne "OK") { return }
    if ($txtConfirm.Text.Trim() -ne "TAK") {
        Add-Log "Czyszczenie anulowane (brak wpisania TAK)."
        return
    }
    $txtLog.Clear()
    Add-Log "Rozpoczynam czyszczenie celu (encje: $($toClear.Count))..."
    $btnClearTarget.Enabled = $false
    try {
        $totalDeleted = 0
        foreach ($entityName in $toClear) {
            Add-Log "Czyszczenie encji: $entityName"
            $deleted = Clear-TargetEntityRecords -Conn $script:TargetConn -EntityLogicalName $entityName -Logger ${function:Add-Log}
            $totalDeleted += $deleted
        }
        Add-Log "Czyszczenie zakonczone. Usunieto lacznie rekordow: $totalDeleted"
        [System.Windows.Forms.MessageBox]::Show("Usunieto rekordow: $totalDeleted", "Czyszczenie celu", "OK", "Information")
    } catch {
        Add-Log "Blad: $_"
        [System.Windows.Forms.MessageBox]::Show("Blad: $_", "Blad czyszczenia", "OK", "Error")
    } finally {
        $btnClearTarget.Enabled = $true
    }
})

$form.Add_Shown({ $btnConnectSource.Focus() })
[void]$form.ShowDialog()

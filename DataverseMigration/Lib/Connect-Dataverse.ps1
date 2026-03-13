# Modul polaczen do Dataverse
# Wymaga: Microsoft.Xrm.Data.PowerShell

function New-DataverseConnectionString {
    <#
    .SYNOPSIS
        Buduje connection string OAuth z URL, loginu i hasla (do zapisanych danych w pliku).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string] $Url,
        [Parameter(Mandatory = $true)]
        [string] $Username,
        [Parameter(Mandatory = $true)]
        [string] $Password,
        [string] $AppId = '51f81489-12ee-4a9e-aaae-a2591f45987d',
        [string] $RedirectUri = 'app://58145B91-0C36-4500-8554-080854F2AC97'
    )
    $escape = { param($v) if ($v -match '[;=]') { return '"{0}"' -f ($v -replace '"', '""') }; return $v }
    $u = & $escape $Url.Trim()
    $l = & $escape $Username.Trim()
    $p = & $escape $Password
    return "AuthType=OAuth;Url=$u;Username=$l;Password=$p;AppId=$AppId;RedirectUri=$RedirectUri;LoginPrompt=Auto"
}

function Connect-DataverseEnvironment {
    <#
    .SYNOPSIS
        Nawiaze polaczenie ze srodowiskiem Dataverse/Dynamics 365.
    .PARAMETER ConnectionString
        Connection string (np. z aplikacji Azure AD / OAuth).
    .PARAMETER Interactive
        Jesli $true, uzywa logowania interaktywnego.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string] $ConnectionString,
        [Parameter(Mandatory = $false)]
        [switch] $Interactive
    )

    if (-not (Get-Module -ListAvailable -Name 'Microsoft.Xrm.Data.PowerShell')) {
        throw "Modul Microsoft.Xrm.Data.PowerShell nie jest zainstalowany. Zainstaluj: Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser"
    }

    Import-Module Microsoft.Xrm.Data.PowerShell -Force -ErrorAction Stop

    if ($Interactive) {
        $conn = Get-CrmConnection -InteractiveMode
        return $conn
    }

    if ([string]::IsNullOrWhiteSpace($ConnectionString)) {
        throw "Podaj ConnectionString lub uzyj -Interactive."
    }

    try {
        $conn = Get-CrmConnection -ConnectionString $ConnectionString
    } catch {
        $msg = $_.Exception.Message
        if ($_.Exception.InnerException) {
            $msg = $_.Exception.InnerException.Message
            if ($_.Exception.InnerException.InnerException) {
                $msg = $_.Exception.InnerException.InnerException.Message
            }
        }
        $full = $_.Exception.ToString()
        if ($msg -match "AADSTS|MFA|multi.factor|interactive|consent") {
            $msg = $msg + " (Tip: Uzyj logowania interaktywnego w konsoli PowerShell: .\Run-InteractiveMigration.ps1)"
        }
        if ($msg -match "Seamless|intranet|on premises|on-premises") {
            $msg = $msg + " Logowanie login+haslo wymaga Seamless SSO (siec firmowa / AD). Uzyj logowania interaktywnego: uruchom z konsoli PowerShell (nie z exe) skrypt .\Run-InteractiveMigration.ps1 - otworzy sie przegladarka do logowania."
        }
        throw "Get-CrmConnection nie powiodl sie: $msg"
    }
    if (-not $conn) {
        throw "Nie udalo sie nawiazac polaczenia z Dataverse."
    }
    return $conn
}

function Test-DataverseConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $Connection
    )
    try {
        # Encja organization ma jeden rekord; nie wymaga Id (Get-CrmRecord "whoami" w tym module pytal o Id).
        $null = Get-CrmRecords -Conn $Connection -EntityLogicalName "organization" -TopCount 1 -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

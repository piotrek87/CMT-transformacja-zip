# Konfiguracja migracji CMT (Configuration Migration Tool)
# Polaczenia: zmienne srodowiskowe LUB plik Config\Polaczenia.txt (ZrodloUrl, ZrodloLogin, ZrodloHaslo, CelUrl, CelLogin, CelHaslo)

function Build-CMTConnectionString {
    param([string]$Url, [string]$Username, [string]$Password)
    if ([string]::IsNullOrWhiteSpace($Url) -or [string]::IsNullOrWhiteSpace($Username)) { return '' }
    $escape = { param($v) if ($v -match '[;=]') { return '"{0}"' -f ($v -replace '"', '""') }; return $v }
    $u = & $escape $Url.Trim()
    $l = & $escape $Username.Trim()
    $p = & $escape ([string]$Password)
    $appId = '51f81489-12ee-4a9e-aaae-a2591f45987d'
    $redirect = 'app://58145B91-0C36-4500-8554-080854F2AC97'
    return "AuthType=OAuth;Url=$u;Username=$l;Password=$p;AppId=$appId;RedirectUri=$redirect;LoginPrompt=Auto"
}

function Read-PolaczeniaFromFile {
    param([string]$FilePath)
    $h = @{}
    if (-not (Test-Path $FilePath)) { return $h }
    foreach ($line in [System.IO.File]::ReadAllLines($FilePath)) {
        $s = $line.Trim()
        if ([string]::IsNullOrEmpty($s) -or $s.StartsWith('#')) { continue }
        $idx = $s.IndexOf('=')
        if ($idx -le 0) { continue }
        $key = $s.Substring(0, $idx).Trim()
        $val = $s.Substring($idx + 1).Trim()
        $h[$key] = $val
    }
    return $h
}

$srcConn = if ($env:DATAVERSE_SOURCE_CONNECTION) { $env:DATAVERSE_SOURCE_CONNECTION } else { '' }
$tgtConn = if ($env:DATAVERSE_TARGET_CONNECTION) { $env:DATAVERSE_TARGET_CONNECTION } else { '' }

if ([string]::IsNullOrWhiteSpace($srcConn) -or [string]::IsNullOrWhiteSpace($tgtConn)) {
    $polPath = Join-Path $PSScriptRoot 'Polaczenia.txt'
    $pol = Read-PolaczeniaFromFile -FilePath $polPath
    if ($pol.Count -gt 0) {
        $zUrl = $pol['ZrodloUrl']; $zLogin = $pol['ZrodloLogin']; $zHaslo = $pol['ZrodloHaslo']
        if ([string]::IsNullOrWhiteSpace($zLogin)) { $zLogin = $pol['Login']; $zHaslo = $pol['Haslo'] }
        if ([string]::IsNullOrWhiteSpace($srcConn) -and -not [string]::IsNullOrWhiteSpace($zUrl)) {
            $srcConn = Build-CMTConnectionString -Url $zUrl -Username $zLogin -Password $zHaslo
        }
        $cUrl = $pol['CelUrl']; $cLogin = $pol['CelLogin']; $cHaslo = $pol['CelHaslo']
        if ([string]::IsNullOrWhiteSpace($cLogin)) { $cLogin = $pol['Login']; $cHaslo = $pol['Haslo'] }
        if ([string]::IsNullOrWhiteSpace($tgtConn) -and -not [string]::IsNullOrWhiteSpace($cUrl)) {
            $tgtConn = Build-CMTConnectionString -Url $cUrl -Username $cLogin -Password $cHaslo
        }
    }
}

# Encje w schemacie CMT: tylko te (kontakty, konta, leady, opportunity, dzialania, uwagi + zaleznosci + BPF).
# Pusta tablica = wszystkie encje ze zrodla. Wypelniona = tylko te z listy (kolejnosc w SchemaEntityOrder).
$SchemaEntityIncludeOnly = @(
    'transactioncurrency', 'businessunit', 'subject',
    'systemuser', 'team', 'role',
    'account', 'contact', 'lead', 'opportunity',
    'activitypointer', 'email', 'task', 'appointment', 'phonecall', 'letter', 'fax', 'annotation',
    'workflow', 'processstage', 'opportunitysalesprocess', 'leadtoopportunitysalesprocess',
    'product', 'pricelevel', 'quote', 'salesorder', 'invoice',
    'opportunityproduct', 'quotedetail', 'salesorderdetail', 'invoicedetail',
    'campaign', 'campaignactivity', 'campaignresponse', 'list', 'listmember',
    'uom', 'uomschedule'
)
# Encje systemowe do pominiecia przy porownaniu schematow (opcjonalnie)
$SystemEntitiesToSkip = @()

# Kolejnosc encji w schemacie (zaleznosci pierwsze – zeby import sie spinál)
$SchemaEntityOrder = @(
    'transactioncurrency', 'businessunit', 'subject', 'uom', 'uomschedule',
    'systemuser', 'team', 'role',
    'account', 'contact', 'lead', 'opportunity',
    'product', 'pricelevel', 'quote', 'salesorder', 'invoice',
    'opportunityproduct', 'quotedetail', 'salesorderdetail', 'invoicedetail',
    'activitypointer', 'email', 'task', 'appointment', 'phonecall', 'letter', 'fax', 'annotation',
    'workflow', 'processstage', 'opportunitysalesprocess', 'leadtoopportunitysalesprocess',
    'campaign', 'campaignactivity', 'campaignresponse', 'list', 'listmember'
)

# Encje do calkowitego pominiecia przy imporcie (usuwane z data.xml i data_schema.xml).
# Np. salesliteratureitem gdy cel ma inna wersje encji (CMT: Missing Fields) i nie migrujesz tej tabeli.
$EntitiesToExcludeFromImport = @('salesliteratureitem')

# Pola lookup do usuniecia z zipa przed importem (encje w zrodle maja GUID, w celu brak – CMT failuje Stage).
# Usuniecie z XML = import nie probuje lookup; w celu pole bedzie puste.
$LookupFieldsToStripFromImport = @(
    'msdyn_accountkpiid'
    'msdyn_contactkpiid'
    'transactioncurrencyid'
    'originatingleadid'
)

# Walidacja option setow (branza, statusy itd.) przy transformacji zipa:
# - Report     = wykryj niepasujace, zapisz CSV; po transformacji (opcja 3) pokaz bledy i pozwol wybrac
#               opcje ZE ZRODLA (numer + nazwa) na ktora zamienic – wymaga Polaczenia.txt (Zrodlo + Cel)
# - Clear     = usun wartosc pola (rekord zaimportuje sie z pusta branza)
# - Replace   = ustaw na wartosc z OptionSetFallbackValues (np. industrycode=82 jako "Inne")
# - Interactive = przy kazdej niepasujacej wartosci pytaj (opcje ze zrodla lub z celu), wybierz zamienic / wyczysc / pomin
# OptionSetFallbackValues: uzywane gdy Replace (nazwa_pola -> wartosc int w celu, np. industrycode = 82)
$OptionSetValidationAction = 'Interactive'
$OptionSetFallbackValues   = @{ industrycode = 82 }

@{
    SourceConnectionString    = $srcConn
    TargetConnectionString     = $tgtConn
    SchemaFilePath            = ''
    ExportOutputDirectory     = Join-Path $PSScriptRoot '..\Output'
    ImportDataFileName        = 'CMT_Export.zip'
    PreserveCreatedOn         = $true
    PreserveOwner             = $true
    UserMapFilePath           = ''
    DisableTelemetry          = $true
    SchemaEntityIncludeOnly   = $SchemaEntityIncludeOnly
    SchemaEntityOrder         = $SchemaEntityOrder
    SystemEntitiesToSkip      = $SystemEntitiesToSkip
    OptionSetValidationAction     = $OptionSetValidationAction
    OptionSetFallbackValues       = $OptionSetFallbackValues
    LookupFieldsToStripFromImport = $LookupFieldsToStripFromImport
    EntitiesToExcludeFromImport  = $EntitiesToExcludeFromImport
}

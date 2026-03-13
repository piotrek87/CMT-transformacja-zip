# Konfiguracja migracji Dataverse
# Ścieżka: DataverseMigration\Config\MigrationConfig.ps1
# Zgodne z PowerShell 5.1 (operator ?? jest dopiero w PS 7)

$srcConn = if ($env:DATAVERSE_SOURCE_CONNECTION) { $env:DATAVERSE_SOURCE_CONNECTION } else { '' }
$tgtConn = if ($env:DATAVERSE_TARGET_CONNECTION) { $env:DATAVERSE_TARGET_CONNECTION } else { '' }

@{
    # Parametry połączenia - można nadpisać przez parametry skryptu
    SourceConnectionString = $srcConn
    TargetConnectionString = $tgtConn

    # Limity i batch
    PageSize = 5000
    BatchSize = 100
    MaxRetryCount = 3
    RetryDelaySeconds = 5

    # Pola systemowe do pominięcia przy migracji
    # statecode/statuscode: w Org Service nie ustawia się przez Update() – wymagany SetStateRequest (osobne wywołanie)
    SystemFieldsToSkip = @(
        'createdon', 'modifiedon', 'createdby', 'modifiedby',
        'versionnumber', 'utcconversiontimezonecode', 'timezoneruleversionnumber',
        'importsequencenumber', 'overriddencreatedon',  # overriddencreatedon przekazujemy osobno
        'statecode', 'statuscode'
    )

    # Encje systemowe/konfiguracyjne – nie migrowane i nie pokazywane w spisie (cel: tylko dane biznesowe, uzytkownik pracuje tak samo)
    SystemEntitiesToSkip = @(
        'bulkdeleteoperation', 'asyncoperation', 'workflow', 'workflowdependency', 'pluginassembly', 'plugintype',
        'sdkmessage', 'sdkmessagefilter', 'sdkmessageprocessingstep', 'sdkmessageprocessingstepimage',
        'systemform', 'savedquery', 'userquery', 'ribboncustomization',
        'multientitysearch', 'exportsolutionupload', 'dvtablesearchentity',
        'subscription', 'subscriptionstatisticsoffline', 'subscriptionsyncentryoffline',
        'importjob', 'importfile', 'importmap', 'importdata',
        'duplicaterule', 'duplicaterulecondition', 'duplicaterecord',
        'sharepointsite', 'sharepointdocumentlocation', 'sharepointdocument',
        'solutioncomponent', 'solutioncomponentdefinition',
        'callbackregistration', 'serviceendpoint', 'serviceplanmapping',
        'channelproperty', 'channelpropertygroup', 'entityanalyticsconfig',
        'businessdatalocalizedlabel', 'managedproperty', 'localizedlabel', 'stringmap',
        'msdyn_aioptimizationprivatedata', 'appnotificationsignal', 'serviceplancustomcontrol'
    )

    # Tylko te encje (plus zaleznosci do relacji) – pusta tablica = migruj wszystkie wspolne
    # Wypelniona = migruj tylko: kontakty, klienci, leady, szanse, oferty, zamowienia, faktury, aktywnosci, adnotacje, produkty, cenniki itd.
    EntityIncludeOnly = @(
        'systemuser', 'team', 'role', 'businessunit', 'transactioncurrency', 'subject',
        'uom', 'uomschedule', 'unit',
        'account', 'contact', 'lead', 'opportunity', 'quote', 'salesorder', 'invoice',
        'product', 'pricelevel', 'opportunityproduct', 'quotedetail', 'salesorderdetail', 'invoicedetail',
        'activitypointer', 'email', 'task', 'appointment', 'phonecall', 'letter', 'fax', 'annotation',
        'campaign', 'campaignactivity', 'campaignresponse', 'list', 'listmember',
        'opportunitysalesprocess', 'leadtoopportunitysalesprocess'
    )

    # Kolejność encji (prefixy/zależności) - encje wcześniejsze są migrowane pierwsze
    EntityOrderPriority = @(
        'transactioncurrency', 'businessunit', 'subject', 'uom', 'uomschedule',
        'systemuser', 'team', 'role',
        'account', 'contact', 'lead',
        'invoice', 'opportunity', 'quote', 'salesorder', 'product', 'pricelevel',
        'invoicedetail', 'quotedetail', 'salesorderdetail', 'opportunityproduct',
        'activitypointer', 'email', 'task', 'appointment', 'phonecall', 'letter', 'fax', 'annotation',
        'campaign', 'campaignactivity', 'campaignresponse', 'list', 'listmember'
    )

    # Jawne zależności (gdy metadane nie zwracają Target dla lookup) – encja => encje, które muszą być wcześniej
    EntityDependencyOverrides = @{
        role              = @('businessunit')
        invoicedetail     = @('invoice')
        quotedetail       = @('quote', 'product')
        salesorderdetail  = @('salesorder', 'product')
        opportunityproduct = @('opportunity', 'product')
        product           = @('uomschedule')
    }

    # Sufiks encji BPF
    BpfEntitySuffix = 'process'

    # Ponowne przebiegi: jesli po pierwszym przejsciu sa bledy (np. brakujacy lead/klient), uruchom jeszcze 1–2 przebiegi – wtedy IdMap jest juz wypelniony i relacje moga sie dopasowac.
    RetryFailedRecordPasses = 2

    # Tryb migracji: Create (tylko nowe), Update (tylko aktualizacja istniejacych), Upsert (create lub update)
    MigrationMode = 'Upsert'

    # Dopasowanie rekordu w celu: Id (mapa z poprzedniej migracji), IdThenName (najpierw Id potem nazwa), Name (tylko pole), Custom (wlasne pole z CustomMatchAttribute)
    MatchBy = 'IdThenName'

    # Sciezka do pliku mapy ID (source->target). Domyslnie Logs\IdMap_latest.json
    IdMapPath = ''

    # Dla MatchBy Custom: nazwa atrybutu do dopasowania (np. accountnumber). Dla Name/IdThenName uzyj EntityMatchKey.
    CustomMatchAttribute = ''

    # Klucz dopasowania po nazwie (dla Name / IdThenName). Encja -> nazwa atrybutu unikalnego. Pomin encje bez danego atrybutu w srodowisku.
    EntityMatchKey = @{
        account      = 'accountnumber'
        contact      = 'emailaddress1'
        lead         = 'emailaddress1'
        opportunity  = 'name'
        systemuser   = 'fullname'
        team         = 'name'
    }

    # Istniejace rekordy w celu – dolaczanie zamiast tworzenia (gdy brak w IdMap):
    # Domyślny Id w celu dla danej encji (jeden rekord = np. jedna BU, jeden uomschedule). Encja -> Guid (string).
    EntityDefaultTargetLookup = @{
        # businessunit  = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
        # uomschedule   = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
        # uom           = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
        # transactioncurrency = 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
    }
    # Dla tych encji przy braku mapy probuj znalezc rekord w celu po nazwie (pobierz nazwe ze zrodla, wyszukaj w celu). Wymaga EntityMatchKey lub atrybutu 'name'.
    EntityLookupResolveByName = @(
        'businessunit', 'uomschedule', 'uom', 'subject', 'transactioncurrency', 'systemuser'
    )

    # Logowanie
    LogFolder = '.\Logs'
    LogFileName = "Migration_{0:yyyyMMdd_HHmmss}.log"
}

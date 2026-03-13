# Narzędzie migracyjne Dataverse (PowerShell)

Skrypt do przenoszenia danych biznesowych między środowiskami **Microsoft Dataverse / Dynamics 365 Sales** z użyciem modułu **Microsoft.Xrm.Data.PowerShell**.

## Wymagania

- PowerShell 5.1 lub PowerShell 7+
- Moduł: **Microsoft.Xrm.Data.PowerShell**

```powershell
Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser
```

- Połączenie ze środowiskiem źródłowym i docelowym (connection string lub logowanie interaktywne).

## Jak używać – interfejsy

### Uruchomienie jednym kliknięciem

- **MigracjaDataverse.bat** – dwuklik uruchamia GUI (PowerShell w tle). Wymaga, aby cały folder `DataverseMigration` (z podfolderami `Lib`, `Config`) był na miejscu.
- **MigracjaDataverse.exe** – po zbudowaniu (patrz niżej) dwuklik uruchamia to samo okno. Exe musi stać w folderze razem z `Lib`, `Config` i plikami `.ps1`.

**Budowanie .exe (opcjonalnie):**

```powershell
cd DataverseMigration
Install-Module ps2exe -Scope CurrentUser   # jednorazowo
.\Build-Exe.ps1
```

Powstanie plik `MigracjaDataverse.exe`. Dystrybuuj cały folder (exe + `Lib`, `Config`, `Start-MigrationGUI.ps1`, `Get-MigrationEntityList.ps1`, `Start-DataverseMigration.ps1` itd.) – exe uruchamia skrypty z tego samego katalogu.

### 1. Interfejs graficzny (GUI) – zalecany na start

Okno z polami na connection stringi, listą encji i przyciskami. Możesz je uruchomić przez **dwuklik `MigracjaDataverse.bat`** albo:

```powershell
cd DataverseMigration
.\Start-MigrationGUI.ps1
```

- Wklej **connection string źródła** i **celu** w pola tekstowe.
- Kliknij **„Pobierz listę encji”** – narzędzie połączy się z oboma środowiskami i wypełni listę wspólnych encji.
- Opcjonalnie zaznacz **„Migruj tylko zaznaczone encje”** i zaznacz wybrane encje na liście.
- **„Tylko podgląd (WhatIf)”** – sprawdza połączenia i pokazuje, co zostałoby zmigrowane, bez kopiowania.
- **„Uruchom migrację”** – uruchamia faktyczną migrację danych.
- **„Otwórz logi”** – otwiera folder `Logs` z plikami logów.

**Zapisane logowanie (bez logowania za każdym razem):** Otwórz plik tekstowy **`Config\LoginHaslo.txt`** (szablon jest w repozytorium). Wpisz w kolejnych liniach `Login=twoj@email.com` i `Haslo=twojehaslo`, zapisz. W aplikacji zaznacz **„Użyj zapisanych login/hasło”** przy Źródle i/lub Celu, **wpisz URL środowiska** (np. `https://org.crm4.dynamics.com`) w polu URL i kliknij „Połącz…”. Środowisko (URL) wybierasz w aplikacji; login i hasło brane są z pliku. Jedna para login/hasło wystarczy dla obu połączeń. Nie commituj pliku z hasłem do repozytorium.

W GUI mozesz tez **logowac sie interaktywnie** (okno Dynamics) – wtedy nie zaznaczaj „Użyj zapisanych”. Logowanie login+hasło z pliku wymaga zwykle sieci firmowej / Seamless SSO. Jesli pojawi sie blad **„Seamless single sign on failed”** lub **„No access on premises AD or intranet zone”**, logowanie login+haslo nie dziala w Twojej sieci (wymaga Seamless SSO / AD). Wtedy **uruchom migracje z logowaniem interaktywnym** (przegladarka):

```powershell
cd DataverseMigration
.\Run-InteractiveMigration.ps1
```

Otworzy sie dwukrotnie przegladarka (zrodlo, potem cel). Po zalogowaniu migracja ruszy. Podglad bez migracji: `.\Run-InteractiveMigration.ps1 -WhatIf`.

### 2. Menu konsolowe (krok po kroku)

```powershell
cd DataverseMigration
.\Start-MigrationMenu.ps1
```

W menu wybierasz kolejno:

1. **Ustaw połączenia** – wpisujesz lub wklejasz connection stringi (albo ścieżkę do pliku .txt).
2. **Pobierz listę encji** – ładuje listę encji do migracji.
3. **Tylko podgląd (WhatIf)** – podgląd bez migracji.
4. **Uruchom migrację (wszystkie encje)**.
5. **Uruchom migrację (wybrane encje)** – podaj numery lub nazwy encji (np. `3,5,7` lub `account,contact,opportunity`).
6. **Otwórz folder z logami**.

### 3. Wiersz poleceń (parametry)

Dla automatyzacji lub jednego uruchomienia bez menu:

```powershell
.\Start-DataverseMigration.ps1 -SourceConnectionString $src -TargetConnectionString $tgt
.\Start-DataverseMigration.ps1 -SourceConnectionString $src -TargetConnectionString $tgt -WhatIf
.\Start-DataverseMigration.ps1 -SourceConnectionString $src -TargetConnectionString $tgt -EntityFilter 'account','contact'
```

## Architektura

```
DataverseMigration/
├── Start-DataverseMigration.ps1   # Skrypt główny (entry point)
├── Start-MigrationGUI.ps1         # Interfejs graficzny (Windows Forms)
├── Start-MigrationMenu.ps1       # Interaktywne menu konsolowe
├── Get-MigrationEntityList.ps1   # Helper: zwraca listę encji (dla GUI/menu)
├── Config/
│   └── MigrationConfig.ps1       # Konfiguracja (stronnicowanie, retry, pola systemowe, kolejność encji)
└── Lib/
    ├── Connect-Dataverse.ps1     # Połączenie z Dataverse (Get-CrmConnection)
    ├── Get-EntityMetadata.ps1    # Metadane encji, wykrywanie wspólnych encji
    ├── Get-MigrationOrder.ps1    # Kolejność migracji (zależności + BPF na końcu)
    └── Migrate-EntityData.ps1    # Migracja rekordów (paging, retry, mapowanie lookup)
```

### Przepływ

1. **Połączenie** – źródło i cel (connection string lub `-Interactive`).
2. **Metadane** – pobranie listy encji i atrybutów z obu środowisk.
3. **Wspólne encje** – migrowane są tylko encje istniejące w obu środowiskach, z uwzględnieniem wspólnych atrybutów.
4. **Kolejność** – sortowanie topologiczne po zależnościach (lookup); encje BPF (sufiks `process`) na końcu.
5. **Migracja** – dla każdej encji: odczyt ze źródła (FetchXML + paging 5000), mapowanie pól, zapis do celu z retry.

### Zachowanie danych

- **CreatedOn** – ustawiane przez pole `overriddencreatedon` przy tworzeniu rekordu.
- **Relacje** – pola EntityReference (lookup) mapowane przez słownik źródłowy GUID → docelowy GUID (po migracji encji referencowanej).
- **ownerid** – zachowany, jeśli użytkownik/zespół został wcześniej zmigrowany (np. `systemuser`, `team`).
- **statecode / statuscode** – kopiowane jako OptionSetValue.
- **GUID** – w standardowym tworzeniu przez `Add-CrmRecord` Dataverse generuje nowe GUID; mapowanie stary→nowy służy do ustawiania lookupów.

### Business Process Flow (BPF)

- Encje z logiczna nazwą kończącą się na **`process`** traktowane są jako tabele BPF.
- Są migrowane **po** encjach głównych (np. account, contact, opportunity), aby lookupi do tych encji były już dostępne.
- Aktywny etap procesu jest zachowany przez kopiowanie atrybutów encji (w tym pól stanu/stage).

### Pomijane pola systemowe

Przy kopiowaniu atrybutów pomijane są m.in.:  
`createdon`, `modifiedon`, `createdby`, `modifiedby`, `versionnumber`,  
`utcconversiontimezonecode`, `timezoneruleversionnumber`, `importsequencenumber`.

### Stronicowanie i retry

- Odczyt ze źródła: **FetchXML** z `count` i `paging-cookie` (do 5000 rekordów na stronę).
- Zapis: po jednym rekordzie z **retry** (domyślnie 3 próby, opóźnienie 5 s).
- Logowanie postępu do konsoli i do pliku w `.\Logs\`.

## Użycie z wiersza poleceń

### Connection string (źródło i cel)

```powershell
$source = "AuthType=OAuth;Url=https://org.crm4.dynamics.com;..."
$target = "AuthType=OAuth;Url=https://org2.crm4.dynamics.com;..."
.\Start-DataverseMigration.ps1 -SourceConnectionString $source -TargetConnectionString $target
```

### Tryb interaktywny (jedno środowisko – do testów)

```powershell
.\Start-DataverseMigration.ps1 -Interactive
```

### Tylko wybrane encje

```powershell
.\Start-DataverseMigration.ps1 -SourceConnectionString $source -TargetConnectionString $target -EntityFilter 'account','contact','opportunity'
```

### Podgląd bez migracji (WhatIf)

```powershell
.\Start-DataverseMigration.ps1 -SourceConnectionString $source -TargetConnectionString $target -WhatIf
```

### Konfiguracja w pliku / zmienne środowiskowe

- Edycja `Config\MigrationConfig.ps1`: connection stringi, `PageSize`, `BatchSize`, `MaxRetryCount`, `SystemEntitiesToSkip`, `EntityOrderPriority`.
- Opcjonalnie zmienne środowiskowe: `$env:DATAVERSE_SOURCE_CONNECTION`, `$env:DATAVERSE_TARGET_CONNECTION`.

## Uwagi

- **Metadane** – skrypt używa `Get-CrmEntityAllMetadata` (jeśli dostępne) lub pobiera metadane pojedynczo dla wybranych encji; w zależności od wersji modułu może być potrzebna drobna korekta nazw cmdletów.
- **GUID** – zachowanie oryginalnych GUID wymaga zapisu przez API z przekazaniem ID (np. Web API PUT); obecna implementacja tworzy rekordy przez `Add-CrmRecord`, więc GUID się zmieniają, a relacje są utrzymywane przez mapowanie.
- **Batch create** – moduł może nie udostępniać batch create; rekordy są zapisywane pojedynczo z retry.

## Porównanie z Data Transporter (XrmToolBox)

Kod źródłowy **Data Transporter** (plugin do XrmToolBox):  
**[github.com/bcolpaert/Colso.Xrm.DataTransporter](https://github.com/bcolpaert/Colso.Xrm.DataTransporter)**

Główne pliki z logiką transferu:
- **DataTransporter.cs** – formularz, wywołanie transferu, playlisty.
- **AppCode/EntityRecord.cs** – pobieranie rekordów (FetchXML), Create/Update/Delete, mapowania lookupów, state/status.
- **AppCode/CrmExceptionHelper.cs** – obsługa błędów (FaultException).

Różnice w podejściu:

| Aspekt | Data Transporter (XrmToolBox) | To narzędzie (PowerShell) |
|--------|-------------------------------|----------------------------|
| **Dopasowanie rekordów** | Po **Id** – zakłada, że w celu mogą być rekordy z tym samym GUID co w źródle (np. restore backup). | Po **Id** (mapa z poprzedniej migracji), **IdThenName** lub **Name** – dopasowanie po polu (np. email) i mapowanie source→target Id. |
| **Lookupy** | Lista **mapowań** (source EntityReference → target EntityReference). Użytkownik konfiguruje lub „Auto Mappings”. ApplyMappings zamienia w encji wszystkie pasujące referencje. | **IdMap** (source GUID → target GUID po migracji) + **ResolveLookupId** (np. po nazwie). Lookupy bez mapy można pominąć lub dołączyć do rekordu w celu. |
| **Atrybuty** | Użytkownik **zaznacza atrybuty** w GUI. Wysyłane jest tylko to, co zaznaczone; typy z metadanych. | Wspólne atrybuty z metadanych lub **odkrywanie pól** z pierwszego rekordu (gdy metadane puste). Retry: po błędzie API (OptionSet/Boolean/Does Not Exist) usuwanie problematycznego pola i ponowienie. |
| **State/Status** | **SetStateRequest** osobno po Create/Update (statecode/statuscode usuwane z encji przed wysłaniem). | statecode/statuscode w **SystemFieldsToSkip** – nie wysyłane w Create/Update (SetState można dodać osobno). |
| **Błędy** | Zbierane w listę (Messages), ExecuteMultiple z ContinueOnError. Bez retry „usuń pole i spróbuj again”. | Retry z usuwaniem/konwersją atrybutu (OptionSet, Boolean, „doesn't contain attribute”, „Does Not Exist”), logowanie każdego błędu. |

Wnioski dla naszego skryptu: Data Transporter opiera się na jawnym wyborze atrybutów i ręcznych mapowaniach lookupów; nasze narzędzie idzie w stronę automatycznego mapowania Id i konwersji „w locie” po błędach API. Warto rozważyć w GUI **wybór atrybutów** (jak w Data Transporter) oraz **eksport/import mapowań** (IdMap już jest).

## Licencja

Do użytku wewnętrznego. Użycie modułu Microsoft.Xrm.Data.PowerShell podlega jego licencji.

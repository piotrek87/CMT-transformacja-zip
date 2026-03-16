# DataverseMigrationCMT

**Repozytorium:** [https://github.com/piotrek87/CMT-transformacja-zip](https://github.com/piotrek87/CMT-transformacja-zip)

Aplikacja **modyfikuje zip** wyeksportowany z CMT: podmienia właścicieli (IdMap), ustawia oryginalne daty utworzenia (overriddencreatedon), usuwa pola nieistniejące w celu. **Importu nie robi** – gotowy zip (`*_ForTarget.zip`) importujesz sam (CMT Import data lub inne narzędzie).

## Przepływ (CMT Export → 2 → 3 → Ty importujesz)

| Krok | Opis |
|------|------|
| **Ty** | W **DataMigrationUtility**: Export data → źródło → wybierz encje → zapisz zip np. do `Input\` lub `Output\CMT_Export.zip`. |
| **2** | W menu: **Generuj User Map** – mapowanie użytkowników (imię+nazwisko) → User Map XML + IdMap JSON. |
| **3** | **Transformuj zip** – podmiana ownerid, overriddencreatedon, usunięcie pól nie z celu → `Output\..._ForTarget.zip`. |
| **Ty** | **Import** – w CMT (Import data) lub innym narzędziu wybierasz plik `*_ForTarget.zip` i wykonujesz import. |

## Różnica względem DataverseMigration

| | **DataverseMigration** | **DataverseMigrationCMT** |
|---|------------------------|----------------------------|
| Schemat | Metadane źródło + cel | **Porównanie źródło vs cel** → jeden schemat importu (tylko pola z celu). |
| Dane | Własna logika (FetchXML, IdMap) | CMT Export → zip → **transformacja zipa** → Ty importujesz. |
| Właściciel | Mapowanie po fullname w skrypcie | User Map + IdMap; transformacja podmienia ownerid w zipie. |
| Daty | overriddencreatedon w API | Transformacja ustawia overriddencreatedon w zipie. |

## Wymagania

- PowerShell 5.1+
- Moduł **Microsoft.Xrm.Tooling.ConfigurationMigration** (1.0.0.88+) – opcjonalnie (eksport/import CMT)
- Moduł **Microsoft.Xrm.Data.PowerShell** – do User Map i (opcjonalnie) budowy schematu

## Szybki start

1. **Połączenia** – `Config\Polaczenia.txt`: `ZrodloUrl`, `ZrodloLogin`, `ZrodloHaslo`, `CelUrl`, `CelLogin`, `CelHaslo` (lub zmienne środowiskowe / `CMTConfig.ps1`).
2. **Lista encji** – w `Config\CMTConfig.ps1`: `SchemaEntityIncludeOnly`, `SchemaEntityOrder`.
3. **Uruchom menu** – dwuklik `MigracjaCMT.bat` lub `.\Start-CMTMigrationMenu.ps1`.
4. **Przepływ:** CMT Export → w menu **2** (User Map), potem **3** (Transformuj) → gotowy zip w `Output\*_ForTarget.zip`. Import wykonujesz sam w CMT lub innym narzędziu.

## Instalacja modułu CMT

```powershell
.\Install-CMTModule.ps1
```

Lub ręcznie:

```powershell
Install-Module -Name Microsoft.Xrm.Tooling.ConfigurationMigration -MinimumVersion 1.0.0.88 -Scope CurrentUser
```

## Konfiguracja

1. **Config\CMTConfig.ps1**  
   - `SourceConnectionString`, `TargetConnectionString` – ten sam format co w głównej aplikacji (OAuth; Url; Username; Password; itd.). Można nadpisać zmiennymi środowiskowymi `DATAVERSE_SOURCE_CONNECTION`, `DATAVERSE_TARGET_CONNECTION`.
   - `SchemaFilePath` – ścieżka do pliku schematu CMT (data_schema.xml). **Obowiązkowe** do eksportu i pełnej migracji.
   - `ExportOutputDirectory` – katalog na wyeksportowany plik zip.
   - `ImportDataFileName` – nazwa pliku zip (domyślnie `CMT_Export.zip`).
   - `UserMapFilePath` – opcjonalnie plik mapowania użytkowników (źródło → cel) dla właścicielstwa.

2. **Schemat CMT**  
   - Utwórz schemat w [Configuration Migration Tool](https://learn.microsoft.com/en-us/power-platform/admin/create-schema-export-configuration-data) (DataMigrationUtility.exe): wybierz organizację źródłową, encje i pola, zapisz schemat.
   - Aby migrować „tylko to, co można” do celu: możesz użyć schematu wyeksportowanego ze źródła; encje/pola nieobecne w celu będą pomijane lub zgłoszą błąd – wtedy przycinasz schemat do encji/pól istniejących w celu (ręcznie lub w przyszłości skryptem porównującym schematy).

## Uruchomienie

**Pełna migracja (eksport ze źródła + import do celu):**

```powershell
.\Start-CMTMigration.ps1
```

Wymaga ustawionego `SchemaFilePath` oraz connection stringów w configu (lub parametrach).

**Tylko eksport ze źródła:**

```powershell
.\Start-CMTMigration.ps1 -ExportOnly -SchemaFilePath "C:\path\to\data_schema.xml"
```

**Tylko import do celu (z wcześniej wyeksportowanego pliku):**

```powershell
.\Start-CMTMigration.ps1 -ImportOnly -ImportDataFile "C:\path\to\CMT_Export.zip"
```

Opcjonalnie mapowanie użytkowników:

```powershell
.\Start-CMTMigration.ps1 -ImportOnly -ImportDataFile "C:\path\to\CMT_Export.zip" -UserMapFilePath "C:\path\to\usermap.xml"
```

**Parametry z wiersza poleceń** (nadpisują config):  
`-ConfigPath`, `-SchemaFilePath`, `-SourceConnectionString`, `-TargetConnectionString`, `-ExportOnly`, `-ImportOnly`, `-ImportDataFile`, `-UserMapFilePath`, `-LogDirectory`, `-DisableTelemetry`.

## Daty utworzenia i właścicielstwo

- **Daty:** CMT (od wersji 1.0.0.60) obsługuje opcje dotyczące dat przy imporcie (np. „Updating Dates on insert”). Zachowanie zależy od ustawień w samym CMT – sprawdź dokumentację [Import configuration data](https://learn.microsoft.com/en-us/power-platform/admin/import-configuration-data).
- **Właściciel:** Podczas importu, jeśli dane zawierają użytkowników źródła, CMT może wygenerować lub użyć pliku **User Map** (mapowanie użytkowników źródło → cel), żeby ustawić właścicielstwo w celu. Ustaw `UserMapFilePath` w configu lub `-UserMapFilePath`.

## Migracja danych przez CMT – co masz z tego, czego brakowało

Możesz **migrować dane przez CMT** i nadal mieć to, czego wcześniej brakowało:

| Wymaganie | Rozwiązanie |
|-----------|-------------|
| **Data oryginalna tworzenia** | W schemacie CMT muszą być pola **createdon** i/lub **overriddencreatedon** w encjach. Opcja „Pobierz schemat ze źródła” (menu 5) bierze atrybuty z metadanych – jeśli encja ma `createdon` / `overriddencreatedon`, trafią do schematu. CMT przy imporcie ustawia te wartości z eksportu (zachowanie zależy od wersji CMT i opcji importu). |
| **Właścicielstwo po imieniu i nazwisku** | CMT domyślnie mapuje użytkowników po identyfikatorze (np. domena/login). Aby mapować **po imieniu i nazwisku**, użyj skryptu **`Lib\New-CMTUserMapByDisplayName.ps1`**: łączy się ze źródłem i celem, dopasowuje użytkowników po „Imię Nazwisko” i generuje plik User Map. Ustaw w configu `UserMapFilePath` na wygenerowany plik (lub podaj go przy `-UserMapFilePath` przy imporcie). |
| **Stagi na szansach (opportunity)** | W schemacie CMT są encje **processstage**, **opportunitysalesprocess**, **leadtoopportunitysalesprocess** (w `SchemaEntityIncludeOnly` i `SchemaEntityOrder`). Eksport/import w tej kolejności przenosi procesy i stagi; szanse (opportunity) mają wtedy powiązane rekordy BPF i aktywny stage. |

**Kiedy migrować przez CMT, a kiedy przez DataverseMigration?**

- **CMT** – jeden wspólny schemat (data_schema.xml), oficjalne narzędzie Microsoft, wieloprzejściowy import (zależności), User Map z pliku. Dobrze się sprawdza, gdy środowiska są dość zbieżne (te same/s podobne rozwiązania) i chcesz jednego pliku zip + importu z mapowaniem użytkowników.
- **DataverseMigration** – pełna kontrola: tylko encje z danymi + zależności, upsert, `overriddencreatedon`, mapowanie owner po IdMap (GUID→GUID). Przydatne, gdy cel ma inne rozwiązania (brakujące pola), gdy chcesz mapować owner po imieniu/nazwisku w jednym skrypcie (bez osobnego pliku User Map) albo gdy wolisz nie używać CMT.

## Logi

Logi zapisywane są w `DataverseMigrationCMT\Logs\` (np. `CMT_yyyyMMdd.log`) oraz w katalogu podanym w `-LogDirectory` / configu (używanym też przez CMT jako `LogWriteDirectory`).

## Krótkie podsumowanie

- **Druga aplikacja** – nie zmienia głównej `DataverseMigration`.
- **Schematy:** Źródło/cel – schemat tworzysz w CMT (GUI) lub używasz istniejącego; migrujesz tylko to, co opisane w schemacie (cel może przycinać brakujące encje/pola).
- **Daty i właściciel:** CMT obsługuje daty i mapowanie użytkowników (User Map).
- **Moduł:** [Microsoft.Xrm.Tooling.ConfigurationMigration 1.0.0.88](https://www.powershellgallery.com/packages/Microsoft.Xrm.Tooling.ConfigurationMigration/1.0.0.88).

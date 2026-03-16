# Instrukcja: migracja danych z użyciem CMT i DataverseMigrationCMT

Prosty przepływ: **eksportujesz dane z CMT** (źródło) → **ta aplikacja przygotowuje zip** (mapowanie użytkowników, daty, pola dla celu) → **importujesz w CMT** (cel).

---

## 1. Co jest potrzebne

- **PowerShell 5.1** (lub nowszy), uruchomiony z uprawnieniami do skryptów.
- **CRM Configuration Migration (CMT)** – narzędzie Microsoft do eksportu i importu danych (Data Migration Utility / Configuration Migration Tool). Pobierz z dokumentacji Power Platform.
- **Moduł PowerShell:** `Microsoft.Xrm.Data.PowerShell` (do mapowania użytkowników i metadanych celu).  
  Instalacja: `.\Install-CMTModule.ps1` albo ręcznie:
  ```powershell
  Install-Module -Name Microsoft.Xrm.Data.PowerShell -Scope CurrentUser
  ```
- **Dane logowania** do środowiska **źródłowego** (skąd eksportujesz) i **docelowego** (do którego importujesz) – np. URL, login, hasło lub connection string.

---

## 2. Konfiguracja (przed pierwszym uruchomieniem)

### 2.1 Połączenia – `Config\Polaczenia.txt`

Wypełnij (linie z `#` są pomijane):

- **Źródło:** `ZrodloUrl`, `ZrodloLogin`, `ZrodloHaslo`
- **Cel:** `CelUrl`, `CelLogin`, `CelHaslo`

Przykład:

```
ZrodloUrl=https://moja-org.crm4.dynamics.com/
ZrodloLogin=admin@firma.pl
ZrodloHaslo=Haslo123
CelUrl=https://cel-org.crm4.dynamics.com/
CelLogin=admin@firma.pl
CelHaslo=Haslo123
```

**Nie commituj tego pliku z hasłami do repozytorium.**

### 2.2 Opcjonalnie – `Config\CMTConfig.ps1`

Możesz ustawić m.in. katalog eksportu CMT, ścieżkę do schematu, akcję walidacji option setów (`Report` / `Clear` / `Replace` / `Interactive`). Domyślne wartości wystarczą do standardowego przepływu.

---

## 3. Przepływ krok po kroku

### Krok A: Eksport w CMT (źródło)

1. Uruchom **CRM Configuration Migration** (Data Migration Utility).
2. Wybierz **Export data** i połącz się ze **środowiskiem źródłowym**.
3. Wybierz encje i pola do eksportu (albo użyj istniejącego schematu).
4. Zapisz wynikowy plik **zip** – najlepiej do folderu **`Input`** tej aplikacji (np. `Input\moj_export.zip`).

### Krok B: Uruchomienie aplikacji

- **Dwuklik:** `MigracjaCMT.bat`  
  **lub** w PowerShell z folderu aplikacji:  
  `.\Start-CMTMigrationMenu.ps1`

### Krok C: W menu aplikacji

| Opcja | Co zrobić |
|-------|-----------|
| **1** | Wybierz zip z CMT (jeśli nie wrzuciłeś do `Input` – wskaż plik). |
| **2** | **Generuj User Map** – mapowanie użytkowników (imię i nazwisko) źródło → cel oraz encji (lead, account, opportunity, contact) po polu kluczowym. Tworzy User Map XML i **IdMap JSON** (ownerid + **objectid** w uwagach). Wykonaj **przed** opcją 3. Przy imporcie uwag uruchom opcję 2 ponownie (odśwież mapowania), potem opcję 3. |
| **5** | **Pobierz metadane celu** – cache encji i pól w celu (używany przy opcji 3 do usuwania pól nieistniejących w celu). Wykonaj **raz na organizację docelową** (np. po zmianie celu). |
| **3** | **Transformuj zip** – podmiana ownerów (IdMap), ustawienie dat utworzenia (overriddencreatedon), usunięcie pól nieistniejących w celu, walidacja option setów. Wynik: plik **`*_ForTarget.zip`** w folderze **`Output`**. |

**Typowa kolejność:** 1 → 2 → (5 jeśli jeszcze nie robiłeś dla tego celu) → 3.

### Krok D: Import w CMT (cel)

1. W **CRM Configuration Migration** wybierz **Import data**.
2. Połącz się ze **środowiskiem docelowym**.
3. Wskaż plik **`Output\<nazwa>_ForTarget.zip`** (ten po transformacji).
4. Jeśli CMT zapyta o mapowanie użytkowników – możesz użyć wygenerowanego w opcji 2 pliku User Map (np. `Output\CMT_UserMap_ByDisplayName.xml`), jeśli jest wymagany przez CMT.
5. Uruchom import. Czekaj na zakończenie (duże zbiory mogą trwać długo).

---

## 4. Gdzie co leży

| Miejsce | Zawartość |
|---------|-----------|
| **Input** | Zipy wyeksportowane z CMT (źródło). Tu wrzucasz pliki przed uruchomieniem opcji 3. |
| **Output** | `*_ForTarget.zip` – gotowy zip do importu w CMT; User Map; IdMap; cache metadanych celu (`TargetMetadata_*.json`); statystyki ostatniego uruchomienia (do szacowania czasu). |
| **Config** | `Polaczenia.txt` (hasła!), `CMTConfig.ps1` (opcje). |
| **Logs** | Logi z menu (każde uruchomienie tworzy nowy plik). |

---

## 5. Typowe problemy

- **„Missing Fields on Entity … overriddencreatedon”**  
  Aplikacja usuwa ze schematu encje, które w zipie nie mają żadnych rekordów (np. `salesliteratureitem`). Uruchom **ponownie opcję 3** – w nowym zipie ta encja zniknie ze schematu i błąd nie powinien się pojawić.

- **„Missing Comparison Key for entity Uwaga, missing subject”**  
  Encja annotation (Uwaga) wymaga w schemacie **primarynamefield=subject**. Przy **opcji 3 (Transformuj zip)** schemat w zipie jest automatycznie poprawiany. Jeśli nadal widzisz błąd, upewnij się, że w zipie jest plik schematu (np. data_schema.xml) i że transformacja go przetwarza.

- **„The parent object type was present, but the ID was missing” (uwagi/annotation)**  
  Uwagi są powiązane z rekordem nadrzędnym (lead, opportunity, konto itd.) przez pole **objectid**. Transformacja podmienia **objectid** na GUID z celu na podstawie **IdMap**. **Opcja 2 (Generuj User Map)** tworzy IdMap z użytkownikami oraz z encjami lead, account, opportunity, contact (dopasowanie po e-mailu, numerze konta, numerze szansy itd.). Odśwież mapowania: uruchom **opcję 2 ponownie** (po zaimportowaniu lead/kont/szans do celu), potem **opcję 3** (transformuj zip z uwagami). Jeśli encje w celu mają inne klucze niż w źródle, możesz ręcznie uzupełnić IdMap (format JSON: klucz = GUID źródła, wartość = GUID celu).

  **Test: czy problem jest w mapowaniu, czy w schemacie?** Uruchom skrypt **`Lib\Patch-AnnotationZipWithTargetLeadIds.ps1`** (wymaga połączenia do celu w Config). Skrypt pobiera ID leadów z celu, podmienia w zipie uwag wszystkie objectid na jeden z tych ID i zapisuje `Output\annotation_ForTarget_Test.zip`. Zaimportuj ten zip w CMT. Jeśli coś się zaimportuje → problem był w mapowaniu (opcja 2 / IdMap). Jeśli dalej 0 rekordów → możliwy problem ze schematem/formatem objectid.

- **Uwagi (annotation) – brak oryginalnej daty utworzenia po imporcie**  
  W logu CMT możesz zobaczyć: *"DataField createdon is present in coming data but is not present in provided schema definition, skipping field on import"* oraz *"Principal user is missing prvOverrideCreatedOnCreatedBy privilege"*.

  - **Dlaczego daty znikają:** CMT nie importuje pól, których nie ma w schemacie – stąd pomijane są `createdon`/`modifiedon`. Aplikacja ustawia w zipie pole **overriddencreatedon** (po opcji 6 + opcji 3), które Dataverse honoruje tylko wtedy, gdy użytkownik importujący ma uprawnienie **Override Created On and Created By** (prvOverrideCreatedOnCreatedBy). Bez tego uprawnienia Dataverse ignoruje `overriddencreatedon` i ustawia bieżącą datę.

  - **Co zrobić:** W środowisku docelowym (Power Platform / Zarządzanie zabezpieczeniami) nadaj **roli użytkownika, którym importujesz w CMT**, uprawnienie **„Override Created On and Created By”** (lub równoważne) – na poziomie organizacji lub dla encji. Po nadaniu uprawnienia ponów import uwag (zip po opcji 3). Alternatywa: import uwag **przez API** skryptem **`Lib\Import-CMTZipToDataverse.ps1`** (użytkownik z Config musi mieć to uprawnienie).

- **„Stage Failed” przy imporcie w CMT**  
  W menu wybierz **opcję 4 – Pokaz ostatni log bledow CMT**. W logu szukaj linii z: Exception, Failed, Error, Invalid, Missing. Częste przyczyny: duplikaty rekordów, pole wymagane puste, wartość option set nie istniejąca w celu, lookup do rekordu nieistniejącego w celu.  
  Więcej: **TROUBLESHOOTING-StageFailed.md**.

- **Błędne szacowanie czasu**  
  Czas oczekiwania na wybór option setów (gdy jest tryb Interactive) nie wlicza się do szacunku. Statystyki zapisują tylko „czas aktywny” – kolejne uruchomienia dają realistyczne szacunki.

---

## 6. Podsumowanie

1. **CMT Export** (źródło) → zip do `Input`.
2. **Menu:** 1 (wybierz zip) → 2 (User Map) → 5 (metadane celu, raz) → 3 (transformuj).
3. **CMT Import** (cel) → plik `Output\*_ForTarget.zip`.

- **README.md** – dokumentacja techniczna, różnice względem innej aplikacji migracyjnej, konfiguracja zaawansowana.
- **TROUBLESHOOTING-StageFailed.md** – szczegóły przy błędzie „Stage Failed” w CMT.
- **RESTORE_POINT.md** – opis stanu działającej wersji (punkt przywracania), gdy coś się zepsuje.

# Schemat encji Annotation (Uwaga) – CMT i Dataverse

Krótkie podsumowanie na podstawie dokumentacji Microsoft i typowego formatu CMT.

## Dataverse (dokumentacja Microsoft)

- **Encja:** `annotation` (Note / Uwaga).
- **Klucz główny:** `annotationid`.
- **Pole nazwy (primary name):** `subject` – wymagane przy imporcie CMT (primarynamefield w schemacie).

### Powiązanie z rekordem nadrzędnym

| Właściwość        | Typ        | Opis |
|-------------------|------------|------|
| **objectid**      | Lookup     | Unikalny identyfikator rekordu, do którego przypisana jest uwaga (lead, account, contact, opportunity itd.). |
| **objecttypecode**| EntityName | Typ encji rekordu nadrzędnego (np. `lead`, `account`, `contact`). |

Źródło: [annotation EntityType (Web API)](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/reference/annotation), [Note (Annotation) table reference](https://learn.microsoft.com/en-us/power-apps/developer/data-platform/reference/entities/annotation).

### Przykładowa odpowiedź API (uwaga przy leadzie)

W odpowiedzi Web API pole lookup do rekordu nadrzędnego jest zwracane jako **`_objectid_value`** (GUID) oraz **`objecttypecode`** (typ encji). Reszta pól jak w dokumentacji.

```json
{
  "@odata.context": "https://.../api/data/v9.2/$metadata#annotations/$entity",
  "objecttypecode": "lead",
  "_objectid_value": "8aa4cbc8-5820-f111-8342-6045bdde9566",
  "annotationid": "ae43b1cf-31b4-d803-aa3f-790717cc7c4b",
  "subject": "tst",
  "notetext": "<div class=\"ck-content\" ...>...</div>",
  "isdocument": false,
  "filesize": 0,
  "mimetype": null,
  "documentbody": null,
  "filename": null,
  "_owninguser_value": "5a0fbddd-1f14-f111-8341-7ced8d71b7a3",
  "_ownerid_value": "5a0fbddd-1f14-f111-8341-7ced8d71b7a3",
  "_owningbusinessunit_value": "0fa9c4d7-1f14-f111-8341-7ced8d71b7a3",
  "createdon": "2026-03-15T10:22:04Z",
  "modifiedon": "2026-03-15T10:22:04Z",
  "isautonomouslycreated": false,
  "overriddencreatedon": null
}
```

Przy tworzeniu rekordu (SDK / Add-CrmRecord) przekazujemy **objectid** jako **EntityReference(logicalName, id)** np. `EntityReference("lead", guid)`, co odpowiada parze `_objectid_value` + `objecttypecode` w API. Skrypt **Import-CMTZipToDataverse.ps1** ustawia `objectid` w ten sposób na podstawie pól z data.xml.

### Tworzenie uwagi (API)

Przy tworzeniu uwagi do leada typowy format referencji to np.:

- **Web API:** `objectid = { id: leadGuid, logicalname: "lead", name: "Lead Name" }`
- **Lookup:** `_objectid_value` = GUID rekordu nadrzędnego.

Ważne: przy podawaniu **objectid** warto (lub trzeba, zależnie od kontekstu) uzupełnić **objecttypecode**, żeby system wiedział, że chodzi np. o encję `lead`.

---

## Format w zipie CMT (eksport / import)

- W **data_schema.xml** (schemat) pole **objectid** jest zdefiniowane jako `type="entityreference"` z `lookupType="...|lead|...|opportunity|...|account|...|contact|..."`.
- W **data.xml** (rekordy) wartość lookupu jest zwykle w jednej z postaci:
  - **`guid leadid=<GUID>`** – dla leada (nazwa encji + atrybut klucza = wartość),
  - **`guid accountid=<GUID>`** – dla konta,
  - **`guid opportunityid=<GUID>`** – dla szansy,
  - **`guid contactid=<GUID>`** – dla kontaktu,
  - czasem sam **`<GUID>`** (bez prefiksu).

Prefiks **`guid <entity>id=`** określa **typ encji** rekordu nadrzędnego, więc CMT może z niego wywnioskować „objecttypecode” przy imporcie.

### Gdzie w zipie jest objectid

- **Atrybut rekordu:** `<record ... objectid="guid leadid=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx">`
- **Element / pole:** `<field name="objectid" ... value="guid leadid=...">` lub `<objectid>...</objectid>` z wartością (GUID lub pełna forma).

---

## Co robią nasze skrypty

- **Transform (opcja 3)** i **Patch-AnnotationZipWithTargetLeadIds.ps1** podmieniają **objectid** na GUID z org docelowej, w formacie **`guid leadid=<targetGUID>`** (dla testu z leadami).
- **objecttypecode** w data.xml nie jest dziś ustawiane osobno – CMT najpewniej korzysta z prefiksu `leadid` w wartości objectid. Jeśli przy poprawnym GUID celu import dalej się nie udaje, warto w zipie **dla każdego rekordu uwagi** ustawić też pole **objecttypecode** na `lead` (gdy objectid wskazuje na leada).

---

## Podsumowanie

1. **objectid** = lookup do rekordu nadrzędnego (lead, account, contact, opportunity itd.).
2. **objecttypecode** = nazwa encji rekordu nadrzędnego (np. `lead`).
3. W zipie CMT typowy format to **`guid leadid=<GUID>`** (dla leada) – prefiks określa typ encji.
4. W naszym schemacie (`Input\annotation.xml`) **objectid** ma `lookupType` zawierające `lead` – jest zgodne z dokumentacją.
5. Jeśli po podstawieniu prawidłowych GUID leadów z celu import nadal zwraca 0 rekordów, warto sprawdzić w data.xml, czy CMT wymaga także wypełnionego **objecttypecode** (np. `lead`) obok objectid.

---

## Uwagi – oryginalna data utworzenia (createdon) i uprawnienie

- W imporcie CMT w logu mogą pojawić się: *"DataField createdon is present in coming data but is not present in provided schema definition, skipping field on import"* oraz *"Principal user is missing prvOverrideCreatedOnCreatedBy privilege"*.
- **createdon/modifiedon w schemacie:** CMT nie importuje pól niewymienionych w schemacie. W naszym flow używamy **overriddencreatedon** (ustawiane w opcji 6 + opcja 3) – to pole jest dodawane do schematu przez transformację. Komunikat o pomijaniu `createdon` jest oczekiwany (pole jest w data z opcji 6, ale w schemacie CMT celowo używamy tylko overriddencreatedon).
- **Dlaczego część uwag ma błędną datę:** Dataverse stosuje wartość **overriddencreatedon** tylko wtedy, gdy użytkownik wykonujący import ma uprawnienie **prvOverrideCreatedOnCreatedBy** („Override Created On and Created By”). Bez niego system ignoruje overriddencreatedon i ustawia bieżącą datę.
- **Rozwiązanie:** W org docelowej nadać roli użytkownika importującego (CMT lub konto z Config przy imporcie API) uprawnienie **Override Created On and Created By**. Szczegóły i alternatywy (import przez API): **INSTRUKCJA.md**, sekcja „Typowe problemy” → „Uwagi – brak oryginalnej daty utworzenia”.

---

## Co znaleziono w internecie (podobne problemy)

### 1. LookupRecordNotAvailable / „Skipping Lookup … Not Available in target System”

- **Źródło:** [Nishant Rana – LookupRecordNotAvailable w CMT](https://nishantrana.me/2018/09/24/lookuprecordnotavailable-error-while-using-configuration-migration-tool-in-dynamics-365-customer-engagement/)
- **Przyczyna:** CMT nie znajduje w systemie docelowym rekordu wskazywanego przez pole lookup (np. `createdonbehalfby`, a u nas `objectid`).
- **Proponowane rozwiązania:**
  - **Opcja A:** Usunąć pole lookup ze schematu, jeśli nie trzeba go migrować (np. `createdonbehalfby`).
  - **Opcja B:** Upewnić się, że rekord wskazywany przez lookup **już istnieje w org docelowej** przed importem. Dokumentacja Microsoft: import jest wieloprzejściowy – najpierw „foundation data”, potem dane zależne; rekordy referencjonowane muszą być dostępne w celu.

### 2. Dopasowanie po GUID vs. po nazwie (primary name)

- **Źródła:** [Develop1 – non-unique display name](https://www.develop1.net/public/post/2015/03/25/Using-the-Configuration-Data-Migration-Tool-with-non-unique-display-name-values), dokumentacja CMT, [GitHub issue #957 (powerplatform-build-tools)](https://github.com/microsoft/powerplatform-build-tools/issues/957).
- **Zachowanie CMT:** Dla **encji importowanej** (np. annotation) narzędzie domyślnie porównuje rekordy po **polu primary name** (np. `subject`). Aby dopasowywać po **GUID** (primary key), w schemacie trzeba ustawić **Configure Import Settings** i dodać w polu klucza atrybut `updateCompare="true"`.
- **Dla pól typu lookup (entityreference):** CMT może rozwiązywać lookup w celu używając **zarówno GUID, jak i primary attribute (display name)** rekordu docelowego. Przy duplikatach nazw (np. wielu leadach z tym samym fullname) może dochodzić do pomyłek lub pomijania rekordów. Rekomendacja Microsoft: definiować **uniqueness conditions / composite keys**, żeby jednoznacznie identyfikować rekordy.

### 3. „The parent object type was present, but the ID was missing”

- W kontekście CMT ten komunikat często idzie w parze z problemami lookupów: typ encji (np. lead) jest rozpoznany, ale **identyfikator rekordu** nie zostaje poprawnie dopasowany w systemie docelowym (np. CMT szuka po nazwie zamiast po GUID albo nie znajduje rekordu o podanym GUID).

### 4. Praktyczne rekomendacje dla importu annotation → lead

1. **Ta sama organizacja:** Przy imporcie w CMT upewnij się, że łączysz się z **tą samą org docelową** (URL), co w `Config/Polaczenia.txt` (CelUrl). Inna org = brak „znanego” leada = LookupRecordNotAvailable.
2. **Kolejność migracji:** Najpierw upewnij się, że **leady** są w org docelowej (zaimportowane wcześniej lub już istniejące). Potem importuj **annotation** z `objectid` wskazującym na te leady.
3. **Format `objectid` w data.xml:** Trzymać się formatu **`guid leadid=<GUID>`** lub samego **GUID** + **objecttypecode = lead**. Jeśli CMT w Twojej wersji rozwiąże lookup po **primary name** leada (fullname), można **eksperymentalnie** w zipie podać w polu lookup wartość w formacie „lead, \<fullname z celu\>” (jeśli dokumentacja lub próby potwierdzą taki format).
4. **Logi CMT:** Sprawdzić **ImportDataDetail.log** i **DataMigrationUtility.log** w `%AppData%\Roaming\Microsoft\Microsoft Dataverse Configuration Migration Tool\` – często tam widać dokładny powód „Not Available” (np. brak rekordu o podanym GUID lub brak dopasowania po nazwie).
5. **Opcja obejścia:** Jeśli import annotation z lookupem na leady w CMT nadal się nie udaje, rozważyć import uwag **poza CMT** (np. skrypt PowerShell z API – tworzenie rekordów `annotation` z `objectid` = EntityReference(lead, targetLeadId)).

### 5. „Not Available” mimo poprawnego GUID i lookupentity=lead

Jeśli w zipie masz **objectid** = sam GUID (np. `edd57a5b-6314-f011-998a-000d3ab3bc1e`), **lookupentity** = `lead`, a CMT i tak raportuje „Skipping Lookup … Not Available in target System” i „The parent object type was present, but the ID was missing”, to **praktycznie jedyna spójna przyczyna** to:

- **Import w CMT jest wykonywany do innej organizacji niż ta, z której pochodzą ID leadów.**

ID leadów w zipie (np. z `TargetLeadIds.json`) pochodzą z połączenia **Config (Polaczenia.txt: CelUrl, CelLogin, CelHaslo)**. CMT przy imporcie używa organizacji wybranej w oknie CMT (logowanie przed importem). Jeśli w CMT zalogujesz się do innego środowiska/org niż CelUrl, lead o tym GUID tam nie istnieje → CMT słusznie raportuje „Not Available”.

**Co zrobić:**

1. **Sprawdź URL organizacji w CMT** – po zalogowaniu w CMT (przed importem) upewnij się, że adres środowiska/org jest **identyczny** z **CelUrl** z `Config\Polaczenia.txt` (np. `https://xentivocrm.crm4.dynamics.com`).
2. **Lookupentityname w zipie** – CMT rozstrzyga lookup m.in. po **lookupentityname** (nazwa rekordu w celu). Uruchom **Patch-AnnotationZipWithTargetLeadIds.ps1** **bez** `-UseExistingTargetLeadIds`, żeby skrypt pobrał z celu leady wraz z **fullname** i zapisał w zipie `lookupentityname` przy polu objectid. Wtedy w logu CMT zamiast „Name: ” będzie „Name: &lt;fullname leada&gt;”. Następnie zaimportuj wygenerowany zip w CMT.
3. **Format wartości objectid** – jeśli nadal 0 rekordów, uruchom Patch z **`-UseGuidLeadIdFormat`** – wartość w zipie będzie w formacie `guid leadid=&lt;GUID&gt;` zamiast samego GUID.
4. **Import uwag przez API (gdy CMT nie wchodzi w grę)** – użyj skryptu **`Lib\Import-CMTZipToDataverse.ps1`** z zipem zawierającym tylko encję annotation (np. `Output\annotation_ForTarget_Test.zip`). Skrypt łączy się do celu z **Config** (ta sama org co CelUrl), więc `objectid` zostanie poprawnie ustawiony jako EntityReference(lead, guid).

   ```powershell
   .\Lib\Import-CMTZipToDataverse.ps1 -ZipPath ".\Output\annotation_ForTarget_Test.zip" -ConfigPath ".\Config\CMTConfig.ps1"
   ```

   Wymaga: `Microsoft.Xrm.Data.PowerShell`. Pola `objectid` i `objecttypecode` są odczytywane z data.xml i mapowane na EntityReference (lead, guid) przy tworzeniu rekordów.

---

## Weryfikacja zipa i test (15.03.2026)

### Budowa zipa `annotation_ForTarget_Test.zip`

- **data.xml:** encja `annotation`, 300 rekordów; każdy rekord ma atrybut `objecttypecode="lead"` oraz pole `<field name="objectid" value="edd57a5b-6314-f011-998a-000d3ab3bc1e" lookupentity="lead">`.
- **data_schema.xml:** encja `annotation`, `primaryidfield="annotationid"`, `primarynamefield="subject"`; pole `objectid` typu `entityreference` z `lookupType` zawierającym `lead`.

Budowa jest zgodna z oczekiwaniami CMT/Dataverse. Wszystkie uwagi wskazują na leada `edd57a5b-6314-f011-998a-000d3ab3bc1e`.

### Logi importu CMT

- **"Skipping Lookup … objectid requesting lookup to lead ID: edd57a5b-6314-f011-998a-000d3ab3bc1e … Not Available in target System"** – CMT nie znalazł tego leada w organizacji, do której importuje.
- **"The parent object type was present, but the ID was missing"** – konsekwencja: po „pominięciu” lookupu CMT nie ma ID rodzica i odrzuca insert.

Jeśli import w CMT jest do **tej samej** org co CelUrl z Config, to w tej org lead **powinien** istnieć (ID pochodzi z Config). Wtedy możliwe, że CMT ma ograniczenia przy polimorficznym lookupie (objectid) i nie rozpoznaje samego GUID + lookupentity=lead.

### Sprawdzenie, czy lead istnieje w celu

Uruchom (wymaga Microsoft.Xrm.Data.PowerShell i poprawnego Config):

```powershell
.\Lib\Verify-TargetLeadAndAnnotation.ps1 -LeadId "edd57a5b-6314-f011-998a-000d3ab3bc1e" -ConfigPath ".\Config\CMTConfig.ps1"
```

Skrypt łączy się z org docelową z Config i sprawdza, czy rekord lead o podanym ID istnieje.

### Zip do szybkiego testu (2 rekordy)

Wygenerowany plik: **`Output\annotation_ForTarget_Test_2records.zip`** (ta sama struktura, 2 rekordy annotation). Do ponownego wygenerowania:

```powershell
.\Lib\New-AnnotationTestZip.ps1 -SourceZip ".\Output\annotation_ForTarget_Test.zip" -OutZip ".\Output\annotation_ForTarget_Test_2records.zip" -RecordCount 2
```

### Test importu przez API

1. Sprawdź lead w celu: `.\Lib\Verify-TargetLeadAndAnnotation.ps1`
2. Zaimportuj uwagi przez API (nie CMT):

   ```powershell
   .\Lib\Import-CMTZipToDataverse.ps1 -ZipPath ".\Output\annotation_ForTarget_Test_2records.zip" -ConfigPath ".\Config\CMTConfig.ps1"
   ```

Jeśli import API przejdzie, problem leży po stronie CMT (lookup polimorficzny / inna org). Pełny zip 300 rekordów: `.\Output\annotation_ForTarget_Test.zip`.

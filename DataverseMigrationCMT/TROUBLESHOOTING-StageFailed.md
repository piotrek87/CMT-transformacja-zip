# Stage Failed przy imporcie CMT

Gdy w CRM Configuration Migration pojawia się **"Przetwarzanie encji: Klient - Stage Failed"** (bez szczegółów w oknie), przyczyna jest w szczegółowym logu.

## 1. Zobacz szczegółowy log

- **W menu:** wybierz opcję **4. Pokaz ostatni log bledow CMT** – skrypt wyszuka logi w AppData i wypisze wpisy z błędami.
- **Ręcznie:**  
  - W katalogu, z którego uruchamiasz **CRM Configuration Migration** (np. `Tools\ConfigurationMigration\`), otwórz:
    - **DataMigrationUtility.log**
    - **ImportDataDetail.log**
  - Albo w `%AppData%\Microsoft` lub `%LocalAppData%\Microsoft` poszukaj podkatalogów z nazwą zawierającą "Migration" / "Configuration" i plików `.log`.

W logu szukaj linii z: **Exception**, **Failed**, **Error**, **Invalid**, **Missing**. Tam będzie dokładny komunikat (np. brak wymaganego pola, nieprawidłowa wartość option set, odwołanie do rekordu, który nie istnieje w celu).

## 2. Typowe przyczyny Stage Failed

| Przyczyna | Co zrobić |
|-----------|-----------|
| **Element o tym samym kluczu został już dodany** | W zipie są **duplikaty rekordów** (ta sama encja, ten sam klucz główny). Usuń duplikaty w źródle lub wyeksportuj ponownie bez zduplikowanych rekordów. |
| **Lookup do rekordu nieistniejącego w celu** | CMT pomija referencje do rekordów nieobecnych w celu (np. msdyn_accountkpiid, transactioncurrencyid, originatingleadid). Jeśli to powoduje Stage Failed, usuń te pola w źródle lub zaimportuj najpierw brakujące encje. |
| **Wymagane pole puste** | W celu encja może mieć pole wymagane, którego w eksporcie nie ma lub jest puste. Sprawdź w logu nazwę pola i uzupełnij w źródle lub dodaj wartość domyślną w transformacji. |
| **Wartość option set nie istnieje w celu** | Użyj w transformacji walidacji option setów (Config: `OptionSetValidationAction` = Replace/Interactive) albo usuń wartość pola (Clear). |
| **Plugin/workflow w celu blokuje zapis** | W dokumentacji CMT sprawdź, czy import obsługuje parametr typu *BypassCustomPluginExecution* (jeśli masz uprawnienia). |
| **Nazwa encji / schemat** | "Klient" to nazwa wyświetlana; w zipie encja to zwykle `account`. Upewnij się, że schemat importu zawiera encję `account` i że w celu nie ma konfliktu (np. inna definicja encji). |

## 3. Po znalezieniu błędu w logu

- Jeśli chodzi o **konkretne pole** – usuń je w źródle, popraw option set / wartość domyślną albo zaimportuj brakujące rekordy (encje) w celu.
- Ponownie uruchom **opcję 3** (transformacja zipa), potem ponowny import w CMT.

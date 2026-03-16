# Punkt przywracania – migracja Uwag (annotation) z oryginalną datą OK

**Data:** 2026-03-16  
**Tag:** `restore-uwagi-daty-2026-03-16`  
**Stan:** Opcja 3 (Transform) + opcja 6 (Popraw daty Uwag) działają; daty w ISO 8601; można robić 3 potem 6 bez ponownej opcji 3.

## Co to za moment

- **Uwagi (annotation):** oryginalna data utworzenia zachowana przy imporcie – w zipie pole `overriddencreatedon` w formacie ISO 8601 (np. `2025-11-28T12:10:00.0000000`).
- **Przepływ:** najpierw opcja 3 (Transformuj zip) → potem opcja 6 (Popraw daty Uwag ze źródła) na zipie `*_ForTarget.zip`. Wynik od razu do importu w CMT (bez ponownego uruchamiania opcji 3).
- **Opcja 6 (Inject):** pobiera createdon/modifiedon ze źródła (lub z pól już w zipie, gdy źródło nie zwróci rekordu), zapisuje w formacie ISO; naprawiony błąd parsera (string bez „mądrych” cudzysłowów/myślników).
- **Opcja 3 (Transform):** normalizuje daty do ISO przy ustawianiu overriddencreatedon.
- **Reguła .cursor:** w plikach `.ps1` tylko ASCII w stringach (cudzysłowy `"` `'`, myślnik `-`), żeby uniknąć „The string is missing the terminator”.

## Gdzie wrócić w razie problemów

- `Lib\Inject-AnnotationDatesFromSource.ps1` – opcja 6, funkcja `ConvertTo-DateTimeIso`, fallback dat z zipa gdy brak w źródle.
- `Lib\Transform-CMTZip.ps1` – opcja 3, `ConvertTo-DateTimeIso`, ustawianie overriddencreatedon.
- `.cursor\rules\powershell-ascii-strings.mdc` – reguła ASCII dla .ps1.
- INSTRUKCJA.md – sekcja „Uwagi – brak oryginalnej daty utworzenia” (uprawnienie prvOverrideCreatedOnCreatedBy).

## Wracanie do tego stanu

```bash
git checkout restore-uwagi-daty-2026-03-16
```

(lub odtwórz commit z tym tagiem).

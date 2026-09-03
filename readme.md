# excel-user-pdf-export

Makro VBA do Excela: dla listy użytkowników i listy kolumn (nagłówków) generuje
osobny plik PDF na A4 dla każdego użytkownika, zawierający tylko jego wiersze
i tylko wybrane kolumny.

## Zawartość repo

- `macros/EksportujUzytkownikow.bas` — kod makra do zaimportowania w VBA.
- `examples/dane_testowe.xlsx` — plik testowy: arkusz `Dane` (20 kolumn,
  100 wierszy losowych danych, 8 różnych użytkowników) + arkusz
  `Export_Config` już wypełniony przykładową konfiguracją (filtr na 5 z 20
  kolumn: `Data`, `Zadanie`, `Status`, `Priorytet`, `Klient`). Służy do
  szybkiego przetestowania makra bez potrzeby przygotowywania własnych danych.
- `examples/generate_dane_testowe.py` — skrypt (Python + openpyxl), którym
  wygenerowano powyższy plik testowy; można go zmodyfikować i odpalić
  ponownie, żeby wygenerować inny zestaw danych testowych.

## Wymagania

- Desktopowy Microsoft Excel (Windows lub Mac). **Makra VBA nie działają
  w Excel Online** — jeśli plik jest na SharePoincie, otwórz go w aplikacji
  desktopowej (np. przez "Open in Desktop App"), a nie w przeglądarce.
- Plik musi być zapisany jako `.xlsm` (skoroszyt z obsługą makr).

## Instalacja krok po kroku

1. Otwórz swój plik Excela (arkusz z danymi) w desktopowym Excelu.
2. Zapisz go jako `.xlsm`: **Plik → Zapisz jako → typ pliku: Skoroszyt
   programu Excel z obsługą makr (*.xlsm)**.
3. Otwórz edytor VBA: `Alt+F11`.
4. W drzewku projektu kliknij prawym na nazwę skoroszytu →
   **Insert → Module**.
5. Wklej do nowego modułu całą zawartość pliku
   [`macros/EksportujUzytkownikow.bas`](macros/EksportujUzytkownikow.bas).
6. Zamknij edytor VBA (`Alt+Q`).
7. Dodaj nowy arkusz o nazwie dokładnie `Export_Config` (patrz sekcja
   "Konfiguracja" niżej).
8. (Opcjonalnie) Na arkuszu z danymi dodaj przycisk: **Wstążka → Deweloper →
   Wstaw → Przycisk (formant formularza)**, narysuj go na arkuszu, w oknie
   "Przypisz makro" wybierz `EksportujRaportyUzytkownikow`.
   Jeśli nie widzisz karty "Deweloper": **Plik → Opcje → Dostosuj wstążkę →
   zaznacz "Deweloper"**.
9. Przy pierwszym uruchomieniu Excel może zapytać o włączenie makr — kliknij
   "Włącz zawartość" / "Enable Content" (pasek zaufania pod wstążką).

## Konfiguracja (arkusz `Export_Config`)

Utwórz arkusz `Export_Config` i wypełnij komórki:

| Komórka | Opis | Przykład |
|---|---|---|
| `B1` | Nazwa arkusza z danymi źródłowymi | `Dane` |
| `B2` | Dokładny nagłówek kolumny, w której są nazwy użytkowników | `Użytkownik` |
| `B3` | Ścieżka folderu docelowego na PDF-y. Może zostać pusta — makro poprosi o wybór folderu i zapamięta go tutaj | `C:\Users\Jan\Documents\Raporty` |
| `B4` | Prefiks nazwy pliku (opcjonalny) | `Raport` |
| `A7`, `A8`, `A9`, ... | Lista użytkowników do wyeksportowania (jeden na wiersz) | `Jan Kowalski` |
| `C7`, `C8`, `C9`, ... | Lista nagłówków kolumn do eksportu, w kolejności docelowej. Wpisz `*` w samej `C7`, żeby wyeksportować wszystkie kolumny | `Data`, `Zadanie`, `Status` |

### Przykład

Arkusz `Dane` (dane źródłowe):

| Użytkownik | Data | Zadanie | Status | Notatki wewnętrzne |
|---|---|---|---|---|
| Jan Kowalski | 2026-08-01 | Instalacja czujnika | Zrobione | (kolumna pomijana w eksporcie) |
| Anna Nowak | 2026-08-02 | Przegląd | W trakcie | ... |
| Jan Kowalski | 2026-08-03 | Naprawa | Zrobione | ... |

Arkusz `Export_Config`:

```
B1: Dane
B2: Użytkownik
B3: C:\Users\Jan\Documents\Raporty
B4: Raport

A7: Jan Kowalski
A8: Anna Nowak

C7: Data
C8: Zadanie
C9: Status
```

Uruchomienie makra `EksportujRaportyUzytkownikow` wygeneruje w
`C:\Users\Jan\Documents\Raporty`:

- `Raport_Jan Kowalski_09_03_2026.pdf` — tylko 2 wiersze Jana, kolumny
  Data / Zadanie / Status, format A4.
- `Raport_Anna Nowak_09_03_2026.pdf` — tylko 1 wiersz Anny.

Jeśli w `C7` wpiszesz samo `*` (i usuniesz `C8`, `C9`), PDF-y będą zawierać
wszystkie kolumny z arkusza `Dane`, łącznie z "Notatki wewnętrzne".

## Szybki test na przykładowych danych

1. Otwórz `examples/dane_testowe.xlsx` w desktopowym Excelu.
2. Zapisz jako `.xlsm` i dodaj makro zgodnie z krokami 3-9 z sekcji
   "Instalacja krok po kroku" powyżej (arkusz `Export_Config` już jest
   gotowy w tym pliku, nie trzeba go tworzyć).
3. W `Export_Config` w komórce `B3` wpisz istniejący folder na swoim dysku
   (np. `C:\Temp`).
4. Uruchom makro (`Alt+F8` → `EksportujRaportyUzytkownikow`). Powinny powstać
   3 pliki PDF: dla `Jan Kowalski`, `Anna Nowak`, `Piotr Wisniewski`, każdy
   tylko z kolumnami `Data`, `Zadanie`, `Status`, `Priorytet`, `Klient`.

## Uruchomienie

- Kliknij przycisk na arkuszu (jeśli dodany), albo
- `Alt+F8` → wybierz `EksportujRaportyUzytkownikow` → Uruchom.

Po zakończeniu pojawi się okno z podsumowaniem: ilu użytkowników
wyeksportowano i którzy zostali pominięci (brak pasujących wierszy).

## Uwagi

- Nazwa pliku: `<Prefiks_>Uzytkownik_MM_DD_YYYY.pdf` (data = dzień
  uruchomienia makra). Prefiks pomijany, jeśli `B4` jest puste.
- Znaki niedozwolone w nazwach plików Windows (`\ / : * ? " < > |`) są
  automatycznie zamieniane na `_`.
- Makro tworzy i usuwa tymczasowy arkusz `__TempExport` — nie twórz arkusza
  o tej nazwie ręcznie.
- Dopasowanie użytkownika jest dokładne (bez rozróżniania wielkości liter w
  praktyce Excela) — wartość w `A7:A...` musi się zgadzać z wartością w
  kolumnie użytkowników co do treści (spacje na końcu są przycinane).

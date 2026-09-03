# Arkusz Oceny Kandydatów

Profesjonalny Excel do oceny kandydatów na różne stanowiska.

## Instalacja

```bash
pip install openpyxl
```

## Użycie

Aby wygenerować plik Excel z ocenami kandydatów:

```bash
python Ocena_Kandydatow.py
```

To utworzy plik `Ocena_Kandydatow.xlsx` zawierający:

### Arkusz 1: Ocena kandydatów
- **30 wierszy** do wypełnienia danymi kandydatów
- **Kolumny oceny (1-5)**:
  - Wiedza techniczna
  - Znajomość Microsoft 365
  - Rozwiązywanie problemów
  - Świadomość cyberbezpieczeństwa
  - Umiejętności komunikacyjne
  - Współpraca zespołowa
  - Motywacja

- **Automatyczne funkcje**:
  - Obliczanie wyniku końcowego (średnia z ocen)
  - Automatyczne rekomendacje:
    - ✅ **Zatrudnić** (wynik ≥ 4.5) - zielone
    - ⚠️ **Rozważyć** (wynik 3.5-4.49) - żółte
    - ❌ **Odrzucić** (wynik < 3.5) - czerwone

- **Formatowanie**:
  - Kolorowe kolumny z nagłówkami
  - Walidacja danych (dropdown 1-5 dla ocen)
  - Zamrożona linia nagłówkowa
  - Filtry dostępne dla wszystkich kolumn

### Arkusz 2: Pytania rekrutacyjne
- Gotowe pytania do kandydatów
- Pole na notatki i obserwacje z rozmowy

## Struktura pliku

```
candidate-evaluation-spreadsheet/
├── Ocena_Kandydatow.py          # Generator arkusza Excel
├── run_generator.py              # Skrypt runner
├── .gitignore                    # Ignoruj wygenerowane pliki
└── README.md                     # Ten plik
```

## Wymagania

- Python 3.6+
- openpyxl

## Licencja

Projekt dostępny do użytku wewnątrz organizacji.

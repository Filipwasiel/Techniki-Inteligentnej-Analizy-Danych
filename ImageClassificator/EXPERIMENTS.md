# Uruchamianie Eksperymentów

Projekt zawiera automatyczny system do uruchamiania eksperymentów z 4 różnymi podziałami zbioru testowego.

## 📊 Czym są eksperymenty?

System uruchamia 4 eksperymentów klasyfikacji z następującymi podziałami danych:

| Eksperyment | Trenowanie | Test |
|-----------|-----------|------|
| 1 | 60% | 40% |
| 2 | 70% | 30% |
| 3 | 80% | 20% |
| 4 | 90% | 10% |

## 🚀 Uruchomienie wszystkich eksperymentów

Z katalogu głównego projektu:

```bash
python run_experiments.py
```

System automatycznie:
1. Przygotowuje dane z odpowiednich podziałów
2. Trenuje model CNN dla każdego podziału
3. Zapisuje wyniki do `results/experiments_YYYYMMDD_HHMMSS/`
4. Przechodzi do następnego eksperymentu

## 📁 Struktura wyników

Po uruchomieniu eksperymentów, zobaczysz strukturę:

```
results/
└── experiments_20260506_123456/           # Timestamp serii
    ├── SUMMARY.txt                        # Podsumowanie wszystkich eksperymentów
    ├── split_60_40/                       # Eksperyment 1 (60% train, 40% test)
    │   ├── confusion_matrix.png           # Macierz pomyłek (wizualizacja)
    │   ├── accuracy_plot.png              # Wykres accuracy (trenowanie vs walidacja)
    │   ├── metrics.json                   # Metryki w formacie JSON
    │   └── classification_report.txt      # Raport klasyfikacji (tekst)
    ├── split_70_30/                       # Eksperyment 2 (70% train, 30% test)
    │   ├── confusion_matrix.png
    │   ├── accuracy_plot.png
    │   ├── metrics.json
    │   └── classification_report.txt
    ├── split_80_20/                       # Eksperyment 3 (80% train, 20% test)
    │   ├── confusion_matrix.png
    │   ├── accuracy_plot.png
    │   ├── metrics.json
    │   └── classification_report.txt
    └── split_90_10/                       # Eksperyment 4 (90% train, 10% test)
        ├── confusion_matrix.png
        ├── accuracy_plot.png
        ├── metrics.json
        └── classification_report.txt
```

## 📊 Czym są poszczególne pliki wyników?

### confusion_matrix.png
- Wizualizacja macierzy pomyłek
- Pokazuje liczbę prawidłowych i błędnych klasyfikacji
- Klasy: Cats (0) i Dogs (1)

### accuracy_plot.png
- Krzywe uczenia modelu
- Niebieska linia: dokładność na zbiorze treningowym
- Pomarańczowa linia: dokładność na zbiorze walidacyjnym
- Pomaga wykryć overfitting

### metrics.json
Zawiera metryki w formacie JSON:
```json
{
  "split_info": "split_60_40",
  "accuracy": 0.8234,
  "precision": 0.8456,
  "recall": 0.7890,
  "f1_score": 0.8156,
  "confusion_matrix": [[...], [...]]
}
```

### classification_report.txt
Tekstowy raport z:
- Precision, Recall, F1-score dla każdej klasy
- Średnie wartości (macro, weighted)
- Macierz pomyłek w formie tabeli

## 🔧 Uruchomienie jednego eksperymentu

Jeśli chcesz uruchomić tylko jeden eksperyment z konkretnym podziałem:

```python
from src.main import main

# Eksperyment z 70% treningu, 30% testu
model, history = main(train_split=0.7, output_dir="results/custom_experiment")
```

## ⚙️ Konfiguracja

Parametry modelu znajdują się w `src/config.py`:
- `IMG_SIZE`: Rozmiar obrazów (128, 128)
- `BATCH_SIZE`: Rozmiar batcha (64)
- `EPOCHS`: Liczba epok trenowania (20)

## 📝 Notatki

- Każdy eksperyment przechowuje niezależny model i wyniki
- Dane są losowo tasowane dla każdego podziału
- Wykresy są zapisywane w rozdzielczości 100 DPI
- Wszystkie wyniki mają polskie opisy
- System automatycznie tworzy folder z timestampem, by nie nadpisywać starych wyników

## 🎯 Cel eksperymentów

Porównanie wydajności modelu przy różnych proporcjach danych treningowych:
- Mniej danych treningowych = potencjalnie niższa dokładność, ale lepszy test rzeczywisty
- Więcej danych treningowych = wyższa dokładność, ale dane testowe mogą być niereprezentacyjne

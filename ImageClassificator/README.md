# Image Classificator

Binary image classification (Cats vs Dogs) z **automatycznym systemem eksperymentów** obsługującym wiele modeli i różne podziały danych.

---

## 📋 Wymagania Systemowe

- **Python:** 3.10
- **zainstalowane zależności**
---

## 🚀 Instalacja i Uruchomienie

### 1. Przygotowanie Środowiska

#### Windows
```bash
# Klon repozytorium
git clone <repository-url>
cd ImageClassificator

# Tworzenie virtual environment
py -3.10 -m venv venv  
venv\Scripts\activate

# Instalacja zależności
pip install -r requirements.txt
```

#### macOS/Linux
```bash
git clone <repository-url>
cd ImageClassificator

python -3.10 -m venv venv  
source venv/bin/activate

pip install -r requirements.txt
```

### 2. Uruchomienie Eksperymentów

```bash
python -m src.main
```

**Co się stanie:**
1. Automatycznie pobierze dataset z Kaggle
2. Przygotuje dane (weryfikacja obrazów, split)
3. Uruchomi eksperymenty z konfiguracją z `src/main.py`
4. Zapisze wyniki do `results/experiments_YYYYMMDD_HHMMSS/`

---
## 📊 Dataset

### Source

The dataset is obtained from **Kaggle**. You can download it from:
- **Kaggle Dataset URL**: https://www.kaggle.com/datasets/shaunthesheep/microsoft-catsvsdogs-dataset

## 🎯 Jak Testować Różne Modele i Podziały

Edytuj plik `src/main.py`, funkcja `main()` (linia ~182):

```python
def main():
    """Main entry point"""
    print(f"\nGPU dostępne: {tf.config.list_physical_devices('GPU')}")
    
    # ⬇️ ZMIEŃ TUTAJ ⬇️
    experiments = {
        'model_names': ['simple_cnn', 'mobilenet', 'resnet'],  # Wybrane modele
        'train_splits': [0.7, 0.8, 0.9],                       # Podziały danych
        'results_dir': 'results'
    }
    
    results = run_experiments(**experiments)
    return results
```

### Dostępne Modele

```
'simple_cnn'  # Mały CNN (3 warstwy Conv) - szybki
'mobilenet'   # Lekki model - dobry balans szybkości i dokładności
'resnet'      # Ciężki model - najwyższa dokładność
```

### Dostępne Podziały

```python
'train_splits': [0.3, 0.5, 0.7, 0.9]  # Procent danych treningowych
```

Np. `0.7` = 70% train, 30% test

---

## 📁 Struktura Projektu

```
ImageClassificator/
├── src/                          # Kod źródłowy
│   ├── config.py                 # Konfiguracja: IMG_SIZE, BATCH_SIZE, EPOCHS
│   ├── main.py                   # ⭐ Eksperymentator - tutaj zmienia się co testować
│   ├── models_factory.py         # ⭐ Wszystkie modele (SimpleCNN, MobileNet, ResNet)
│   ├── data_manager.py           # Pobieranie i split danych
│   ├── data_loader.py            # TensorFlow Dataset (resize, normalizacja)
│   └── evaluate.py               # Zapis wyników (PNG, JSON, TXT)
│
├── data/                         # 📥 Dane (tworzone automatycznie)
│   ├── raw/                      # Oryginalne zdjęcia (pobrane z Kaggle)
│   ├── train/                    # Trenowanie
│   │   ├── Cat/
│   │   └── Dog/
│   └── test/                     # Test
│       ├── Cat/
│       └── Dog/
│
├── results/                      # 📊 Wyniki eksperymentów (tworzone automatycznie)
│   └── experiments_20250506_123456/
│       ├── SUMMARY.txt
│       ├── simple_cnn_split_70_30/
│       │   ├── confusion_matrix.png
│       │   ├── accuracy_plot.png
│       │   ├── metrics.json
│       │   └── classification_report.txt
│       ├── mobilenet_split_70_30/
│       │   └── [same files]
│       └── resnet_split_70_30/
│           └── [same files]
│
├── requirements.txt              # Zależności Python
├── README.md                     # Dokumentacja
└── STRUCTURE.md                  # Szczegółowa dokumentacja struktury
```

---

## 📊 Struktura Wyników

Każdy eksperyment tworzy folder `<model_name>_split_<train>_<test>/` zawierający:

### confusion_matrix.png
Macierz pomyłek - wizualizacja gdzie model się myli:
- **True Positives (górny lewy):** Prawidłowe klasyfikacje klasy 0 (Koty)
- **False Negatives (górny prawy):** Błędy klasy 0
- **False Positives (dolny lewy):** Błędy klasy 1
- **True Negatives (dolny prawy):** Prawidłowe klasyfikacje klasy 1 (Psy)

### accuracy_plot.png
Krzywe uczenia - jak model się uczy:
- **Niebieska linia:** Dokładność na zbiorze treningowym
- **Pomarańczowa linia:** Dokładność na zbiorze walidacyjnym
- Jeśli się rozchodzą = overfitting

### metrics.json
Metryki w formacie JSON:
```json
{
  "accuracy": 0.856,      # Procent prawidłowych klasyfikacji
  "precision": 0.871,     # Jakość pozytywnych predykcji
  "recall": 0.823,        # Ile pozytywnych przypadków znaleziono
  "f1_score": 0.846       # Średnia ważona precision i recall
}
```

### classification_report.txt
Raport tekstowy dla każdej klasy:
```
             precision    recall  f1-score   support

       Cats       0.82      0.85      0.83       250
       Dogs       0.90      0.88      0.89       250

    accuracy                           0.86       500
```

### SUMMARY.txt
Ogólne podsumowanie wszystkich eksperymentów z tego przebiegu.

---

## ⚙️ Konfiguracja

### src/config.py

```python
TRAIN_DIR = '../data/train'    # Ścieżka do danych treningowych
TEST_DIR = '../data/test'      # Ścieżka do danych testowych
IMG_SIZE = (128, 128)          # Rozmiar obrazu
BATCH_SIZE = 64                # Liczba obrazów na batch
EPOCHS = 20                    # Liczba epok trenowania
```

### src/main.py (eksperymenty)

```python
experiments = {
    'model_names': ['simple_cnn'],        # Które modele testować
    'train_splits': [0.6, 0.7, 0.8, 0.9], # Jakie podziały danych
    'results_dir': 'results'               # Gdzie zapisywać wyniki
}
```

---

## 🧠 Dostępne Modele

### SimpleCNN ✅
- **Rozmiar:** Mały (~2M parametrów)
- **Szybkość:** Bardzo szybki
- **Dokładność:** Przeciętna (60-75%)
- **Opis:** 3 warstwy konwolucyjne + Dense layers

### MobileNet ✅
- **Rozmiar:** Średni (~4M parametrów)
- **Szybkość:** Szybki
- **Dokładność:** Dobra (75-85%)
- **Opis:** Lekki model bazujący na MobileNetV2

### ResNet ✅
- **Rozmiar:** Duży (~25M parametrów)
- **Szybkość:** Wolniejszy
- **Dokładność:** Bardzo dobra (85-92%)
- **Opis:** ResNet50 z residual connections

---

## 🔧 Dodatkowe Polecenia

### Sprawdzenie dostępnych modeli

```bash
python -c "from src.models_factory import list_available_models; print(list_available_models())"
```

### Usuwanie starych wyników

```bash
rm -rf results/  # macOS/Linux
rmdir /s results # Windows
```

### Czyszczenie cache

```bash
rm -rf src/__pycache__  # macOS/Linux
rmdir /s src\__pycache__ # Windows
```

---

## 🚨 Rozwiązywanie Problemów

### Problem: "ModuleNotFoundError: No module named 'tensorflow'"

```bash
# Zainstaluj znowu requirements
pip install --upgrade -r requirements.txt
```

### Problem: GPU nie jest wykrywane

```python
# W src/main.py, linia ~179
print(f"\nGPU dostępne: {tf.config.list_physical_devices('GPU')}")
# Jeśli puste - TensorFlow będzie używać CPU (wolniejsze ale działa)
```

**Ostatnia aktualizacja:** 2026-05-06

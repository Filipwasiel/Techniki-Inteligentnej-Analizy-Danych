Zadanie 4

Zaproponuj modyfikację jednego z algorytmów: GGL-PSOD lub SABOA zwiększającą różnorodność roju/poprawiającą skuteczność metody.  (Algorytmy zostały opisane w załączonych poniżej artykułach.) Zbadaj skuteczność zaproponowanej metody. W tym celu:

dobierz odpowiednio parametry (zakres poszukiwań, wymiar D, liczbę iteracji…).
 przeprowadź odpowiednie badania na standardowych zestawach funkcji, takich jak np. CEC 2013/CEC 2017.
przeanalizuj/porównaj wyniki z GGL-PSOD (lub SABOA).
przeprowadź testy (np. Test sumy rang Wilcoxona (Wilcoxon signed-rank test) lub t-studenta) skuteczności metody potwierdzające wyższość zaproponowanego wariantu.
Badania należy przeprowadzić dla min 10 funkcji. Wyniki należy uśrednić (aby były wiarygodne zwykle przeprowadza się co najmniej 51 przebiegów  algorytmu dla jednej funkcji.)

Sprawozdanie powinno zawierać opis zaproponowanej modyfikacji algorytmu, wyniki badań, wykresy, wyniki testu skuteczności, wnioski.
---

## 📋 Wymagania Systemowe

- **Python:** 3.14
- **zainstalowane zależności**
---

## 🚀 Instalacja i Uruchomienie

### 1. Przygotowanie Środowiska

#### Windows
```bash
# Klon repozytorium
git clone <repository-url>
cd AlgorythmGGL_PSOD

# Tworzenie virtual environment
py -3.14 -m venv venv  
venv\Scripts\activate

# Instalacja zależności
pip install -r requirements.txt
```
### 3. Do obsługi cec2017 wykorzystano biblioteke cec2017-py - pobierane jest automatycznie przy instalacji zależności z pliku ```requirements.txt```, więc ten krok nalezy pominąć, chyba że automatyczna instalacja nie działa. 
```bash
git clone https://github.com/tilleyd/cec2017-py
cd cec2017-py
pip install setuptools
# Zainstaluj (upewnij się, że masz aktywny venv)
python setup.py install
```

### 2. Uruchomienie Eksperymentów

```bash
python -m src.main
```

---


chce żebyśzrobił żeby klasa GGL_psod była główna, i miała 2 klasy pochodne - piersza surowa, druga modyfikowana (jak na razie żeby nie było implementacji, dostosuj do tego maina, chce zeby był na tyle inteligentny, żeby przy wpisaniu mniejszej ilości przebiegów, realizował sie do wykonania tylu przebiegów ile musi, a nie był zapetlony - aktualny main jest 
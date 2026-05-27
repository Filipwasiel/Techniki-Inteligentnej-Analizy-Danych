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
## Opis Modyfikacji w `GGL_PSOD_Modified`

Wersja zmodyfikowana `GGL_PSOD_Modified` rozszerza klasyczny, dwuwarstwowy algorytm GGL-PSOD o dwa komplementarne mechanizmy ukierunkowane na **zwiększenie różnorodności populacji** oraz **zapobieganie przedwczesnej zbieżności** w minimach lokalnych. 

Modyfikacja wprowadza inteligentne, selektywne zarządzanie rojem zarówno w warstwie genetycznej (zarządzanie pamięcią wzorców), jak i w warstwie rojowej (równanie prędkości PSO).

---

### 1. Uwarunkowane Uczenie Negatywne (Warstwa PSO)
W przeciwieństwie do klasycznego podejścia *repulsion*, które odpycha cząstki bezwarunkowo i wprowadza chaos, w zmodyfikowanym algorytmie wektor odpychania aktywuje się w sposób wysoce selektywny na podstawie trzech filtrów:

* **Uwarunkowanie Przestrzenne (Strefa Zagrożenia):** Algorytm wyznacza środek ciężkości trzech najgorszych cząstek w populacji (`wbest_zone`). Siła odpychająca działa na cząstkę wyłącznie w wymiarach, w których odległość od tej strefy jest mniejsza niż **15% szerokości dziedziny poszukiwań** (`0.15 * (ub - lb)`).
* **Uwarunkowanie Hierarchiczne (Status Marudera):** Mechanizm odpychania **nie dotyczy liderów roju**. Aktywuje się on wyłącznie dla cząstek, których historyczne przystosowanie ($pbest\_fit$) jest gorsze niż aktualna średnia populacji (`mean_fitness`).
* **Dynamiczna Adaptacja Siły (`c_bad`):** Intensywność odpychania sterowana jest współczynnikiem $c_{bad}$, który maleje liniowo wraz z postępem iteracji (od $1.5$ do $0.1$). Zapewnia to silną dywersyfikację na początku i stabilną zbieżność pod koniec działania algorytmu.

**Zmodyfikowane równanie prędkości cząstki:**
$$V_{i} = w \cdot V_{i} + c_1 \cdot r_1 \cdot (E_{i} - X_{i}) + c_2 \cdot r_2 \cdot (gbest - X_{i}) + V_{repulsion}$$

Gdzie składowe $V_{repulsion}$ dla każdego wymiaru wyliczane są jako:
`- c_bad * r3 * (wbest_zone - X[i])` (gdy spełnione są powyższe warunki uwarunkowania).

---

### 2. Adaptacyjna Mutacja Wzorców DMS (Warstwa Genetyczna)
Modyfikacja dotyczy również mechanizmu krzyżowania i mutacji przy wystąpieniu stagnacji. W klasycznym GGL-PSOD prawdopodobieństwo mutacji wzorca jest stałe i wynosi $p_m = 0.01$. W wersji zmodyfikowanej wdrożono mechanizm **Dynamic Mutation Scaling (DMS)**:

* Wraz ze wzrostem licznika stagnacji danej cząstki (`stagnation_counter`), bazowe prawdopodobieństwo mutacji rośnie adaptacyjnie (maksymalnie do poziomu $0.15$).
* Dzięki temu cząstki efektywne zachowują niską mutację (precyzyjna eksploatacja obszaru), natomiast cząstki uwięzione w minimach lokalnych zyskują znacznie wyższą szansę na losową dywersyfikację struktury genów jeszcze przed ostatecznym resetem turniejowym.

---

### Podsumowanie korzyści
Połączenie uwarunkowanego odpychania maruderów w warstwie fizycznej (PSO) oraz adaptacyjnego skalowania mutacji w warstwie informacyjnej (GA) pozwala algorytmowi zachować wysoką elastyczność i skutecznie uciekać z pułapek optymalizacyjnych w problemach wielimodalnych, hybrydowych oraz kompozytowych (co potwierdzają wyniki testów statystycznych Wilcoxona).


## Opis Modyfikacji w `GGL_PSOD_Modified` - stare 

Klasa `GGL_PSOD_Modified` wprowadza trzy główne usprawnienia w stosunku do bazowej, surowej wersji algorytmu `GGL_PSOD_Raw`. Ich nadrzędnym celem jest **zwiększenie różnorodności roju**, **zapobieganie przedwczesnej zbieżności do lokalnych minimów** oraz **poprawa zbieżności globalnej**.

### 1. Dynamiczna i Nieliniowa Aktualizacja Parametrów (Nieliniowa Waga Bezwładności)
W standardowym algorytmie `GGL_PSOD_Raw` współczynnik bezwładności $w$ maleje liniowo. W wersji modyfikowanej zastosowano **nieliniowy spadek kwadratowy**:
```python
w = 0.9 - 0.5 * (iter / self.max_iter)**2
```
* **Działanie:** W początkowych iteracjach spadek wartości $w$ zachodzi znacznie wolniej niż w przypadku liniowym.
* **Cel:** Zapewnia to dłuższą i bardziej intensywną eksplorację (przeszukiwanie przestrzeni w skali makro) na początku działania algorytmu, a następnie szybkie przejście do precyzyjnej eksploatacji (lokalnego dostrajania rozwiązań) pod koniec procesu optymalizacji.

### 2. Mutacja z Rozkładem Cauchy'ego (Cauchy Mutation)
W warstwie genetycznej (operacja mutacji egzemplarza) zastąpiono mutację jednostajną mutacją opartą na **rozkładzie Cauchy'ego**:
```python
O_i[d] += np.random.standard_cauchy()
O_i[d] = np.clip(O_i[d], self.lb, self.ub)
```
* **Działanie:** Zamiast generowania nowej losowej wartości w pełnym zakresie przestrzeni poszukiwań, modyfikowana jest dotychczasowa pozycja egzemplarza.
* **Cel:** Rozkład Cauchy'ego charakteryzuje się grubymi ogonami (często generuje wartości bliskie zero, ale z niezerowym prawdopodobieństwem zwraca bardzo duże wartości). Dzięki temu mutacja najczęściej wykonuje małe, precyzyjne kroki lokalne, lecz od czasu do czasu pozwala na wykonanie długiego skoku ("długiego kroku"), co umożliwia efektywne uciekanie z lokalnych minimów.

### 3. Elitarna, Zbalansowana Selekcja Turniejowa (Balanced Tournament Selection)
W przypadku wykrycia stagnacji egzemplarza (brak poprawy dopasowania przez zadany próg `sz` iteracji), wywoływana jest dedykowana procedura `_handle_stagnation`:
1. Losowana jest podgrupa uczestników (20% populacji).
2. Dla każdego uczestnika obliczana jest odległość euklidesowa od aktualnego lidera globalnego $Gbest$:
   $$\text{distance}_j = \| Pbest_j - Gbest \|_2$$
3. Wartości funkcji dopasowania ($fitness$) oraz obliczone odległości są normalizowane do przedziału $[0, 1]$.
4. Wyznaczana jest ocena zbalansowana każdego uczestnika według wzoru:
   $$\text{score}_j = 0.6 \cdot \text{norm\_fits}_j - 0.4 \cdot \text{norm\_dists}_j$$
5. Egzemplarz stagnujący jest zastępowany przez uczestnika o **najmniejszej** wartości $\text{score}_j$.
* **Cel:** Taki dobór współczynników (waga dodatnia dla znormalizowanego dopasowania oraz waga ujemna dla znormalizowanej odległości) premiuje cząstki, które posiadają dobrą jakość (niskie $fitness$), a jednocześnie są położone **jak najdalej** od aktualnie najlepszego rozwiązania globalnego $Gbest$. Zwiększa to różnorodność roju i bezpośrednio przeciwdziała zjawisku przedwczesnej konwergencji populacji w jednym punkcie.
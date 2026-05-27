import numpy as np

class GGL_PSOD:
    def __init__(self, obj_func, dim, ps=50, lb=-100, ub=100, func_idx=None, run_id=None):
        self.max_fes = 10000*dim
        self.obj_func = obj_func        # funkcja celu
        self.dim = dim                  # wymiarowość - określa z ilu zmiennych składa sie jedno rozwiązanie
        self.ps = ps                    # wielkość roju cząstek i egzemplarzy GA
        self.max_iter = self.max_fes // ps   # maksymalna liczba iteracji 
                                             # max_fes - całkowity limit wywołań funkcji celu
        self.lb = lb                    # dolny zakres przeszukiwania
        self.ub = ub                    # górny zakres przeszukiwania
        self.func_idx = func_idx        # indeks funkcji testowej
        self.run_id = run_id            # numer przebiegu
        
        # Parametry algorytmu GA i PSO
        self.pm = 0.01  # Prawdopodobieństwo mutacji
        self.sz = 7     # Próg stagnacji egzemplarza - bez poprawy 7 razy jest selektywne losowanie turniejowe
        
        # Inicjalizacja populacji
        self.X = np.random.uniform(lb, ub, (ps, dim))
        self.V = np.zeros((ps, dim))
        
        # Inicjalizacja Pbest i Gbest
        self.pbest = np.copy(self.X)
        self.pbest_fit = np.array([obj_func(x) for x in self.X])
        
        self.gbest_idx = np.argmin(self.pbest_fit)
        self.gbest = np.copy(self.pbest[self.gbest_idx])
        self.gbest_fit = self.pbest_fit[self.gbest_idx]
        
        # Inicjalizacja egzemplarzy (E) i liczników stagnacji
        self.E = np.copy(self.pbest)
        self.e_fit = np.copy(self.pbest_fit)
        self.stagnation_counter = np.zeros(ps)
        
        # --- NOWOŚĆ: Inicjalizacja tablicy na historię zbieżności ---
        self.history = []

    def run(self):
        raise NotImplementedError("Metoda run() musi zostać zaimplementowana w klasie pochodnej.")


class GGL_PSOD_Raw(GGL_PSOD):
    def run(self):
        # Czyszczenie historii przed nowym przebiegiem
        self.history = []
        
        for iter in range(self.max_iter):
            # Liniowa aktualizacja parametrów wg równań PSO
            w = 0.9 - (iter / self.max_iter) * (0.9 - 0.4) # bezwładność 
            c1 = 2.5 - (iter / self.max_iter) * (2.5 - 0.5) # poznawczy - do własnego
            c2 = 0.5 + (iter / self.max_iter) * (2.5 - 0.5) # przyciąganie do globalnego najlepszego
            
            for i in range(self.ps):
                # --- WARSTWA GENETYCZNA (Egzemplarze) ---

                # Krzyżowanie w topologii pierścieniowej
                n_i1 = i - 1 if i > 0 else self.ps - 1
                n_i2 = i + 1 if i < self.ps - 1 else 0
                
                O_i = np.zeros(self.dim)
                for d in range(self.dim):
                    k = np.random.randint(0, self.ps)
                    if self.pbest_fit[i] < self.pbest_fit[k]:
                        # Uczenie od sąsiadów z pierścienia
                        r_d = np.random.rand()
                        O_i[d] = r_d * self.pbest[n_i1, d] + (1 - r_d) * self.pbest[n_i2, d]
                    else:
                        # Uczenie od losowego Pbest
                        O_i[d] = self.pbest[k, d]
                
                # Mutacja 
                for d in range(self.dim):
                    if np.random.rand() < self.pm:
                        O_i[d] = np.random.uniform(self.lb, self.ub)
                
                # Selekcja egzemplarza 
                o_fit = self.obj_func(O_i)
                if o_fit < self.e_fit[i]:
                    self.E[i] = O_i
                    self.e_fit[i] = o_fit
                    self.stagnation_counter[i] = 0
                    if o_fit < self.gbest_fit:
                        self.gbest_fit = o_fit
                        self.gbest = np.copy(O_i)
                else:
                    self.stagnation_counter[i] += 1
                
                # Re-selekcja przy stagnacji (Tournament selection) 
                if self.stagnation_counter[i] >= self.sz:
                    participants = np.random.choice(self.ps, int(0.2 * self.ps), replace=False)
                    best_participant = participants[np.argmin(self.pbest_fit[participants])]
                    self.E[i] = np.copy(self.pbest[best_participant])
                    self.e_fit[i] = self.pbest_fit[best_participant]
                    self.stagnation_counter[i] = 0

                # --- WARSTWA PSO (Aktualizacja cząstki) ---
                # Równanie prędkości 
                r1, r2 = np.random.rand(self.dim), np.random.rand(self.dim)
                self.V[i] = (
                        w * self.V[i]
                        + c1 * r1 * (self.E[i] - self.X[i])
                        + c2 * r2 * (self.gbest - self.X[i])
                        )
                
                # Aktualizacja pozycji 
                self.X[i] = self.X[i] + self.V[i]
                self.X[i] = np.clip(self.X[i], self.lb, self.ub) # Granice
                
                # Ocena i aktualizacja Pbest oraz Gbest
                current_fit = self.obj_func(self.X[i])
                if current_fit < self.pbest_fit[i]:
                    self.pbest_fit[i] = current_fit
                    self.pbest[i] = np.copy(self.X[i])
                    
                    if current_fit < self.gbest_fit:
                        self.gbest_fit = current_fit
                        self.gbest = np.copy(self.X[i])
            
            # --- NOWOŚĆ: Zapis aktualnego gbest_fit na koniec iteracji ---
            self.history.append(self.gbest_fit)
            
            if (iter + 1) % 100 == 0:
                print(f"F{self.func_idx} Run{self.run_id+1} - Iteracja {iter+1}/{self.max_iter}, Najlepszy wynik: {self.gbest_fit:.5e}")
                
        # Zwracanie wzbogacone o tablicę history
        return self.gbest, self.gbest_fit, self.history


import numpy as np

class GGL_PSOD_Modified(GGL_PSOD):
    def run(self):
        """
        Zoptymalizowana wersja GGL-PSOD wyposażona w Dual-Ring Topology (Modyfikacja 1)
        oraz Pamięć Tabu Stagnacji (Modyfikacja 2). Usunięto mechanizm Shake.
        """
        # Czyszczenie historii przed nowym przebiegiem
        self.history = []
        
        # --- STRUKTURY DLA MODYFIKACJI 2 (PAMIĘĆ TABU) ---
        # Bufor przechowujący pozycje gbest z ostatnich 100 iteracji w celu wykrycia najgłębszej pułapki
        gbest_position_history = []
        tabu_zone = np.zeros(self.dim)
        
        # Inicjalizacja zmiennych do monitorowania globalnej stagnacji roju (Global Stagnation Filter)
        best_gbest_ever = self.gbest_fit
        global_stagnation = 0
        
        for iter in range(self.max_iter):
            # 1. Liniowa aktualizacja podstawowych parametrów GGL-PSOD (eq. 13-15)
            w = 0.9 - (iter / self.max_iter) * (0.9 - 0.4)
            c1 = 2.5 - (iter / self.max_iter) * (2.5 - 0.5)
            c2 = 0.5 + (iter / self.max_iter) * (2.5 - 0.5)
            
            # Wykładnicze wygaszanie siły odpychania c_bad
            c_bad = 0.5 * np.exp(-4.0 * (iter / self.max_iter))
            
            # Aktualizacja licznika globalnego braku postępu całego roju
            if self.gbest_fit < best_gbest_ever:
                best_gbest_ever = self.gbest_fit
                global_stagnation = 0
            else:
                global_stagnation += 1
            
            # --- WDROŻENIE MODYFIKACJI 2: AKTUALIZACJA PAMIĘCI TABU ---
            # Zapisujemy aktualną pozycję lidera roju
            gbest_position_history.append(np.copy(self.gbest))
            if len(gbest_position_history) > 100:
                gbest_position_history.pop(0) # Trzymamy okno przesuwne o szerokości maksymalnie 100 wpisów
            
            # Strefą Tabu staje się średnia pozycja lidera w tym oknie. Jeśli rój stoi w miejscu,
            # tabu_zone precyzyjnie namierza środek pułapki (lokalnego minimum).
            tabu_zone = np.mean(gbest_position_history, axis=0)
            
            # Wyznaczenie parametrów pomocniczych do selekcji i odpychania
            mean_fitness = np.mean(self.pbest_fit)
            max_dim_dist = self.ub - self.lb
            
            # Pobieramy indeksy 3 najlepszych osobników w roju dla Modyfikacji 1
            best_3_indices = np.argsort(self.pbest_fit)[:3]
            
            # Główna pętla po populacji cząstek
            for i in range(self.ps):
                
                # --- WDROŻENIE MODYFIKACJI 1: WARSTWA GENETYCZNA (Dual-Ring Topology) ---
                n_i1 = i - 1 if i > 0 else self.ps - 1
                n_i2 = i + 1 if i < self.ps - 1 else 0
                
                O_i = np.zeros(self.dim)
                for d in range(self.dim):
                    # Warunek hierarchiczny: jeśli cząstka jest słaba (gorsza od średniej roju),
                    # pozwalamy jej krzyżować geny bezpośrednio z elitą roju (jednym z TOP 3)
                    if self.pbest_fit[i] > mean_fitness:
                        k_elite = np.random.choice(best_3_indices)
                        r_d = np.random.rand()
                        # Hybrydowe krzyżowanie: wiedza topologii pierścienia + zastrzyk genów elity
                        O_i[d] = r_d * self.pbest[n_i1, d] + (1 - r_d) * self.pbest[k_elite, d]
                    else:
                        # Jeśli cząstka jest dobra (lepsza od średniej), zachowujemy standardowy GL-PSO Ring Topology
                        k = np.random.randint(0, self.ps)
                        if self.pbest_fit[i] < self.pbest_fit[k]:
                            r_d = np.random.rand()
                            O_i[d] = r_d * self.pbest[n_i1, d] + (1 - r_d) * self.pbest[n_i2, d]
                        else:
                            O_i[d] = self.pbest[k, d]
                
                # Mutacja egzemplarza (standard GL-PSO) - kluczowa do utrzymania różnorodności przy elicie
                for d in range(self.dim):
                    if np.random.rand() < self.pm:
                        O_i[d] = np.random.uniform(self.lb, self.ub)
                
                # Selekcja egzemplarza i sprawdzanie stagnacji
                o_fit = self.obj_func(O_i)
                if o_fit < self.e_fit[i]:
                    self.E[i] = O_i
                    self.e_fit[i] = o_fit
                    self.stagnation_counter[i] = 0
                    if o_fit < self.gbest_fit:
                        self.gbest_fit = o_fit
                        self.gbest = np.copy(O_i)
                else:
                    self.stagnation_counter[i] += 1
                
                # Re-selekcja przy stagnacji (Turniej 20% Ps)
                if self.stagnation_counter[i] >= self.sz:
                    participants = np.random.choice(self.ps, int(0.2 * self.ps), replace=False)
                    best_participant = participants[np.argmin(self.pbest_fit[participants])]
                    self.E[i] = np.copy(self.pbest[best_participant])
                    self.e_fit[i] = self.pbest_fit[best_participant]
                    self.stagnation_counter[i] = 0

                # --- WARSTWA PSO (Aktualizacja cząstki z Odpychaniem od Strefy Tabu) ---
                r1 = np.random.rand(self.dim)
                r2 = np.random.rand(self.dim)
                r3 = np.random.rand(self.dim)
                
                repulsion_vector = np.zeros(self.dim)
                
                # Aktywacja uwarunkowana: odpychamy maruderów, gdy cały rój przeżywa kryzys (global_stagnation >= 5)
                if global_stagnation >= 5 and self.stagnation_counter[i] >= 3 and self.pbest_fit[i] > mean_fitness:
                    
                    # Dynamicznie malejący promień strefy zakazanej (od 5% do 0.5% dziedziny)
                    current_zone_ratio = 0.05 * (1.0 - iter / self.max_iter)
                    
                    for d in range(self.dim):
                        # MODYFIKACJA 2: Sprawdzenie dystansu do STREFY TABU (a nie do skaczących najgorszych cząstek)
                        if abs(self.X[i, d] - tabu_zone[d]) < current_zone_ratio * max_dim_dist:
                            
                            # Stabilne odpychanie o charakterze kierunkowym od historycznego punktu utknięcia
                            direction_to_gbest = np.sign(self.gbest[d] - self.X[i, d])
                            repulsion_force = - c_bad * r3[d] * (tabu_zone[d] - self.X[i, d])
                            
                            repulsion_vector[d] = repulsion_force + (0.05 * c_bad * direction_to_gbest * max_dim_dist)
                
                # Zbalansowane, zorientowane na sukces równanie prędkości cząstki
                self.V[i] = (
                    w * self.V[i]
                    + c1 * r1 * (self.E[i] - self.X[i])
                    + c2 * r2 * (self.gbest - self.X[i])
                    + repulsion_vector
                )
                
                # Aktualizacja pozycji cząstki i nałożenie ograniczeń dziedziny
                self.X[i] = self.X[i] + self.V[i]
                self.X[i] = np.clip(self.X[i], self.lb, self.ub)
                
                # Ocena nowej pozycji i aktualizacja Pbest, Gbest
                current_fit = self.obj_func(self.X[i])
                if current_fit < self.pbest_fit[i]:
                    self.pbest_fit[i] = current_fit
                    self.pbest[i] = np.copy(self.X[i])
                    if current_fit < self.gbest_fit:
                        self.gbest_fit = current_fit
                        self.gbest = np.copy(self.X[i])
            
            # Zapis aktualnego gbest_fit na koniec iteracji
            self.history.append(self.gbest_fit)
            
            # Monitorowanie postępu w konsoli
            if (iter + 1) % 100 == 0:
                print(f"F{self.func_idx} Run{self.run_id+1} - Iteracja {iter+1}/{self.max_iter}, Najlepszy wynik: {self.gbest_fit:.5e}")
                
        return self.gbest, self.gbest_fit, self.history
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
        Zoptymalizowana wersja GGL-PSOD wyposażona w Globalny Filtr Stagnacji
        oraz ukierunkowane, wykładniczo wygaszane uczenie negatywne.
        """
        # Czyszczenie historii przed nowym przebiegiem
        self.history = []
        
        # Inicjalizacja zmiennych do monitorowania globalnej stagnacji roju
        best_gbest_ever = self.gbest_fit
        global_stagnation = 0
        
        for iter in range(self.max_iter):
            # 1. Liniowa aktualizacja podstawowych parametrów GGL-PSOD (eq. 13-15)
            w = 0.9 - (iter / self.max_iter) * (0.9 - 0.4)
            c1 = 2.5 - (iter / self.max_iter) * (2.5 - 0.5)
            c2 = 0.5 + (iter / self.max_iter) * (2.5 - 0.5)
            
            # Aktualizacja licznika globalnego braku postępu
            if self.gbest_fit < best_gbest_ever:
                best_gbest_ever = self.gbest_fit
                global_stagnation = 0
            else:
                global_stagnation += 1
            
            # USPRAWNIENIE: Wykładnicze wygaszanie siły odpychania (funkcja exp)
            # W początkowej i środkowej fazie dynamicznego spadku siła maleje płynnie,
            # zapobiegając rozregulowaniu i powstawaniu opóźnień ("brzucha") na wykresie.
            c_bad = 0.5 * np.exp(-4.0 * (iter / self.max_iter))
            
            # 2. Wyznaczenie obszaru zakazanego (środek ciężkości 3 najgorszych cząstek)
            worst_3_indices = np.argsort(self.pbest_fit)[-3:]
            wbest_zone = np.mean(self.pbest[worst_3_indices], axis=0)
            
            # Średni fitness populacji do warunku hierarchicznego
            mean_fitness = np.mean(self.pbest_fit)
            max_dim_dist = self.ub - self.lb
            
            # Główna pętla po populacji cząstek
            for i in range(self.ps):
                
                # --- WARSTWA GENETYCZNA (Generowanie Egzemplarzy - Ring Topology) ---
                n_i1 = i - 1 if i > 0 else self.ps - 1
                n_i2 = i + 1 if i < self.ps - 1 else 0
                
                O_i = np.zeros(self.dim)
                for d in range(self.dim):
                    k = np.random.randint(0, self.ps)
                    if self.pbest_fit[i] < self.pbest_fit[k]:
                        r_d = np.random.rand()
                        O_i[d] = r_d * self.pbest[n_i1, d] + (1 - r_d) * self.pbest[n_i2, d]
                    else:
                        O_i[d] = self.pbest[k, d]
                
                # Mutacja egzemplarza (standard GL-PSO)
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

                # --- WARSTWA PSO (Aktualizacja cząstki z Inteligentnym Odpychaniem) ---
                r1 = np.random.rand(self.dim)
                r2 = np.random.rand(self.dim)
                r3 = np.random.rand(self.dim)
                
                repulsion_vector = np.zeros(self.dim)
                
                # MODYFIKACJA (Globalny Filtr Stagnacji): Odpychanie uruchamia się TYLKO wtedy,
                # gdy cały rój utknął (global_stagnation >= 5), dana cząstka stoi w miejscu (>= 3)
                # ORAZ jej wynik jest gorszy niż średnia populacji.
                if global_stagnation >= 5 and self.stagnation_counter[i] >= 3 and self.pbest_fit[i] > mean_fitness:
                    
                    # Dynamicznie malejący promień strefy zakazanej (od 5% do 0.5% szerokości dziedziny)
                    current_zone_ratio = 0.05 * (1.0 - iter / self.max_iter)
                    
                    for d in range(self.dim):
                        # Sprawdzenie dynamicznego warunku przestrzennego
                        if abs(self.X[i, d] - wbest_zone[d]) < current_zone_ratio * max_dim_dist:
                            
                            # Ukierunkowane odpychanie: ucieczka od strefy złej w stronę gbest
                            direction_to_gbest = np.sign(self.gbest[d] - self.X[i, d])
                            repulsion_force = - c_bad * r3[d] * (wbest_zone[d] - self.X[i, d])
                            
                            # Integracja impulsu repulsyjnego ze zwrotem ku globalnemu optimum
                            repulsion_vector[d] = repulsion_force + (0.05 * c_bad * direction_to_gbest * max_dim_dist)
                
                # Zbalansowane równanie prędkości cząstki
                self.V[i] = (
                    w * self.V[i]
                    + c1 * r1 * (self.E[i] - self.X[i])
                    + c2 * r2 * (self.gbest - self.X[i])
                    + repulsion_vector
                )
                
                # Aktualizacja pozycji cząstki
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
                print(f"F{self.func_idx} Run{self.run_id+1} - Iteracja {iter+1}/{self.max_iter}, Najlepszy wynik (Modyfikacja): {self.gbest_fit:.5e}")
                
        return self.gbest, self.gbest_fit, self.history
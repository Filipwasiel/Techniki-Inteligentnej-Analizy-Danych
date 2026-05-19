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
        
        # Parametry algorytmu GA i PS0
        self.pm = 0.01  # Prawdopodobieństwo mutacji
        self.sz = 7     # Próg stagnacji egzemplarza - bez poprawy 7 razy jest selekcja turniejowa poprzednich pozycji
        
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

    def run(self):
        raise NotImplementedError("Metoda run() musi zostać zaimplementowana w klasie pochodnej.")

class GGL_PSOD_Raw(GGL_PSOD):
    def run(self):
        for iter in range(self.max_iter):
            # Liniowa aktualizacja parametrów wg równań PSO
            w = 0.9 - (iter / self.max_iter) * (0.9 - 0.4) # bezwładność 
            c1 = 2.5 - (iter / self.max_iter) * (2.5 - 0.5) # poznawczy - do własnego
            c2 = 0.5 + (iter / self.max_iter) * (2.5 - 0.5) # przyciaganie do globalnego najlepszego
            
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
            
            if (iter + 1) % 100 == 0:
                print(f"F{self.func_idx} Run{self.run_id+1} - Iteracja {iter+1}/{self.max_iter}, Najlepszy wynik: {self.gbest_fit:.5e}")
                
        return self.gbest, self.gbest_fit

class GGL_PSOD_Modified(GGL_PSOD):
    def run(self):
        for iter in range(self.max_iter):
            # Modyfikaja 1 Dynamiczna i nieliniowa aktualizacja parametrów
            #   - po połowie iteracji przyspieszamy eksplorację, a potem eksploatację
            w = 0.9 - 0.5 * (iter / self.max_iter)**2 
            c1 = 2.5 - (iter / self.max_iter) * (2.5 - 0.5) 
            c2 = 0.5 + (iter / self.max_iter) * (2.5 - 0.5) 
            
            for i in range(self.ps):
                # --- WARSTWA GENETYCZNA (Egzemplarze) ---
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
                
                # Mutacja z rozkładem Cauchy'ego
                for d in range(self.dim):
                    if np.random.rand() < self.pm:
                        # Modyfikujemy obecną pozycję egzemplarza (drobne korekty + rzadkie skoki)
                        O_i[d] += np.random.standard_cauchy()
                        # Ograniczamy do przestrzeni poszukiwań
                        O_i[d] = np.clip(O_i[d], self.lb, self.ub)
                
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
                
                if self.stagnation_counter[i] >= self.sz:
                    self._handle_stagnation(i)

                # --- WARSTWA PSO (Aktualizacja cząstki) ---
                r1, r2 = np.random.rand(self.dim), np.random.rand(self.dim)
                self.V[i] = (
                        w * self.V[i]
                        + c1 * r1 * (self.E[i] - self.X[i])
                        + c2 * r2 * (self.gbest - self.X[i])
                        )
                
                self.X[i] = self.X[i] + self.V[i]
                self.X[i] = np.clip(self.X[i], self.lb, self.ub) 
                
                current_fit = self.obj_func(self.X[i])
                if current_fit < self.pbest_fit[i]:
                    self.pbest_fit[i] = current_fit
                    self.pbest[i] = np.copy(self.X[i])
                    
                    if current_fit < self.gbest_fit:
                        self.gbest_fit = current_fit
                        self.gbest = np.copy(self.X[i])
            
            if (iter + 1) % 100 == 0:
                print(f"F{self.func_idx} Run{self.run_id+1} - Iteracja {iter+1}/{self.max_iter}, Najlepszy wynik: {self.gbest_fit:.5e}")
                
        return self.gbest, self.gbest_fit
    
    def _handle_stagnation(self, i):
        """
        Elitarna, zbalansowana selekcja turniejowa.
        Wybiera cząstkę, która ma dobrą wartość fitness, ale jest odległa od lidera.
        """
        participants = np.random.choice(self.ps, int(0.2 * self.ps), replace=False)
        
        fits = self.pbest_fit[participants]
        positions = self.pbest[participants]
        distances = np.linalg.norm(positions - self.gbest, axis=1)
        
        fit_min, fit_max = np.min(fits), np.max(fits)
        dist_min, dist_max = np.min(distances), np.max(distances)
        
        norm_fits = (fits - fit_min) / (fit_max - fit_min + 1e-10)
        norm_dists = (distances - dist_min) / (dist_max - dist_min + 1e-10)
        
        scores = 0.6 * norm_fits - 0.4 * norm_dists
        
        winner_idx = participants[np.argmin(scores)]
        
        self.E[i] = np.copy(self.pbest[winner_idx])
        self.e_fit[i] = self.pbest_fit[winner_idx]
        self.stagnation_counter[i] = 0
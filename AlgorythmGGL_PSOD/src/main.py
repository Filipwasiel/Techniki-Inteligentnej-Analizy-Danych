import numpy as np
import time
import csv
import os
from concurrent.futures import ProcessPoolExecutor
from cec2017.functions import all_functions
from src.GGL_PSOD import GGL_PSOD_Raw, GGL_PSOD_Modified

def run_single_experiment(params):
    func_idx, dim, is_modified, run_id = params
    
    # Wybór funkcji celu
    target_func = all_functions[func_idx - 1]
    f_optimum = func_idx * 100
    
    def obj_func(x):
        return target_func(x.reshape(1, -1))[0]

    if is_modified:
        model = GGL_PSOD_Modified(
            obj_func=obj_func, 
            dim=dim
        )
    else:
        model = GGL_PSOD_Raw(
            obj_func=obj_func, 
            dim=dim
        )
    
    try:
        result = model.run()
        if isinstance(result, tuple):
            _, fitness = result
        else:
            fitness = result
            
        return fitness - f_optimum
    except NotImplementedError as e:
        # Zwracamy NaN, żeby pokazać brak wyniku bez rzucania wyjątku
        return np.nan

# 2. Zarządca eksperymentów
def perform_full_study():
    functions_to_test = [5] # numery funkcji dla których bedzie testowane
    algorithms = [False, True] 
    dim = 30
    num_runs = 3
    
    # Tworzymy folder na wyniki
    os.makedirs("results", exist_ok=True)
    
    with open("results/final_report.csv", "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Function", "Algorithm", "Mean_Error", "Std_Dev", "Best", "Worst"])

        for alg_mod in algorithms:
            alg_name = "Modified" if alg_mod else "Original"
            print(f"\n>>> TESTOWANIE ALGORYTMU: {alg_name}")
            
            for f_idx in functions_to_test:
                print(f"  Obliczanie F{f_idx} ({num_runs} biegów)... ", end="", flush=True)
                
                # Przygotowanie paczek danych dla procesów
                tasks = [(f_idx, dim, alg_mod, i) for i in range(num_runs)]
                
                # URUCHOMIENIE RÓWNOLEGŁE
                with ProcessPoolExecutor() as executor:
                    errors = list(executor.map(run_single_experiment, tasks))
                
                errors = np.array(errors)
                
                # Statystyki
                mean_e = np.mean(errors)
                std_e = np.std(errors)
                
                # Zapis do CSV
                writer.writerow([f"F{f_idx}", alg_name, mean_e, std_e, np.min(errors), np.max(errors)])
                print(f"Zakończono. Średni błąd: {mean_e:.2e}")

if __name__ == "__main__":
    # KONIECZNE NA WINDOWS dla multiprocessing!
    perform_full_study()
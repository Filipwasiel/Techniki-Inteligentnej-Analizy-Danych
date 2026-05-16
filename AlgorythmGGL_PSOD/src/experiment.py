import numpy as np
import csv
import os
from concurrent.futures import ProcessPoolExecutor
from cec2017.functions import all_functions

def run_single_experiment(params):
    func_idx, dim, AlgorithmClass, run_id = params
    
    target_func = all_functions[func_idx - 1]
    
    f_optimum = func_idx * 100
    
    def obj_func(x):
        return target_func(x.reshape(1, -1))[0]

    model = AlgorithmClass(
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
        return np.nan

def run_experiments(algorithms, functions_to_test, dim, num_runs):
    """
    Główny zarządca procesów. Zleca zadania na wolne wątki procesora
    i zapisuje statystyki końcowe błędu do pliku CSV.
    """
    os.makedirs("results", exist_ok=True)
    csv_path = "results/final_report.csv"
    
    with open(csv_path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Function", "Algorithm", "Mean_Error", "Std_Dev", "Best", "Worst"])

        for AlgorithmClass in algorithms:
            alg_name = AlgorithmClass.__name__
            print(f"\n>>> TESTOWANIE ALGORYTMU: {alg_name}")
            
            for f_idx in functions_to_test:
                print(f"  Obliczanie F{f_idx} ({num_runs} biegów)... ", end="", flush=True)
                
                tasks = [(f_idx, dim, AlgorithmClass, i) for i in range(num_runs)]
                
                workers = max(1, os.cpu_count() - 1)
                with ProcessPoolExecutor(max_workers=workers) as executor:
                    errors = list(executor.map(run_single_experiment, tasks))
                
                errors = np.array(errors)
                
                valid_errors = errors[~np.isnan(errors)]
                
                if len(valid_errors) > 0:
                    mean_e = np.mean(valid_errors)
                    std_e = np.std(valid_errors)
                    min_e = np.min(valid_errors)
                    max_e = np.max(valid_errors)
                    
                    print(f"Zakończono. Średni błąd: {mean_e:.2e}")
                    
                    writer.writerow([
                        f"F{f_idx}", 
                        alg_name, 
                        f"{mean_e:.4e}", 
                        f"{std_e:.4e}", 
                        f"{min_e:.4e}", 
                        f"{max_e:.4e}"
                    ])
                else:
                    print("Zakończono. Brak wyników (NotImplemented).")
                    writer.writerow([f"F{f_idx}", alg_name, "NaN", "NaN", "NaN", "NaN"])
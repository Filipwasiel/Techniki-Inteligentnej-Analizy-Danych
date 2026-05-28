import numpy as np
import csv
import os
from concurrent.futures import ProcessPoolExecutor
from cec2017.functions import all_functions
from src.plots import generate_combined_plots

def run_single_experiment(params):
    func_idx, dim, AlgorithmClass, run_id = params
    target_func = all_functions[func_idx - 1]
    f_optimum = func_idx * 100
    
    def obj_func(x):
        return target_func(x.reshape(1, -1))[0]

    model = AlgorithmClass(
        obj_func=obj_func, 
        dim=dim,
        func_idx=func_idx,
        run_id=run_id
    )
    
    try:
        gbest, gbest_fit, history = model.run()
        
        final_error = gbest_fit - f_optimum
        history_error = np.array(history) - f_optimum
        
        return final_error, history_error
        
    except (NotImplementedError, KeyError):
        return np.nan, np.nan

def run_experiments(algorithms, functions_to_test, dim, num_runs):
    """
    Główny zarządca procesów. Zleca zadania, zbiera wyniki końcowe i historie,
    a następnie ZAWDZE generuje wykresy porównawcze/pojedyncze w formacie PNG.
    """
    os.makedirs("results", exist_ok=True)
    csv_path = "results/final_report.csv"
    
    raw_storage = {alg.__name__: {} for alg in algorithms}
    history_storage = {alg.__name__: {} for alg in algorithms}
    
    # KROK 1: Wykonanie obliczeń
    for AlgorithmClass in algorithms:
        alg_name = AlgorithmClass.__name__
        print(f"\n>>> TESTOWANIE ALGORYTMU: {alg_name}")
        
        for f_idx in functions_to_test:
            print(f"  Obliczanie F{f_idx} ({num_runs} biegów)... ", end="", flush=True)
            
            tasks = [(f_idx, dim, AlgorithmClass, i) for i in range(num_runs)]
            workers = max(1, os.cpu_count() - 1)
            
            with ProcessPoolExecutor(max_workers=workers) as executor:
                results = list(executor.map(run_single_experiment, tasks))
            
            final_errors = []
            histories = []
            
            for final_e, hist_e in results:
                if not isinstance(final_e, float) and np.isnan(final_e):
                    continue
                final_errors.append(final_e)
                histories.append(hist_e)
            
            final_errors = np.array(final_errors)
            histories = np.array(histories)
            
            raw_storage[alg_name][f_idx] = final_errors
            history_storage[alg_name][f_idx] = histories
            
            if len(final_errors) > 0:
                print(f"Zakończono. Średni błąd: {np.mean(final_errors):.2e}")
            else:
                print("Zakończono. Brak wyników.")

    # KROK 2: Generowanie raportu końcowego i WYKRESÓW PNG (Zawsze dla każdej funkcji)
    print("\n>>> GENEROWANIE RAPORTU I WYKRESÓW ZBIEŻNOŚCI...")
    
    base_alg_name = algorithms[0].__name__
    modified_alg_name = algorithms[1].__name__ if len(algorithms) > 1 else None

    with open(csv_path, "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([
            "Function", "Algorithm", "Mean_Error", "Std_Dev", 
            "Best", "Worst", "Wilcoxon_p_value", "Verdict_w_t_l"
        ])

        for f_idx in functions_to_test:
            results_base = raw_storage[base_alg_name].get(f_idx, np.array([]))
            results_mod = raw_storage[modified_alg_name].get(f_idx, np.array([])) if modified_alg_name else np.array([])
            
            p_val = "-"
            verdict = "-"
            
            # Test Wilcoxona odpala się tylko przy dwóch algorytmach
            if modified_alg_name and len(results_base) == num_runs and len(results_mod) == num_runs:
                if np.array_equal(results_base, results_mod):
                    p_val = 1.0
                    verdict = "t"
                else:
                    from scipy.stats import wilcoxon
                    try:
                        _, p_val = wilcoxon(results_base, results_mod)
                        mean_base = np.mean(results_base)
                        mean_mod = np.mean(results_mod)
                        
                        if p_val < 0.05:
                            verdict = "w" if mean_mod < mean_base else "l"
                            p_val = f"{p_val:.4e}"
                        else:
                            verdict = "t"
                            p_val = f"{p_val:.4f}"
                    except ValueError:
                        p_val = "N/A"
                        verdict = "t"
            
            # Zapisz statystyki do pliku CSV
            if len(results_base) > 0:
                writer.writerow([
                    f"F{f_idx}", base_alg_name, f"{np.mean(results_base):.4e}", f"{np.std(results_base):.4e}",
                    f"{np.min(results_base):.4e}", f"{np.max(results_base):.4e}", "-", "-"
                ])
                
            if modified_alg_name and len(results_mod) > 0:
                writer.writerow([
                    f"F{f_idx}", modified_alg_name, f"{np.mean(results_mod):.4e}", f"{np.std(results_mod):.4e}",
                    f"{np.min(results_mod):.4e}", f"{np.max(results_mod):.4e}", p_val, verdict
                ])
            
            # --- ZMIANA: GENEROWANIE WYKRESÓW WYWOŁYWANE JEST ZAWSZE ---
            try:
                generate_combined_plots(
                    func_idx=f_idx,
                    raw_storage=raw_storage,
                    history_storage=history_storage
                )
                print(f"  -> Zapisano wykres: plots/F{f_idx}.png")
            except Exception as e:
                print(f"  -> Problem podczas zapisu wykresu dla F{f_idx}: {e}")
                
    print(f"\n>>> Sukces. Raport zapisano w: {csv_path}, a wykresy w folderze 'plots/'")
from src.experiment import run_experiments
from src.GGL_PSOD import GGL_PSOD_Raw, GGL_PSOD_Modified

if __name__ == "__main__":
    # Parametry przekazywane do funkcji, która robi eksperymenty
    algorithms = [GGL_PSOD_Raw, GGL_PSOD_Modified]
    functions_to_test = [5] # numery funkcji dla których bedzie testowane
    dim = 30
    num_runs = 8
    
    # Uruchomienie eksperymentów
    run_experiments(
        algorithms=algorithms,
        functions_to_test=functions_to_test,
        dim=dim,
        num_runs=num_runs
    )
from src.GGL_PSOD import GGL_PSOD
from src.func import *

if __name__ == "__main__":
    D = 30
    
    print(f"Uruchamiam GGL-PSOD dla funkcji Rastrigina (D={D})...")
    
    # Tworzenie instancji algorytmu
    model = GGL_PSOD(obj_func=rastrigin, dim=D)
    
    # Wykonanie optymalizacji
    best_pos, final_score = model.run()
    
    print("-" * 30)
    print(f"Pozycja (pierwsze 3): {best_pos[:3]} \nNajlepsza wartość: {final_score:.6e}")
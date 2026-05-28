import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os

def generate_combined_plots(func_idx, raw_storage, history_storage):
    """
    Automatycznie generuje wykres zbieżności oraz boxplot dla WSZYSTKICH 
    przekazanych algorytmów (obsługuje zarówno 1, jak i więcej algorytmów).
    
    Zapisuje wyniki jako pliki PNG w folderze 'plots/' z nazwą 'F{numer}.png'.
    """
    os.makedirs("plots", exist_ok=True)
    
    available_algs = [alg for alg in raw_storage.keys() if func_idx in raw_storage[alg]]
    
    if len(available_algs) == 0:
        return

    colors = {'GGL_PSOD_Raw': '#d9534f', 'GGL_PSOD_Modified': '#0275d8'}
    box_palette = ['#ff9999', '#9999ff']

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 5))

    # -----------------------------------------------------------------
    # LEWY WYKRES: ZBIEŻNOŚĆ (CONVERGENCE CURVE)
    # -----------------------------------------------------------------
    for alg_name in available_algs:
        histories = history_storage[alg_name][func_idx]
        if len(histories) > 0:
            mean_curve = np.mean(histories, axis=0)
            iterations = np.arange(len(mean_curve))
            
            if len(mean_curve) <= 600:
                iterations = iterations * (6000 / (len(mean_curve) - 1))
                
            color = colors.get(alg_name, None)
            ls = '--' if 'Raw' in alg_name else '-'
            ax1.plot(iterations, mean_curve, label=alg_name, color=color, linestyle=ls, linewidth=2)

    ax1.set_yscale('log')
    ax1.set_xlim(0, 6000)
    ax1.set_title(f'Zbieżność błędu średniego dla F{func_idx}', fontsize=12, fontweight='bold')
    ax1.set_xlabel('Iteracje', fontsize=10)
    ax1.set_ylabel('Średni błąd (skala log)', fontsize=10)
    ax1.grid(True, which="both", ls="--", alpha=0.4)
    ax1.legend(fontsize=10)

    # -----------------------------------------------------------------
    # PRAWY WYKRES: ROZRZUT WYNIKÓW (BOXPLOT)
    # -----------------------------------------------------------------
    box_data = [raw_storage[alg_name][func_idx] for alg_name in available_algs]
    
    sns.boxplot(data=box_data, ax=ax2, palette=box_palette[:len(available_algs)], width=0.4)
    ax2.set_xticks(ticks=range(len(available_algs)))
    ax2.set_xticklabels(available_algs, fontsize=10)
    
    # Nakładanie surowych punktów z 51 biegów (jitter)
    for i, alg_name in enumerate(available_algs):
        err_array = raw_storage[alg_name][func_idx]
        x_jitter = np.random.normal(i, 0.04, size=len(err_array))
        ax2.scatter(x_jitter, err_array, alpha=0.5, color='black', edgecolor='none', s=20)

    ax2.set_title(f'Rozrzut wyników końcowych dla F{func_idx}', fontsize=12, fontweight='bold')
    ax2.set_ylabel('Błąd końcowy optymalizacji', fontsize=10)
    ax2.grid(True, axis='y', ls="--", alpha=0.4)

    plt.tight_layout()
    
    plt.savefig(f"plots/F{func_idx}.png", dpi=300)
    plt.close()
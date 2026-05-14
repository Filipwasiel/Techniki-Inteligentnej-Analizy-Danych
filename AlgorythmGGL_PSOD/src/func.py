import numpy as np

def rastrigin(x):
    """Funkcja Rastrigina (F5 w CEC2017) [cite: 242]"""
    return 10 * len(x) + np.sum(x**2 - 10 * np.cos(2 * np.pi * x))
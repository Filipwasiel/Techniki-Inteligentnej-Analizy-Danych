import pandas as pd
import numpy as np

# Tworzenie "wymagających" danych
data = {
    "ID": range(1, 51),
    "Data operacji": pd.date_range(start="2023-01-01", periods=50).strftime("%Y-%m-%d"),
    "Długi tekst (test zawijania)": [
        "To jest bardzo długi tekst, który ma na celu sprawdzenie, czy Twoja aplikacja "
        "poprawnie zawija wiersze w tabeli Worda, nie ucinając niczego po prawej stronie. " * 2
        if i % 5 == 0 else "Krótki tekst."
        for i in range(50)
    ],
    "Liczby i ułamki": np.random.uniform(10.5, 999.9, 50).round(2),
    "Puste komórki (test NaN)": [None if i % 6 == 0 else f"Wartość {i}" for i in range(50)],
    "Nowa linia (Enter)": [
        f"Linia 1\nLinia 2\nZnaki: !@#$%^&*" if i % 4 == 0 else "Zwykły wiersz"
        for i in range(50)
    ],
    "Zakazane znaki (test XML)": [
        "Tekst z ukrytym znakiem: " + chr(11) + " <- tutaj" if i % 7 == 0 else "Czysty tekst"
        for i in range(50)
    ]
}

# Tworzenie DataFrame i zapis do Excela
df = pd.DataFrame(data)
nazwa_pliku = "wymagajacy_test.xlsx"
df.to_excel(nazwa_pliku, index=False)

print(f"Gotowe! Plik '{nazwa_pliku}' został wygenerowany w Twoim folderze.")
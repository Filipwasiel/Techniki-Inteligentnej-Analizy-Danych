import os
import shutil
import random
import kagglehub
from PIL import Image


### ZAPIS DANYCH
os.environ["KAGGLEHUB_CACHE"] = r"E:\KaggleCache"

# ==========================================
# KONFIGURACJA EKSPERYMENTU
# ==========================================

RANDOM_SEED = 420

SELECTED_CLASSES = [
    'apple_pie', 'hamburger', 'chicken_curry', 'donuts', 'french_fries', 
    'ice_cream', 'pizza', 'tacos', 'omelette','beef_carpaccio'
]

# SELECTED_CLASSES = [
#     'apple_pie', 'ravioli', 'ramen'
# ]
LIMIT_PER_CLASS = 1000  

# Miejsce zapisywania plików pobranych z Kaggle (można dostosować do własnych potrzeb)
os.environ["KAGGLEHUB_CACHE"] = r"E:\KaggleCache"
# ==========================================

def is_image_corrupt(file_path):
    """Weryfikacja czy plik jest poprawnym obrazem RGB."""
    try:
        with Image.open(file_path) as img:
            img.verify() 
        with Image.open(file_path) as img:
            img.load() 
            if img.mode != 'RGB':
                return True
        return False
    except Exception:
        return True

def initialize_raw_data():
    """Pobiera Food-101 i przygotowuje bazę raw dla wybranych klas."""
    # Ścieżka do folderu 'data/raw' (zakładając strukturę projektu)
    raw_path = os.path.join("data", "raw")
    
    # Sprawdzenie czy dane już są (żeby nie marnować czasu przy każdym uruchomieniu)
    if os.path.exists(raw_path):
        existing_classes = os.listdir(raw_path)
        if all(cls in existing_classes for cls in SELECTED_CLASSES):
            print("Wybrane klasy Food-101 już istnieją w bazie raw.")
            return raw_path

    print("Pobieranie Food-101 z Kaggle...")
    # Dataset kmader/food41 zawiera obrazy w strukturze images/<klasa>/<plik>.jpg
    downloaded_path = kagglehub.dataset_download("kmader/food41")
    src_images_dir = os.path.join(downloaded_path, "images")
    
    if os.path.exists(raw_path):
        shutil.rmtree(raw_path)
    os.makedirs(raw_path, exist_ok=True)

    for cls in SELECTED_CLASSES:
        src_cls_folder = os.path.join(src_images_dir, cls)
        if not os.path.exists(src_cls_folder):
            print(f"Ostrzeżenie: Klasa {cls} nie istnieje w źródle!")
            continue

        dest_cls_folder = os.path.join(raw_path, cls)
        os.makedirs(dest_cls_folder, exist_ok=True)
        
        all_images = [f for f in os.listdir(src_cls_folder) if f.endswith('.jpg')]
        random.shuffle(all_images)
        
        valid_count = 0
        print(f"Przetwarzanie klasy: {cls}...", end="\r")
        
        for f in all_images:
            if valid_count >= LIMIT_PER_CLASS:
                break
                
            src_file = os.path.join(src_cls_folder, f)
            if not is_image_corrupt(src_file):
                shutil.copy(src_file, os.path.join(dest_cls_folder, f))
                valid_count += 1
        
        print(f"Klasa {cls}: skopiowano {valid_count} poprawnych zdjęć.")

    return raw_path

def split_data(train_split=0.5):
    """Dzieli dane z 'raw' na 'train' i 'test' zachowując balans klas oraz powtarzalność."""
    raw_path = os.path.join("data", "raw")
    train_path = os.path.join("data", "train")
    test_path = os.path.join("data", "test")

    for p in [train_path, test_path]:
        if os.path.exists(p):
            shutil.rmtree(p)

    classes = [d for d in os.listdir(raw_path) if os.path.isdir(os.path.join(raw_path, d))]

    for cls in classes:
        os.makedirs(os.path.join(train_path, cls), exist_ok=True)
        os.makedirs(os.path.join(test_path, cls), exist_ok=True)
        
        images = os.listdir(os.path.join(raw_path, cls))
        
        # Ustawienie seeda bezpośrednio przed mieszaniem listy dla każdej klasy
        random.seed(RANDOM_SEED)
        random.shuffle(images)
        
        split_idx = int(len(images) * train_split)
        train_files = images[:split_idx]
        test_files = images[split_idx:]

        for f in train_files:
            shutil.copy(os.path.join(raw_path, cls, f), os.path.join(train_path, cls, f))
        for f in test_files:
            shutil.copy(os.path.join(raw_path, cls, f), os.path.join(test_path, cls, f))

    print(f"Podział zakończony (Seed: {RANDOM_SEED}): {train_split*100:.0f}% trening / {(1-train_split)*100:.0f}% test.")

if __name__ == "__main__":
    initialize_raw_data()
    split_data(train_split=0.8)
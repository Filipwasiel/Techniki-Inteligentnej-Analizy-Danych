import os
import shutil
import random
import kagglehub
from PIL import Image

def is_image_corrupt(file_path):
    try:
        # Sprawdzenie nagłówka (JPEG zaczyna się od \xff\xd8)
        with open(file_path, 'rb') as f:
            header = f.read(10)
            if b'JFIF' not in header and b'Exif' not in header:
                return True
        
        # Próba otwarcia i faktycznego przetworzenia pikseli
        with Image.open(file_path) as img:
            img.verify() 
        with Image.open(file_path) as img:
            img.load() 
            # Sprawdzenie czy to na pewno RGB (niektóre pliki są w skali szarości, co też wywala TF)
            if img.mode != 'RGB':
                return True
        return False
    except Exception:
        return True

def initialize_raw_data(limit_per_class=1000):
    raw_path = os.path.join("..", "data", "raw")
    
    if os.path.exists(raw_path) and len(os.listdir(os.path.join(raw_path, "Cat"))) >= limit_per_class:
        print("Baza zdjęć (raw) już istnieje i jest zweryfikowana.")
        return raw_path

    print("Przygotowywanie czystej bazy danych z Kaggle...")
    downloaded_path = kagglehub.dataset_download("shaunthesheep/microsoft-catsvsdogs-dataset")
    src_pet_images = os.path.join(downloaded_path, "PetImages")
    
    if os.path.exists(raw_path):
        shutil.rmtree(raw_path)

    for cls in ["Cat", "Dog"]:
        os.makedirs(os.path.join(raw_path, cls), exist_ok=True)
        src_folder = os.path.join(src_pet_images, cls)
        
        # 1. Pobieramy listę plików spełniających podstawowe kryteria
        all_images = [f for f in os.listdir(src_folder) if f.endswith('.jpg')]
        
        valid_images = []
        print(f"Weryfikacja zdjęć dla klasy {cls}...")
        
        # 2. NAPRAWA: Sprawdzamy każdy plik zanim trafi do bazy raw
        for f in all_images:
            full_path = os.path.join(src_folder, f)
            if os.path.getsize(full_path) > 0 and not is_image_corrupt(full_path):
                valid_images.append(f)
            
            if len(valid_images) >= limit_per_class:
                break
        
        # 3. Kopiowanie zweryfikowanych plików
        for f in valid_images:
            shutil.copy(os.path.join(src_folder, f), os.path.join(raw_path, cls, f))
    
    print(f"Pobrano i zweryfikowano {len(valid_images) * 2} poprawnych zdjęć.")
    return raw_path

def split_data(train_split=0.5):
    # Ścieżki z Twojego configu
    raw_path = os.path.join("..", "data", "raw")
    train_path = os.path.join("..", "data", "train")
    test_path = os.path.join("..", "data", "test")

    # Czyścimy foldery operacyjne, by nie mieszać starych podziałów
    for p in [train_path, test_path]:
        if os.path.exists(p):
            shutil.rmtree(p)

    for cls in ["Cat", "Dog"]:
        os.makedirs(os.path.join(train_path, cls), exist_ok=True)
        os.makedirs(os.path.join(test_path, cls), exist_ok=True)
        
        images = os.listdir(os.path.join(raw_path, cls))
        random.shuffle(images) # Losowość podziału
        
        split_idx = int(len(images) * train_split)
        train_files = images[:split_idx]
        test_files = images[split_idx:]

        # Fizyczne kopiowanie plików do odpowiednich podfolderów
        for f in train_files:
            shutil.copy(os.path.join(raw_path, cls, f), os.path.join(train_path, cls, f))
        for f in test_files:
            shutil.copy(os.path.join(raw_path, cls, f), os.path.join(test_path, cls, f))

    print(f"Podział zakończony: {train_split*100}% do treningu.")

if __name__ == "__main__":
    initialize_raw_data()
    split_data(train_split=0.5)
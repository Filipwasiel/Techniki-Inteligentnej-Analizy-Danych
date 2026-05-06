from tensorflow.keras import layers, models
from src import config

def build_cnn_model(input_shape=None):
    if input_shape is None:
        input_shape = (config.IMG_SIZE[0], config.IMG_SIZE[1], 3)

    model = models.Sequential([
        layers.Input(shape=input_shape),
        layers.Rescaling(1./255),
        
        # Blok 1: Wyłapywanie krawędzi
        layers.Conv2D(32, (3, 3), activation='relu', padding='same'),
        layers.MaxPooling2D((2, 2)),
        layers.Dropout(0.2), # Delikatny dropout już na starcie
        
        # Blok 2: Wyłapywanie kształtów
        layers.Conv2D(64, (3, 3), activation='relu', padding='same'),
        layers.MaxPooling2D((2, 2)),
        layers.Dropout(0.2),
        
        # Blok 3: Wyłapywanie detali (uszy, oczy)
        layers.Conv2D(128, (3, 3), activation='relu', padding='same'),
        layers.MaxPooling2D((2, 2)),
        layers.Dropout(0.3),
        
        layers.Flatten(), # Powrót do Flatten dla lepszej precyzji
        
        layers.Dense(128, activation='relu'),
        layers.Dropout(0.5), # Mocny hamulec przed warstwą wyjściową
        
        layers.Dense(1, activation='sigmoid')
    ])

    model.compile(
        optimizer='adam', 
        loss='binary_crossentropy', 
        metrics=['accuracy']
    )
    
    return model
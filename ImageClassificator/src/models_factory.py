# models_factory.py
"""
Factory for building different model architectures.
Add new models here as they are developed.
"""
from abc import ABC, abstractmethod
from tensorflow.keras import layers, models
from src import config
import tensorflow as tf
from tensorflow.keras.applications.resnet50 import preprocess_input

class BaseModel(ABC):
    """Abstract base class for all model architectures"""
    
    def __init__(self, input_shape=None):
        if input_shape is None:
            input_shape = (config.IMG_SIZE[0], config.IMG_SIZE[1], 3)
        self.input_shape = input_shape
        self.model = None
    
    @abstractmethod
    def build(self):
        """Build and compile the model"""
        pass
    
    def train(self, train_ds, epochs, validation_data, verbose=2):
        """Train the model"""
        if self.model is None:
            raise ValueError("Model not built. Call build() first.")
        
        return self.model.fit(
            train_ds,
            epochs=epochs,
            validation_data=validation_data,
            verbose=verbose
        )
    
    def predict(self, dataset):
        """Get predictions on dataset"""
        return self.model.predict(dataset)


class SimpleCNN(BaseModel):
    """
    Simple CNN with 3 convolutional blocks.
    
    Architecture:
    - Block 1: Conv32 + MaxPool + Dropout(0.2)
    - Block 2: Conv64 + MaxPool + Dropout(0.2)
    - Block 3: Conv128 + MaxPool + Dropout(0.3)
    - Dense layers: 128 + Dropout(0.5) + Output(Sigmoid)
    """
    
    def build(self):
        self.model = models.Sequential([
            layers.Input(shape=self.input_shape),
            layers.Rescaling(1./255),
            
            # Block 1: Edge detection
            layers.Conv2D(32, (3, 3), activation='relu', padding='same'),
            layers.MaxPooling2D((2, 2)),
            layers.Dropout(0.2),
            
            # Block 2: Shape detection
            layers.Conv2D(64, (3, 3), activation='relu', padding='same'),
            layers.MaxPooling2D((2, 2)),
            layers.Dropout(0.2),
            
            # Block 3: Detail detection
            layers.Conv2D(128, (3, 3), activation='relu', padding='same'),
            layers.MaxPooling2D((2, 2)),
            layers.Dropout(0.3),
            
            layers.Flatten(),
            
            layers.Dense(128, activation='relu'),
            layers.Dropout(0.5),
            
            layers.Dense(1, activation='sigmoid')
        ])
        
        self.model.compile(
            optimizer='adam',
            loss='binary_crossentropy',
            metrics=['accuracy']
        )


class MobileNetModel(BaseModel):
    """
    MobileNetV2 architecture.
    Lightweight and efficient, perfect for quick testing on RTX 2050.
    """
    
    def build(self):
        # 1. Pobieramy bazę modelu bez "głowy" klasyfikacyjnej
        # weights=None oznacza, że trenujemy od zera (zgodnie z Twoim zadaniem)
        base_model = tf.keras.applications.MobileNetV2(
            input_shape=self.input_shape,
            include_top=False,
            weights='imagenet' # Użyj wag wypracowanych na milionach zdjęć
        )
        base_model.trainable = False # Zamroź bazę, trenuj tylko swoją "głowę"
        
        self.model = models.Sequential([
            layers.Input(shape=self.input_shape),
            
            # MobileNet ma wbudowane specyficzne skalowanie, 
            # ale ponieważ trenujemy od zera, wystarczy standardowa normalizacja
            layers.Rescaling(1./255),
            
            base_model,
            
            # Global Average Pooling zamiast Flatten() - drastycznie przyspiesza 
            # działanie i redukuje liczbę parametrów (z kilku milionów do tysięcy)
            layers.GlobalAveragePooling2D(),
            
            layers.Dropout(0.3),
            layers.Dense(1, activation='sigmoid')
        ])
        
        self.model.compile(
            optimizer='adam',
            loss='binary_crossentropy',
            metrics=['accuracy']
        )



class ResNetModel(BaseModel):
    """
    ResNet50 architecture with residual connections.
    Excellent for deep feature extraction but heavier than MobileNet.
    """
    
    def build(self):
        # 1. Pobieramy bazę modelu
        # include_top=False usuwa oryginalne warstwy klasyfikacji ImageNet
        base_model = tf.keras.applications.ResNet50(
            input_shape=self.input_shape,
            include_top=False,
            weights='imagenet' # Zalecane 'imagenet' dla uniknięcia błędów z poprzedniego testu
        )
        
        # Zamrażamy bazę, aby nie popsuć wag podczas pierwszych epok
        base_model.trainable = False 
        
        # 2. Budowa struktury
        self.model = models.Sequential([
            layers.Input(shape=self.input_shape),
            
            # ResNet50 używa preprocessingu, który konwertuje RGB na BGR 
            # i odejmuje średnią z ImageNet. Lambda załatwi to za nas.
            layers.Lambda(preprocess_input),
            
            base_model,
            
            # Global Average Pooling redukuje wymiary z (4,4,2048) do (2048)
            layers.GlobalAveragePooling2D(),
            
            layers.Dropout(0.4),
            layers.Dense(128, activation='relu'),
            layers.Dense(1, activation='sigmoid')
        ])
        
        # 3. Kompilacja z nieco niższym learning rate dla stabilności
        self.model.compile(
            optimizer=tf.keras.optimizers.Adam(learning_rate=0.0001),
            loss='binary_crossentropy',
            metrics=['accuracy']
        )


# Model registry - add new models here
MODEL_REGISTRY = {
    'simple_cnn': SimpleCNN,
    'mobilenet': MobileNetModel,
    'resnet': ResNetModel,
}


def create_model(model_name: str, input_shape=None) -> BaseModel:
    """
    Factory function to create model instances
    
    Args:
        model_name: Name of the model (must be in MODEL_REGISTRY)
        input_shape: Input shape tuple (default from config)
    
    Returns:
        Instantiated model object (not yet built)
    
    Raises:
        ValueError: If model_name not found in registry
    """
    if model_name not in MODEL_REGISTRY:
        available = list(MODEL_REGISTRY.keys())
        raise ValueError(
            f"Model '{model_name}' not found in registry.\n"
            f"Available models: {available}"
        )
    
    model_class = MODEL_REGISTRY[model_name]
    return model_class(input_shape)


def list_available_models():
    """Return list of available models"""
    return list(MODEL_REGISTRY.keys())

# models_factory.py
from abc import ABC, abstractmethod
from tensorflow.keras import layers, models
from src import config
import tensorflow as tf
from tensorflow.keras.applications.resnet50 import preprocess_input

class BaseModel(ABC):
    # DODANO: num_classes w __init__
    def __init__(self, input_shape=None, num_classes=2):
        if input_shape is None:
            input_shape = (config.IMG_SIZE[0], config.IMG_SIZE[1], 3)
        self.input_shape = input_shape
        self.num_classes = num_classes
        self.model = None
    
    @abstractmethod
    def build(self):
        pass
    
    def train(self, train_ds, epochs, validation_data, verbose=2):
        if self.model is None:
            raise ValueError("Model not built. Call build() first.")
        return self.model.fit(train_ds, epochs=epochs, validation_data=validation_data, verbose=verbose)
    
    def predict(self, dataset):
        return self.model.predict(dataset)

class SimpleCNN(BaseModel):
    def build(self):
        self.model = models.Sequential([
            layers.Input(shape=self.input_shape),
            layers.Rescaling(1./255),
            layers.Conv2D(32, (3, 3), activation='relu', padding='same'),
            layers.MaxPooling2D((2, 2)),
            layers.Conv2D(64, (3, 3), activation='relu', padding='same'),
            layers.MaxPooling2D((2, 2)),
            layers.Flatten(),
            layers.Dense(128, activation='relu'),
            layers.Dropout(0.5),
            # ZMIANA: Softmax i dynamiczna liczba klas
            layers.Dense(self.num_classes, activation='softmax')
        ])
        self.model.compile(
            optimizer='adam',
            loss='sparse_categorical_crossentropy', # ZMIANA na wieloklasowy
            metrics=['accuracy']
        )

class MobileNetModel(BaseModel):
    def build(self):
        base_model = tf.keras.applications.MobileNetV2(
            input_shape=self.input_shape, include_top=False, weights='imagenet'
        )
        base_model.trainable = False
        self.model = models.Sequential([
            layers.Input(shape=self.input_shape),
            layers.Rescaling(1./255),
            base_model,
            layers.GlobalAveragePooling2D(),
            layers.Dropout(0.3),
            layers.Dense(self.num_classes, activation='softmax') # ZMIANA
        ])
        self.model.compile(
            optimizer='adam',
            loss='sparse_categorical_crossentropy', # ZMIANA
            metrics=['accuracy']
        )

class ResNetModel(BaseModel):
    def build(self):
        base_model = tf.keras.applications.ResNet50(
            input_shape=self.input_shape, include_top=False, weights='imagenet'
        )
        base_model.trainable = False 
        self.model = models.Sequential([
            layers.Input(shape=self.input_shape),
            layers.Lambda(preprocess_input),
            base_model,
            layers.GlobalAveragePooling2D(),
            layers.Dropout(0.4),
            layers.Dense(self.num_classes, activation='softmax') # ZMIANA
        ])
        self.model.compile(
            optimizer=tf.keras.optimizers.Adam(learning_rate=0.0001),
            loss='sparse_categorical_crossentropy', # ZMIANA
            metrics=['accuracy']
        )

MODEL_REGISTRY = {
    'simple_cnn': SimpleCNN,
    'mobilenet': MobileNetModel,
    'resnet': ResNetModel,
}

# DODANO: num_classes w argumentach funkcji
def create_model(model_name: str, input_shape=None, num_classes=2) -> BaseModel:
    if model_name not in MODEL_REGISTRY:
        raise ValueError(f"Model '{model_name}' not found.")
    
    model_class = MODEL_REGISTRY[model_name]
    return model_class(input_shape, num_classes=num_classes)
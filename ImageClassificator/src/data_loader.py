import tensorflow as tf
from tensorflow.keras.utils import image_dataset_from_directory 
from src import config

def load_data():
    train_ds = image_dataset_from_directory(
        config.TRAIN_DIR,
        image_size=config.IMG_SIZE,
        batch_size=config.BATCH_SIZE,
        label_mode="int"
    )
    
    # Wyciągamy nazwy klas ZANIM zrobimy prefetch
    class_names = train_ds.class_names
    
    test_ds = image_dataset_from_directory(
        config.TEST_DIR,
        image_size=config.IMG_SIZE,
        batch_size=config.BATCH_SIZE,
        label_mode="int",
        shuffle=False
    )

    AUTOTUNE = tf.data.AUTOTUNE
    train_ds = train_ds.prefetch(buffer_size=AUTOTUNE)
    test_ds = test_ds.prefetch(buffer_size=AUTOTUNE)

    # ZWRACAMY TRZY ELEMENTY
    return train_ds, test_ds, class_names
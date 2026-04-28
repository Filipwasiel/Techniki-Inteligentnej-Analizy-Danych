import tensorflow as tf
from tensorflow.keras.preprocessing import image_dataset_from_directory
from src import config


def load_data():
   train_ds = image_dataset_from_directory(
      config.TRAIN_DIR,
      image_size=config.IMG_SIZE,
      batch_size=config.BATCH_SIZE,
      label_mode="binary"
   )
   test_ds = image_dataset_from_directory(
      config.TEST_DIR,
      image_size=config.IMG_SIZE,
      batch_size=config.BATCH_SIZE,
      label_mode="binary",
      shuffle=False
   )

   AUTOTUNE = tf.data.AUTOTUNE
   train_ds = train_ds.cache().prefetch(AUTOTUNE)
   test_ds = test_ds.cache().prefetch(AUTOTUNE)

   return train_ds, test_ds
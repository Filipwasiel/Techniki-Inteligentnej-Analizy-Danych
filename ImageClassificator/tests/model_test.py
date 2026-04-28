import unittest
from src.model import build_cnn_model
import tensorflow as tf

class TestModel(unittest.TestCase):
   def setUp(self):
      self.model = build_cnn_model()

   def test_model_arch(self):
      self.assertIsInstance(self.model, tf.keras.models.Sequential)

   def test_model_output_shape(self):
      last_layer = self.model.layers[-1]
      self.assertEqual(last_layer.units, 1)
      self.assertEqual(last_layer.activation.__name__, 'sigmoid')

   def test_model_compile(self):
      self.assertIsNotNone(self.model.optimizer)
      self.assertIsNotNone(self.model.loss)

if __name__ == '__main__':
   unittest.main()

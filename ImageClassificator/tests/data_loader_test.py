import unittest
from unittest.mock import patch, MagicMock
import tensorflow as tf
from src.data_loader import load_data


class TestDataLoader(unittest.TestCase):
   @patch('src.data_loader.image_dataset_from_directory')
   def test_load_data(self, mock_image_dataset):
      mock_ds = MagicMock()
      mock_image_dataset.return_value = mock_ds
      train_ds, test_ds = load_data()

      self.assertEqual(mock_image_dataset.call_count, 2)
      self.assertIsNotNone(train_ds)
      self.assertIsNotNone(test_ds)



if __name__ == '__main__':
   unittest.main()

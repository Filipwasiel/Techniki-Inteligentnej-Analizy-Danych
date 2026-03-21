import unittest
from src.Logic.nlp import filter_ingredients


class TestNLP(unittest.TestCase):
    def test_basic(self):
        text = "Mam jajko, 3 jajka i 4 jajka, masło, szynkę i chleb. Chciałbym przygotować prosty posiłek."
        result = filter_ingredients(text)
        # self.assertEqual(len(result), 6)
        self.assertIn("jajko", result)
        self.assertIn("chleb", result)
        self.assertIn("szynka", result)
        self.assertIn("jajka", result)

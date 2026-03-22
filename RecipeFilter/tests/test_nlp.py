import unittest
from src.Logic.nlp import filter_ingredients


class TestNLP(unittest.TestCase):
    def test_filtering_ingredients(self):
        text = "Mam jajko, 3 jajka i 4 jajka, masło, szynkę i chleb. Chciałbym przygotować prosty posiłek."
        result = filter_ingredients(text)
        self.assertIn("jajko", result)
        self.assertIn("chleb", result)
        self.assertIn("szynka", result)
        self.assertIn("jajka", result)
    def test_no_nouns(self):
        text = "szybko, łatwo, smacznie"
        result = filter_ingredients(text)
        self.assertEqual(len(result), 0)
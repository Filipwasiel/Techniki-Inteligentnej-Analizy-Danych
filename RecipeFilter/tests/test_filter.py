import unittest
from src.Logic.filter import filter_recipes
RECIPE_BASE = [
    {
        "id": 1,
        "nazwa": "Klasyczne Naleśniki",
        "skladniki": {"jajko", "mleko", "mąka", "olej", "sól"},
        "instrukcja": "Zmiksuj jajka, mleko i mąkę. Smaż na rozgrzanym oleju z obu stron."
    },
    {
        "id": 2,
        "nazwa": "Jajecznica na maśle",
        "skladniki": {"jajko", "masło", "sól"},
        "instrukcja": "Roztop masło na patelni. Wbij jajka, posól i smaż ciągle mieszając."
    },
    {
        "id": 3,
        "nazwa": "Omlet z szynką",
        "skladniki": {"jajko", "szynka", "masło", "sól", "pieprz"},
        "instrukcja": "Roztrzep jajka. Szynkę podsmaż na maśle, wylej jajka, smaż pod przykryciem."
    },
    {
        "id": 4,
        "nazwa": "Tosty z serem",
        "skladniki": {"chleb", "ser", "masło"},
        "instrukcja": "Chleb posmaruj masłem, połóż ser, zapiekaj w tosterze."
    }
]

class TestFilter(unittest.TestCase):
    def test_find_specified_recipe(self):
        required = ["masło", "chleb"]
        result = filter_recipes(required, RECIPE_BASE)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["nazwa"], "Tosty z serem")

    def test_find_many_recipes(self):
        required = ["jajko", "sól"]
        result = filter_recipes(required, RECIPE_BASE)
        self.assertEqual(len(result), 3)
        recipes_name = [r["nazwa"] for r in result]
        self.assertNotIn("Tosty z serem", recipes_name)
        self.assertIn("Omlet z szynką", recipes_name)

    def test_ignores_duplicated_ingredients(self):
        required = ["masło", "chleb", "masło"]
        result = filter_recipes(required, RECIPE_BASE)
        self.assertEqual(len(result), 1)

    def test_unknown_ingredient(self):
        required = ["masło", "unknown"]
        result = filter_recipes(required, RECIPE_BASE)
        self.assertEqual(len(result), 0)

if __name__ == "__main__":
    unittest.main()
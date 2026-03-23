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

def filter_recipes(ingredients ,recipes):
    ingredients_without_duplicates = set(ingredients)
    matching_recipes = []
    for recipe in recipes:
        if ingredients_without_duplicates.issubset(recipe["skladniki"]):
            matching_recipes.append(recipe)
    return matching_recipes
import functools
from deep_translator import GoogleTranslator

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
    },
    {
        "id": 5,
        "nazwa": "Kanapka z miodem",
        "skladniki": {"chleb", "masło", "miód"},
        "instrukcja": "Chleb posmaruj masłem i grubą warstwą miodu."
    },
    {
        "id": 6,
        "nazwa": "Szybka Pasta Jajeczna",
        "skladniki": {"jajko", "masło", "sól", "pieprz"},
        "instrukcja": "Ugotuj jajka na twardo. Rozgnieć widelcem z masłem, solą i pieprzem."
    },
    {
        "id": 7,
        "nazwa": "Placki z jabłkami",
        "skladniki": {"jajko", "mleko", "mąka", "jabłko", "cukier"},
        "instrukcja": "Zrób gęste ciasto z jajka, mleka i mąki. Dodaj starte jabłka i smaż na złoto."
    },
    {
        "id": 8,
        "nazwa": "Sałatka z szynką i serem",
        "skladniki": {"szynka", "ser", "jajko", "sól", "pieprz"},
        "instrukcja": "Pokrój szynkę, ser i ugotowane jajka w kostkę. Wymieszaj z przyprawami."
    },
    {
        "id": 9,
        "nazwa": "Pieczone jabłko z miodem",
        "skladniki": {"jabłko", "miód", "masło"},
        "instrukcja": "Wydrąż środek jabłka, do środka daj masło i miód. Zapiekaj w 180°C przez 20 minut."
    },
    {
        "id": 10,
        "nazwa": "Grzanki z masłem i solą",
        "skladniki": {"chleb", "masło", "sól"},
        "instrukcja": "Opiecz chleb w tosterze lub na patelni, posmaruj masłem i posyp solą."
    }
]

@functools.lru_cache(maxsize=None)
def get_recipe_base(lang: str = 'pl'):
    """
    Returns the recipe base in the requested language.
    By default returns the Polish RECIPE_BASE. For 'en' builds a translated copy.
    The result is cached to avoid repeated network calls.
    """
    if lang == 'pl':
        return RECIPE_BASE
    if lang == 'en':
        translator = GoogleTranslator(source='auto', target='en')
        translated = []
        for r in RECIPE_BASE:
            try:
                name = translator.translate(r['nazwa'])
            except Exception:
                name = r['nazwa']
            translated_ingredients = set()
            for ing in r['skladniki']:
                try:
                    t = translator.translate(ing)
                except Exception:
                    t = ing
                translated_ingredients.add(t.lower())
            try:
                instr = translator.translate(r['instrukcja'])
            except Exception:
                instr = r['instrukcja']
            translated.append({
                "id": r['id'],
                "nazwa": name,
                "skladniki": translated_ingredients,
                "instrukcja": instr
            })
        return translated
    raise ValueError(f"Unsupported language: {lang}")


def filter_recipes(ingredients, recipes):
    """
    Given a list of detected ingredients and a recipe list, return:
      - exact_matches: recipes where ALL recipe ingredients are present
      - partial_matches: recipes where SOME ingredients are present, plus missing ingredients
    Both lists are returned as tuples: (exact_matches, partial_matches)
    partial_matches contains dicts with keys: recipe, found, missing
    """
    ingredients_set = set([i.lower() for i in ingredients])
    exact_matches = []
    partial_matches = []

    for recipe in recipes:
        recipe_ings = set([i.lower() for i in recipe["skladniki"]])
        found = ingredients_set.intersection(recipe_ings)
        missing = recipe_ings - found
        if not found:
            # no ingredients matched; skip
            continue
        if not missing:
            exact_matches.append(recipe)
        else:
            partial_matches.append({
                "recipe": recipe,
                "found": sorted(list(found)),
                "missing": sorted(list(missing)),
                "missing_count": len(missing)
            })

    # sort partial matches by fewest missing ingredients first
    partial_matches.sort(key=lambda x: x["missing_count"]) 
    return exact_matches, partial_matches

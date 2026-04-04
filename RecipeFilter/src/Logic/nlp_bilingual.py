import spacy
from src.Logic.filter_bilingual import RECIPE_BASE, get_recipe_base

# cache for loaded spacy models
_nlp_cache = {}


def _load_spacy_model(lang: str = 'pl'):
    model_map = {
        'pl': 'pl_core_news_sm',
        'en': 'en_core_web_sm'
    }
    model_name = model_map.get(lang)
    if model_name is None:
        raise ValueError("Unsupported language for NLP model")
    if lang in _nlp_cache:
        return _nlp_cache[lang]
    try:
        nlp = spacy.load(model_name)
    except OSError:
        import os
        os.system(f"python -m spacy download {model_name}")
        nlp = spacy.load(model_name)
    _nlp_cache[lang] = nlp
    return nlp


def get_allowed_ingredients(recipes=None):
    """Creates a set of unique ingredient names from provided recipes (lowercased)."""
    allowed = set()
    if recipes is None:
        recipes = RECIPE_BASE
    for przepis in recipes:
        for skladnik in przepis["skladniki"]:
            allowed.add(skladnik.lower())
    return allowed


def filter_ingredients(text: str, recipes=None, lang: str = 'pl') -> list:
    """
    Analyze the text and return ingredients that exist in the provided recipe base.
    lang selects the spacy model and expected recipe language ('pl' or 'en').
    """
    if not text:
        return []
    nlp = _load_spacy_model(lang)
    doc = nlp(text.lower())
    if recipes is None:
        recipes = RECIPE_BASE if lang == 'pl' else get_recipe_base('en')
    allowed_list = get_allowed_ingredients(recipes)
    caught_ingredients = []
    for token in doc:
        lemma = token.lemma_.lower()
        if token.pos_ in ["NOUN", "ADJ"] and lemma in allowed_list:
            caught_ingredients.append(lemma)
    return list(set(caught_ingredients))

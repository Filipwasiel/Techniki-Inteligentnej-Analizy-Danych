import spacy
from spacy.matcher import PhraseMatcher
from src.Logic.filter import RECIPE_BASE

try:
    nlp = spacy.load("pl_core_news_sm")
except OSError:
    import os
    os.system("python -m spacy download pl_core_news_sm")
    nlp = spacy.load("pl_core_news_sm")

def get_allowed_ingredients():
    """Tworzy zbiór unikalnych składników występujących w bazie przepisów."""
    allowed = set()
    for przepis in RECIPE_BASE:
        for skladenik in przepis["skladniki"]:
            # Przechowujemy lematy (formy podstawowe), aby łatwiej dopasować
            allowed.add(skladenik.lower())
    return allowed

def filter_ingredients(text: str) -> list[str]:
    """
    Analizuje tekst i zwraca tylko te składniki, które 
    faktycznie istnieją w naszej bazie przepisów.
    """
    if not text:
        return []

    doc = nlp(text.lower())
    allowed_list = get_allowed_ingredients()
    caught_ingredients = []

    for token in doc:
        # Sprawdzamy formę podstawową (lemat) każdego słowa
        lemma = token.lemma_.lower()
        
        # LOGIKA FILTRA:
        # 1. Czy to rzeczownik (NOUN) lub przymiotnik (ADJ - np. 'mielone')?
        # 2. Czy to słowo znajduje się w naszej liście dozwolonych składników?
        if token.pos_ in ["NOUN", "ADJ"] and lemma in allowed_list:
            caught_ingredients.append(lemma)

    # Usuwamy duplikaty (np. jeśli ktoś powiedział 'jajka i jajko')
    return list(set(caught_ingredients))
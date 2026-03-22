import spacy

try:
    nlp = spacy.load("pl_core_news_sm")
except OSError:
    print("Brak modelu pl_core_news_sm.")
    raise

def filter_ingredients(text: str) -> list[str]:
    doc = nlp(text)
    caught_ingredients = []
    for w in doc:
        if w.pos_ == "NOUN":
            lemma = w.lemma_.lower()
            caught_ingredients.append(lemma)
    return caught_ingredients
def filter_recipes(ingredients ,recipes):
    ingredients_without_duplicates = set(ingredients)
    matching_recipes = []
    for recipe in recipes:
        if ingredients_without_duplicates.issubset(recipe["skladniki"]):
            matching_recipes.append(recipe)
    return matching_recipes
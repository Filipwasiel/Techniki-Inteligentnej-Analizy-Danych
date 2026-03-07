import os
import json

SETTINGS_FILE = 'settings.json'

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r', encoding='UTF-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {'font_size': "11", 'marign': "2.0"}

def save_settings(settings):
    with open(SETTINGS_FILE, 'w', encoding='UTF-8') as f:
        json.dump(settings, f)
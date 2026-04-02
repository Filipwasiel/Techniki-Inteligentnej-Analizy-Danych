import streamlit as st
from faster_whisper import WhisperModel
from deep_translator import GoogleTranslator
import os

# Wyciszenie ostrzeżeń o symlinkach (ważne na Windows)
os.environ["HF_HUB_DISABLE_SYMLINKS_WARNING"] = "1"

@st.cache_resource
def load_whisper_model(model_size: str):
    print(f"--- ŁADOWANIE MODELU: {model_size} ---")
    return WhisperModel(model_size, device="cpu", compute_type="int8")

def get_transcription(audio_path: str, model_size: str = "turbo"):
    model = load_whisper_model(model_size)
    
    segments, info = model.transcribe(audio_path, beam_size=1, best_of=1)
    text = " ".join([s.text for s in segments]).strip()
    
    return {
        "tekst": text,
        "jezyk": info.language
    }

def get_translation(text: str, target_lang: str = 'en'):
    if not text:
        return ""
    try:
        translated = GoogleTranslator(source='auto', target=target_lang).translate(text)
        return translated
    except Exception as e:
        return f"Błąd tłumaczenia: {e}"
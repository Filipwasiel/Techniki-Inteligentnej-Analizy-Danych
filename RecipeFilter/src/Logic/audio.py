from faster_whisper import WhisperModel
from deep_translator import GoogleTranslator # Darmowe i nielimitowane dla małych tekstów
import os

os.environ["HF_HUB_DISABLE_SYMLINKS_WARNING"] = "1"
model = WhisperModel("turbo", device="cpu", compute_type="int8")

def get_transcription(audio_path: str):
    """Whisper zajmuje się tylko zamianą mowy na tekst (zawsze oryginał)."""
    segments, info = model.transcribe(audio_path, beam_size=5)
    text = " ".join([s.text for s in segments]).strip()
    return {"tekst": text, "jezyk": info.language}

def get_translation(text: str, target_lang: str = 'en'):
    """Nowa funkcja: tłumaczy gotowy tekst na dowolny język."""
    if not text:
        return ""
    try:
        translated = GoogleTranslator(source='auto', target=target_lang).translate(text)
        return translated
    except Exception as e:
        return f"Błąd tłumaczenia: {e}"
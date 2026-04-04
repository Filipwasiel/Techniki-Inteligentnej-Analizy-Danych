import sys
import streamlit as st
from faster_whisper import WhisperModel
from deep_translator import GoogleTranslator
import os

os.environ["HF_HUB_DISABLE_SYMLINKS_WARNING"] = "1"
if sys.platform == 'win32':
    site_packages = os.path.join(sys.prefix, 'Lib', 'site-packages')
    cublas_path = os.path.join(site_packages, 'nvidia', 'cublas', 'bin')
    cudnn_path = os.path.join(site_packages, 'nvidia', 'cudnn', 'bin')
    
    if os.path.exists(cublas_path) and cublas_path not in os.environ["PATH"]:
        os.environ["PATH"] = cublas_path + os.pathsep + os.environ["PATH"]
    if os.path.exists(cudnn_path) and cudnn_path not in os.environ["PATH"]:
        os.environ["PATH"] = cudnn_path + os.pathsep + os.environ["PATH"]

@st.cache_resource
def load_whisper_model(model_size: str, device: str = "cpu"):
    """Load WhisperModel with given size and device. Cached per (model_size, device)."""
    print(f"--- ŁADOWANIE MODELU: {model_size} (device={device}) ---")
    # Map UI device to faster-whisper device string
    fw_device = "cuda" if device in ("gpu", "cuda", "GPU", "Cuda") else "cpu"
    # Choose compute_type: int8 for CPU, float16 for GPU
    compute_type = "int8" if fw_device == "cpu" else "float16"
    return WhisperModel(model_size, device=fw_device, compute_type=compute_type)

def get_transcription(audio_path: str, model_size: str = "turbo", device: str = "cpu"):
    model = load_whisper_model(model_size, device=device)
    
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
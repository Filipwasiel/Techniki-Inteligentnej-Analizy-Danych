import streamlit as st
import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../..')))

from src.Logic.audio import get_transcription, get_translation 
from src.Logic.nlp import filter_ingredients
from src.Logic.filter import filter_recipes, RECIPE_BASE

st.set_page_config(
    page_title="Inteligentne Przepisy", 
    page_icon="🍳", 
    layout="centered"
)

st.title("Wyszukiwarka Przepisów")

with st.sidebar:
    st.header("Ustawienia")
    
    model_options = {
        "Turbo": "turbo",
        "Tiny": "tiny",
        "Base ": "base",
        "Small": "small",
        "Medium": "medium",
        "Large": "large-v3"
    }
    
    selected_label = st.selectbox(
        "Wybierz model rozpoznawania:",
        options=list(model_options.keys()),
        index=0
    )
    model_size = model_options[selected_label]
    
    st.markdown("---")
    tlumacz = st.checkbox("🇬🇧 Dodaj tłumaczenie na angielski")
    
    st.markdown("---")
    st.markdown("**Dostępne przepisy w bazie:**")
    for p in RECIPE_BASE:
        st.markdown(f"- {p['nazwa']}")
    
    st.markdown("---")
    st.caption("Infrastruktura: west-germany")

# 3. WYBÓR ŹRÓDŁA DŹWIĘKU
tab1, tab2 = st.tabs(["🎤 Nagraj mowę", "📁 Wczytaj plik audio"])
audio_source = None

with tab1:
    audio_mic = st.audio_input("Naciśnij ikonę mikrofonu i wymień składniki:")
    if audio_mic: 
        audio_source = audio_mic

with tab2:
    audio_file = st.file_uploader("Wybierz plik audio (wav, mp3, m4a):", type=["wav", "mp3", "m4a"])
    if audio_file: 
        audio_source = audio_file

# 4. GŁÓWNA LOGIKA PRZETWARZANIA
if audio_source is not None:
    temp_path = "temp_audio.wav"
    
    with open(temp_path, "wb") as f:
        f.write(audio_source.getbuffer())

    try:
        # --- KROK 1: TRANSKRYPCJA ---
        with st.spinner(f"🤖 Whisper ({selected_label}): Rozpoznawanie mowy..."):
            # Przekazujemy wybrany rozmiar modelu do logiki audio
            wynik_pl = get_transcription(temp_path, model_size=model_size)
        
        st.markdown(f"### 🌐 Wyniki ({wynik_pl['jezyk'].upper()}):")
        st.info(f"**Tekst oryginalny:** {wynik_pl['tekst']}")

        # --- KROK 2: TŁUMACZENIE (Opcjonalnie) ---
        if tlumacz:
            with st.spinner("🇬🇧 GoogleTranslate: Tłumaczenie..."):
                # Przesyłamy tekst, nie plik (zgodnie z nowym audio.py)
                tekst_en = get_translation(wynik_pl['tekst'], target_lang='en')
                
                st.markdown("### 🇬🇧 English Translation:")
                st.success(f"**Translated text:** {tekst_en}")

        # --- KROK 3: ANALIZA NLP (Składniki) ---
        with st.spinner("🧠 Model NLP: Analiza składników..."):
            skladniki = filter_ingredients(wynik_pl['tekst'])

        # --- KROK 4: PREZENTACJA WYNIKÓW ---
        if skladniki:
            st.markdown("### 🛒 Wykryte składniki:")
            unikalne = list(set(skladniki))
            st.success(", ".join(unikalne).title())

            znalezione = filter_recipes(unikalne, RECIPE_BASE)
            
            st.markdown("### 🍽️ Pasujące przepisy:")
            if znalezione:
                for p in znalezione:
                    with st.expander(f"✅ {p['nazwa']}"):
                        st.write(f"**Potrzebne składniki:** {', '.join(p['skladniki'])}")
                        st.markdown(f"**Instrukcja:** \n {p['instrukcja']}")
            else:
                st.warning("Nie znalazłem przepisu zawierającego te składniki.")
        else:
            st.error("Nie wykryto żadnych znanych składników. Spróbuj jeszcze raz.")

    except Exception as e:
        st.error(f"Wystąpił błąd: {e}")
    
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

st.markdown("---")
st.caption("© 2026 Inteligentne Przepisy | Tryb: " + selected_label)
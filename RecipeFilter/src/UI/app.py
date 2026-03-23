import streamlit as st
import os
import sys

# Dodajemy główny folder projektu do ścieżki systemowej,
# aby Python widział folder 'src' i nasze importy działały poprawnie.
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../..')))

from src.Logic.audio import transcribe_audio
from src.Logic.nlp import filter_ingredients
from src.Logic.filter import filter_recipes, RECIPE_BASE

# Konfiguracja strony
st.set_page_config(page_title="Inteligentne Przepisy", page_icon="🍳", layout="centered")

st.title("🍳 Inteligentna Wyszukiwarka Przepisów")
st.markdown("Opowiedz mi, jakie masz składniki, a znajdę dla Ciebie idealny przepis!")

# Panel boczny z opcjami "na ocenę 5"
with st.sidebar:
    st.header("⚙️ Ustawienia")
    tlumacz = st.checkbox("🇬🇧 Przetłumacz na angielski (Whisper)")
    st.markdown("---")
    st.markdown("**Baza zawiera przepisy:**")
    for p in RECIPE_BASE:
        st.markdown(f"- {p['nazwa']}")

# Wbudowany widżet Streamlit do nagrywania mowy z mikrofonu
audio_value = st.audio_input("Naciśnij ikonę mikrofonu i powiedz, co masz w lodówce:")

if audio_value is not None:
    st.success("Nagranie zarejestrowane! Przetwarzam...")

    # Zapisujemy nagranie z przeglądarki do tymczasowego pliku .wav
    temp_audio_path = "temp_recording.wav"
    with open(temp_audio_path, "wb") as f:
        f.write(audio_value.getbuffer())

    with st.spinner("🤖 Model Whisper analizuje mowę..."):
        # KROK 1: Rozpoznawanie mowy (z ewentualnym tłumaczeniem)
        wynik_audio = transcribe_audio(temp_audio_path)

    st.markdown("### 🎙️ Wyniki transkrypcji:")
    col1, col2 = st.columns(2)
    with col1:
        st.info(f"**Wykryty język:** `{wynik_audio['jezyk'].upper()}`")
    with col2:
        st.info(f"**Rozpoznany tekst:** {wynik_audio['tekst']}")

    with st.spinner("🧠 Model NLP wyciąga składniki..."):
        # KROK 2: Analiza tekstu
        # Jeśli tłumaczyliśmy na angielski, polski model NLP może zgłupieć,
        # więc do wyciągania składników bezpieczniej jest nie używać przetłumaczonego tekstu
        # (Chyba że chcesz w projekcie analizować też angielski, to temat na małą rozbudowę).
        # Załóżmy na razie, że NLP przetwarza tekst oryginalny:
        if tlumacz:
            wynik_oryginalny = transcribe_audio(temp_audio_path)
            skladniki = filter_ingredients(wynik_oryginalny['tekst'])
        else:
            skladniki = filter_ingredients(wynik_audio['tekst'])

    st.markdown("### 🛒 Wykryte składniki:")
    if skladniki:
        # Usuwamy ewentualne duplikaty na poziomie wyświetlania
        unikalne_skladniki = list(set(skladniki))
        st.success(", ".join(unikalne_skladniki).title())

        # KROK 3: Filtrowanie przepisów
        znalezione = filter_recipes(unikalne_skladniki, RECIPE_BASE)

        st.markdown("### 🍽️ Pasujące przepisy:")
        if znalezione:
            for przepis in znalezione:
                with st.expander(f"✅ {przepis['nazwa']}"):
                    st.write(f"**Potrzebne składniki:** {', '.join(przepis['skladniki'])}")
                    st.write(f"**Instrukcja:** {przepis['instrukcja']}")
        else:
            st.warning("Niestety, nie znalazłem przepisu zawierającego WSZYSTKIE te składniki.")
    else:
        st.error("Nie usłyszałem żadnych konkretnych składników. Spróbuj jeszcze raz!")

    # Sprzątanie tymczasowego pliku
    if os.path.exists(temp_audio_path):
        os.remove(temp_audio_path)
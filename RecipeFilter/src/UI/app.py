import streamlit as st
import os
import sys

# Dodajemy główny folder projektu do ścieżki systemowej,
# aby Python widział folder 'src' i nasze importy działały poprawnie.
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../..')))

# Importujemy zoptymalizowane funkcje z logiki
from src.Logic.audio import get_transcription, get_translation 
from src.Logic.nlp import filter_ingredients
from src.Logic.filter import filter_recipes, RECIPE_BASE

# 1. KONFIGURACJA STRONY
st.set_page_config(
    page_title="Inteligentne Przepisy", 
    page_icon="🍳", 
    layout="centered"
)

# Stylizacja nagłówka
st.title("🍳 Inteligentna Wyszukiwarka Przepisów")
st.markdown("Opowiedz mi lub prześlij nagranie ze składnikami, a znajdę dla Ciebie przepis!")

# 2. PANEL BOCZNY (Sidebar)
with st.sidebar:
    st.header("⚙️ Ustawienia")
    tlumacz = st.checkbox("🇬🇧 Dodaj tłumaczenie na angielski")
    st.markdown("---")
    st.markdown("**Dostępne przepisy w bazie:**")
    for p in RECIPE_BASE:
        st.markdown(f"- {p['nazwa']}")
    st.markdown("---")
    st.caption("Infrastruktura: west-germany")

# 3. WYBÓR ŹRÓDŁA DŹWIĘKU (Tabs)
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
    
    # Zapisujemy bufor do pliku tymczasowego
    with open(temp_path, "wb") as f:
        f.write(audio_source.getbuffer())

    try:
        # --- KROK 1: TRANSKRYPCJA (Oryginał dla NLP) ---
        with st.spinner("🤖 Whisper: Rozpoznawanie mowy..."):
            wynik_pl = get_transcription(temp_path)
        
        # Wyświetlanie wyników transkrypcji
        st.markdown(f"### 🌐 Wyniki ({wynik_pl['jezyk'].upper()}):")
        st.info(f"**Tekst oryginalny:** {wynik_pl['tekst']}")

        # --- KROK 2: TŁUMACZENIE (Opcjonalnie) ---
        if tlumacz:
            with st.spinner("🇬🇧 Whisper: Tłumaczenie na angielski..."):
                odpowiedz_en = get_translation(wynik_pl['tekst'], target_lang='en')
                
                # Obsługa formatu zwrotnego (słownik lub string)
                tekst_en = odpowiedz_en.get('tekst', '') if isinstance(odpowiedz_en, dict) else odpowiedz_en
                
                st.markdown("### 🇬🇧 English Translation:")
                st.success(f"**Translated text:** {tekst_en}")

        # --- KROK 3: ANALIZA NLP (Składniki) ---
        with st.spinner("🧠 Model NLP: Analiza składników..."):
            # Analizujemy zawsze tekst oryginalny, bo spaCy jest ustawiony na PL
            skladniki = filter_ingredients(wynik_pl['tekst'])

        # --- KROK 4: PREZENTACJA WYNIKÓW ---
        if skladniki:
            st.markdown("### 🛒 Wykryte składniki:")
            unikalne = list(set(skladniki))
            # Wyświetlamy sformatowaną listę
            st.success(", ".join(unikalne).title())

            # Filtrowanie przepisów
            znalezione = filter_recipes(unikalne, RECIPE_BASE)
            
            st.markdown("### 🍽️ Pasujące przepisy:")
            if znalezione:
                for p in znalezione:
                    with st.expander(f"✅ {p['nazwa']}"):
                        st.write(f"**Potrzebne składniki:** {', '.join(p['skladniki'])}")
                        st.markdown(f"**Instrukcja:** \n {p['instrukcja']}")
            else:
                st.warning("Nie znalazłem przepisu zawierającego te składniki w naszej bazie.")
        else:
            st.error("Nie wykryto żadnych znanych składników. Spróbuj powiedzieć np. 'mam jajka, mleko i mąkę'.")

    except Exception as e:
        st.error(f"Wystąpił krytyczny błąd podczas przetwarzania: {e}")
    
    finally:
        # Sprzątanie - usuwamy plik tymczasowy po zakończeniu pracy
        if os.path.exists(temp_path):
            os.remove(temp_path)

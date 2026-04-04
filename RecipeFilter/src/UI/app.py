import streamlit as st
import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../..')))

from src.Logic.audio import get_transcription, get_translation
from src.Logic.nlp_bilingual import filter_ingredients
from src.Logic.filter_bilingual import filter_recipes, RECIPE_BASE, get_recipe_base

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
    # Device selection for transcription (CPU or GPU)
    device_choice = st.selectbox("Urządzenie do transkrypcji:", options=["CPU", "GPU"], index=0)
    device = 'gpu' if device_choice == 'GPU' else 'cpu'

    st.markdown("---")
    tlumacz = st.checkbox("🇬🇧 Dodaj tłumaczenie na angielski")

    st.markdown("---")
    st.markdown("**Dostępne przepisy w bazie:**")
    for p in RECIPE_BASE:
        st.markdown(f"- {p['nazwa']}")
    
    st.markdown("---")

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
        with st.spinner(f"🤖 Whisper ({selected_label}) on {device_choice}: Rozpoznawanie mowy..."):
            # Przekazujemy wybrany rozmiar modelu i urządzenie do logiki audio
            wynik_pl = get_transcription(temp_path, model_size=model_size, device=device)
        
        st.markdown(f"### 🌐 Wyniki ({wynik_pl['jezyk'].upper()}):")
        st.info(f"**Tekst oryginalny:** {wynik_pl['tekst']}")

        # --- KROK 2: TŁUMACZENIE ---
        detected_lang = wynik_pl.get('jezyk','').lower() if isinstance(wynik_pl.get('jezyk',''), str) else ''
        tekst_pl = wynik_pl['tekst']
        if not detected_lang.startswith('pl'):
            with st.spinner("🇵🇱 GoogleTranslate: Tłumaczenie na polski..."):
                tekst_pl = get_translation(wynik_pl['tekst'], target_lang='pl')
        st.markdown("### 🇵🇱 Polish (for filtering):")
        st.info(f"**Polish text:** {tekst_pl}")

        tekst_en = None
        if tlumacz:
            with st.spinner("🇬🇧 GoogleTranslate: Tłumaczenie na angielski..."):
                tekst_en = get_translation(tekst_pl, target_lang='en')
                st.markdown("### 🇬🇧 English Translation:")
                st.success(f"**Translated text:** {tekst_en}")

        # --- KROK 3: ANALIZA NLP (Składniki) ---
        with st.spinner("🧠 Model NLP: Analiza składników..."):
            if tlumacz and tekst_en:
                base_en = get_recipe_base('en')
                skladniki = filter_ingredients(tekst_en, recipes=base_en, lang='en')
                exact_matches, partial_matches = filter_recipes(skladniki, base_en)
            else:
                skladniki = filter_ingredients(tekst_pl, recipes=RECIPE_BASE, lang='pl')
                exact_matches, partial_matches = filter_recipes(skladniki, RECIPE_BASE)

        # --- KROK 4: PREZENTACJA WYNIKÓW ---
        if skladniki:
            st.markdown("### 🛒 Wykryte składniki:")
            unikalne = list(set(skladniki))
            st.success(", ".join(unikalne).title())

            st.markdown("### 🍽️ Pasujące przepisy (pełne dopasowania):")
            if exact_matches:
                for p in exact_matches:
                    with st.expander(f"✅ {p['nazwa']}"):
                        st.write(f"**Potrzebne składniki:** {', '.join(sorted(list(p['skladniki'])))}")
                        # show PL and EN instructions when available
                        if tlumacz:
                            pl_p = next((r for r in RECIPE_BASE if r['id'] == p['id']), None)
                            if pl_p:
                                st.markdown(f"**Instrukcja (PL):** \n {pl_p['instrukcja']}")
                            st.markdown(f"**Instruction (EN):** \n {p.get('instrukcja', '')}")
                        else:
                            st.markdown(f"**Instrukcja:** \n {p['instrukcja']}")
            else:
                st.info("Brak pełnych dopasowań.")

            st.markdown("### 🍽️ Przepisy częściowo dopasowane (brakujące składniki):")
            if partial_matches:
                for item in partial_matches:
                    r = item['recipe']
                    missing = item['missing']
                    found = item['found']
                    with st.expander(f"⚠️ {r['nazwa']} — brak {len(missing)} składników"):
                        st.write(f"Znalezione składniki: {', '.join(found)}")
                        st.write(f"Brakuje: {', '.join(missing)}")
                        if tlumacz:
                            pl_p = next((x for x in RECIPE_BASE if x['id'] == r['id']), None)
                            if pl_p:
                                st.markdown(f"**Instrukcja (PL):** \n {pl_p['instrukcja']}")
                            st.markdown(f"**Instruction (EN):** \n {r.get('instrukcja', '')}")
                        else:
                            st.markdown(f"**Instrukcja:** \n {r['instrukcja']}")
            else:
                st.info("Brak częściowych dopasowań.")
        else:
            st.error("Nie wykryto żadnych znanych składników. Spróbuj jeszcze raz.")

    except Exception as e:
        st.error(f"Wystąpił błąd: {e}")
    
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

st.markdown("---")
st.caption("© 2026 Inteligentne Przepisy | Tryb: " + selected_label)
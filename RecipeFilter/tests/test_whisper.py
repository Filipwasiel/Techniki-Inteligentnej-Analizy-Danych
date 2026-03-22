import unittest
import os
from src.Logic.audio import transcribe_audio


class TestWhisperModel(unittest.TestCase):

    def setUp(self):
        """
        Ta metoda uruchamia się przed każdym testem.
        Budujemy tu bezpieczną ścieżkę do naszego pliku audio.
        """
        # Pobiera ścieżkę do folderu 'tests'
        katalog_testow = os.path.dirname(__file__)
        # Dokleja folder 'data' i nazwę pliku
        self.sciezka_audio = os.path.join(katalog_testow, "data", "Recording0004.wav")

    def test_prawdziwa_transkrypcja_whisper(self):
        """Uruchamia prawdziwy model Whisper na pliku audio."""

        # Zabezpieczenie: jeśli zapomnisz dodać pliku, test się pominie, a nie wyrzuci błąd
        if not os.path.exists(self.sciezka_audio):
            self.skipTest(f"Brak pliku {self.sciezka_audio}. Dodaj go, aby uruchomić ten test.")

        print("\n⏳ Przetwarzam prawdziwe audio... to może potrwać kilka sekund.")

        # Odpalamy naszą funkcję z PRAWDZIWYM plikiem
        wynik = transcribe_audio(self.sciezka_audio)

        # SPRAWDZAMY WYNIKI:
        # 1. Sprawdzamy, czy Whisper rozpoznał język polski
        self.assertEqual(wynik["jezyk"], "pl")

        # 2. Sprawdzamy, czy wyłapał kluczowe słowo (zamieniamy na małe litery, bo Whisper często zaczyna zdanie wielką)
        tekst_malymi_literami = wynik["tekst"].lower()
        self.assertIn("jajko", tekst_malymi_literami)
        self.assertIn("chleb", tekst_malymi_literami)


if __name__ == "__main__":
    unittest.main()
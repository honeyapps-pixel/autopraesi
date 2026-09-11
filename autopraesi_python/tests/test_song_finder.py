"""Regressionstests für die Liedsuche (song_finder.find_song_detailed).

Ausführen:  cd autopraesi_python && venv/bin/python -m unittest -v

Hintergrund: Am 13.9.2026 wurde für "FJ 3 87 - Ich bin bei dir" die Datei
FJ1/87 - Lord you have my heart.pptx eingefügt, weil "FJ 3" (mit Leerzeichen)
keinem Ordner zugeordnet wurde und der Fallback den ersten Treffer mit Nr. 87
in irgendeinem Ordner nahm. Diese Tests sichern das Verhalten dauerhaft ab.
"""
import os
import sys
import unicodedata
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excel_reader import parse_song_entry  # noqa: E402
from song_finder import find_song, find_song_detailed  # noqa: E402

LIB = "/Gemeinde/Brüderrecords°"


def _index(**dirs):
    """Baut einen Song-Index wie build_song_index(): {ordner: [(name, pfad)]}."""
    return {d: [(n, f"{LIB}/{d}/{n}") for n in names] for d, names in dirs.items()}


INDEX = _index(
    FJ1=["87 - Lord you have my heart.pptx",
         "87 - Lord you have my heart - deutsch.pptx",
         "29 - Hab dank von Herzen Herr.pptx",
         "140 - Dass dein Wort in meinem Herzen.pptx",
         "140 - Dass dein Wort in meinem Herzen starke Wurzeln schlägt.pptx",
         "12-3 Worte.pptx",
         "121- Komm in unsere Mitte.pptx",
         "167 -Ich will dir gehorchen FJ1.pptx"],
    FJ2=["87 - Ganz anderes Lied.pptx", "12 - Lobe den Herrn.pptx"],
    FJ3=["87 - Ich bin bei dir.pptx"],
    FJ4=["87 - Noch ein anderes.pptx"],
    IWDD=["446 - Zünde an dein Feuer.pptx"],   # NFD-Umlaut wie von macOS/Dropbox
    Lobpreisstrophen=["FJ1 95 - Mein Herz, meinen Leib.pptx",
                      "FJ1 29 - Hab dank von Herzen Herr.pptx"],
    Kinderlieder=["Von oben, von unten.pptx"],
)
INDEX["Sonstige Lieder"] = [("Höre mich.pptx", f"{LIB}/Sonstige Lieder/Höre mich.pptx"),
                            ("Dreimal.pptx", f"{LIB}/Sonstige Lieder/Dreimal.pptx")]


def _find(raw, col_b=""):
    return find_song_detailed(parse_song_entry(raw, "song1", col_b=col_b), INDEX)


def _name(match):
    return unicodedata.normalize("NFC", os.path.basename(match.path)) if match.path else None


class BookNormalization(unittest.TestCase):
    """Buch mit/ohne Leerzeichen muss denselben Ordner treffen."""

    def test_book_with_space_hits_correct_folder(self):
        m = _find("FJ 3 87 - Ich bin bei dir")
        self.assertEqual(_name(m), "87 - Ich bin bei dir.pptx")
        self.assertIn("/FJ3/", m.path)
        self.assertEqual(m.warning, "")

    def test_book_without_space(self):
        m = _find("FJ3 87 - Ich bin bei dir")
        self.assertIn("/FJ3/", m.path)

    def test_lowercase_book_with_space(self):
        m = _find("fj 3 87 - Ich bin bei dir")
        self.assertIn("/FJ3/", m.path)

    def test_lobpreisstrophe_book_with_space(self):
        m = _find("Lobpreisstrophe: FJ 1 95 - Mein Herz, meinen Leib")
        self.assertEqual(_name(m), "FJ1 95 - Mein Herz, meinen Leib.pptx")
        self.assertIn("/Lobpreisstrophen/", m.path)

    def test_lobpreisstrophe_falls_back_to_full_song_in_book_folder(self):
        idx = _index(FJ1=["95 - Mein Herz, meinen Leib.pptx"], Lobpreisstrophen=[])
        m = find_song_detailed(parse_song_entry("Lobpreisstrophe: FJ 1 95 - Mein Herz, meinen Leib"), idx)
        self.assertIn("/FJ1/95 - ", m.path)

    def test_lobpreisstrophe_via_column_b(self):
        m = _find("FJ 1 29 - Hab Dank", col_b="Lobpreisstrophe")
        self.assertIn("/Lobpreisstrophen/", m.path)
        self.assertEqual(_name(m), "FJ1 29 - Hab dank von Herzen Herr.pptx")


class TitleDisambiguation(unittest.TestCase):
    """Bei mehreren Dateien mit gleicher Nummer entscheidet der Titel."""

    def test_number_only_never_picks_song_from_wrong_book(self):
        # Unbekanntes Buch: Nummer allein darf NICHT irgendein 87er liefern
        m = _find("XYZ 87 - Gibt es nicht")
        self.assertIsNone(m.path)

    def test_typo_in_book_is_rescued_by_full_title_match(self):
        m = _find("FJ1 87 - Ich bin bei dir")
        self.assertIn("/FJ3/", m.path)
        self.assertIn("anderem Ordner", m.warning)

    def test_german_title_in_excel_english_files_warns(self):
        m = _find("FJ1 87 - Herr, du hast mein Herz")
        self.assertIn("/FJ1/", m.path)          # Nummer im Buch ist verbindlich …
        self.assertIn("passt nicht", m.warning)  # … aber sichtbar prüfen lassen

    def test_english_title_prefers_exact_file_and_hints_alternative(self):
        m = _find("FJ1 87 - Lord you have my heart")
        self.assertEqual(_name(m), "87 - Lord you have my heart.pptx")
        self.assertIn("deutsch", m.warning)

    def test_deutsch_suffix_selects_german_version(self):
        m = _find("FJ1 87 - Lord you have my heart - deutsch")
        self.assertEqual(_name(m), "87 - Lord you have my heart - deutsch.pptx")
        self.assertEqual(m.warning, "")

    def test_longer_title_selects_longer_file(self):
        m = _find("FJ1 140 - Dass dein Wort in meinem Herzen starke Wurzeln schlägt")
        self.assertEqual(_name(m), "140 - Dass dein Wort in meinem Herzen starke Wurzeln schlägt.pptx")
        self.assertEqual(m.warning, "")

    def test_generic_fj_book_searches_all_fj_folders(self):
        m = _find("FJ 87 - Ich bin bei dir")
        self.assertIn("/FJ3/", m.path)
        self.assertEqual(m.warning, "")


class FilenameVariants(unittest.TestCase):
    """Alle in der Bibliothek vorkommenden Nummern-Schreibweisen."""

    def test_dash_without_spaces_is_not_confused_with_range(self):
        # FJ1/"12-3 Worte" ist nicht Nr. 12 → stattdessen Titel-Treffer in FJ2
        m = _find("FJ1 12 - Lobe den Herrn")
        self.assertEqual(_name(m), "12 - Lobe den Herrn.pptx")
        self.assertIn("/FJ2/", m.path)

    def test_number_dash_space(self):
        self.assertEqual(_name(_find("FJ1 121 - Komm in unsere Mitte")),
                         "121- Komm in unsere Mitte.pptx")

    def test_number_space_dash_nospace(self):
        self.assertEqual(_name(_find("FJ1 167 - Ich will dir gehorchen")),
                         "167 -Ich will dir gehorchen FJ1.pptx")

    def test_umlaut_nfd_filename_matches_nfc_excel(self):
        m = _find("IWDD 446 - Zünde an dein Feuer")
        self.assertEqual(_name(m), "446 - Zünde an dein Feuer.pptx")
        self.assertEqual(m.warning, "")


class TitleOnlySongs(unittest.TestCase):
    def test_kinderlied_by_title(self):
        m = _find("Kinderlied: Von oben, von unten")
        self.assertEqual(_name(m), "Von oben, von unten.pptx")

    def test_sonstige_lieder_with_umlaut(self):
        m = _find("Sonstige Lieder - Höre mich")
        self.assertEqual(_name(m), "Höre mich.pptx")

    def test_whole_word_matching_avoids_substring_false_positive(self):
        # "ich bin bei" darf nicht über Teilwörter ("nicht", "sinnbild", "beim") treffen
        idx = _index(**{"Sonstige Lieder": ["Nicht beim Sinnbild.pptx"]})
        m = find_song_detailed(parse_song_entry("Sonstige Lieder - Ich bin bei dir"), idx)
        self.assertIsNone(m.path)

    def test_find_song_wrapper_returns_path(self):
        self.assertEqual(find_song(parse_song_entry("FJ 3 87 - Ich bin bei dir"), INDEX),
                         f"{LIB}/FJ3/87 - Ich bin bei dir.pptx")

    def test_empty_entry(self):
        self.assertIsNone(find_song(parse_song_entry(""), INDEX))


if __name__ == "__main__":
    unittest.main()

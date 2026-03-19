"""Konfiguration und Konstanten für AutoPräsi."""

import os

# Basis-Pfade
HOME = os.path.expanduser("~")
DROPBOX_GEMEINDE = os.path.join(HOME, "Dropbox", "Gemeinde")
DESKTOP_AUTOPRAESI = os.path.join(HOME, "Desktop", "Autopräsi")

# Template
TEMPLATE_PATH = os.path.join(
    DESKTOP_AUTOPRAESI, "v_1_4_3_ab2026",
    "Vorlage_Godi_Standard_Sonntag_v1_4_3.pptx"
)

# Lied-Bibliothek
SONG_LIBRARY = os.path.join(DROPBOX_GEMEINDE, "Brüderrecords°")

# Bild-Verzeichnis
IMAGE_DIR = DROPBOX_GEMEINDE

# Ausgabe-Verzeichnisse
OUTPUT_DIR_DESKTOP = DROPBOX_GEMEINDE
OUTPUT_DIR_DROPBOX = DROPBOX_GEMEINDE

# GoDi-Plan Excel Verzeichnis
GODI_PLAN_DIR = DROPBOX_GEMEINDE

# Logging
LOG_DIR = os.path.join(DESKTOP_AUTOPRAESI, "autopraesi_python", "logs")

# --- Excel Zeilen-Mappings (1-indexed) ---

# Standard-Gottesdienst Zeilen
EXCEL_ROWS = {
    "header": (1, 1),           # Row 1, Col A - "Gottesdienst am 08.03.2026 (Okuli)"
    "theme": (3, 4),            # Row 3, Col D - Thema
    "godi_leitung": (4, 4),     # Row 4, Col D - GoDi-Leitung Name
    "bild_vorhanden": (4, 7),   # Row 4, Col G - "X" wenn Bild ausgesucht
    "predigt_vorschlaege": (5, 4),  # Row 5, Col D
    "verkuendigung1": (6, 4),   # Row 6, Col D - Prediger 1
    "verkuendigung2": (6, 6),   # Row 6, Col F - Prediger 2
    "lesung_referenz": (7, 4),  # Row 7, Col D - Lesungstext Bibelstelle
    "lesung_person": (7, 6),    # Row 7, Col F - Wer liest
    "begruessung": (16, 4),     # Row 16, Col D - Begrüßungsvers
    "song1": (19, 4),           # Row 19, Col D - Gemeindelied 1
    "song2": (21, 4),           # Row 21, Col D - Lobpreisstrophe
    "song3": (24, 4),           # Row 24, Col D - Kinderlied
    "predigt1_referenz": (26, 4),   # Row 26, Col D - Predigt 1 Bibelstelle
    "predigt1_titel": (26, 5),      # Row 26, Col E - Predigt 1 Thema
    "song4": (27, 4),           # Row 27, Col D - Gemeindelied
    "song5": (32, 4),           # Row 32, Col D - Gemeindelied
    "predigt2_referenz": (33, 4),   # Row 33, Col D - Predigt 2 Bibelstelle
    "predigt2_titel": (33, 5),      # Row 33, Col E - Predigt 2 Thema
    "song6": (34, 4),           # Row 34, Col D - Lobpreisstrophe
    "abkuendigungen_start": (37, 2),  # Row 37 - Beginn Abkündigungen
    "abendmahl": (38, 2),       # Row 38, Col B - Abendmahl-Indikator
    "song7": (47, 4),           # Row 47, Col D - Schlusslied
}

# Abkündigungs-Zeilen (Row 37-45, Cols B-E)
ABKUENDIGUNGEN_ROWS = range(37, 46)

# Einladungs-Zeilen im Excel (Rows 37-44, vor "Ausblick:" in Row 45)
EINLADUNG_ROW_START = 37
EINLADUNG_ROW_END = 44  # inklusive
# Spalten: B=Datum, C=Uhrzeit, D=Veranstaltung, F=Hinweis
EINLADUNG_COLS = {"datum": 2, "uhrzeit": 3, "veranstaltung": 4, "hinweis": 6}

# Template-Folien-Positionen für Lieder (1-indexed)
# Jede Position gibt an, wo die Platzhalter-Folie im Template steht
SONG_TEMPLATE_SLIDES = {
    "song1": 5,     # Template Folie 5
    "song2": 10,    # Template Folie 10
    "song3": 15,    # Template Folie 15
    "song4": 22,    # Template Folie 22
    "song5": 26,    # Template Folie 26
    "song6": 32,    # Template Folie 32
    "song7": 36,    # Template Folie 36
}

# Thema-Folien im Template (1-indexed) - enthalten Platzhalter "Thema" und "XY.XY.XY"
THEME_SLIDES = [4, 6, 9, 11, 13, 17, 21, 23, 25, 27, 31, 33, 35, 37]

# Slots, die immer Lobpreisstrophen sind (auch ohne Prefix im Excel-Eintrag)
LOBPREIS_SLOTS = {"song2", "song6"}

# Togglebare Abschnitte im Template (0-basierte Folien-Indizes)
# Jeder Abschnitt kann in der UI an/abgewählt werden
TOGGLEABLE_SECTIONS = {
    # --- Eröffnung ---
    "begruessung": {
        "label": "Begrüßungsvers",
        "slides": [2],             # Folie 3
    },
    # --- Lieder ---
    "song1": {
        "label": "Lied 1 (Gemeindelied)",
        "slides": [4],             # Folie 5
    },
    "song2": {
        "label": "Lied 2 (Lobpreis)",
        "slides": [9],             # Folie 10
    },
    # song3 ist in "kinderstunde" eingebettet (Folie 15 = Index 14)
    "song4": {
        "label": "Lied 4",
        "slides": [21],            # Folie 22
    },
    "song5": {
        "label": "Lied 5",
        "slides": [25],            # Folie 26
    },
    "song6": {
        "label": "Lied 6 (Lobpreis)",
        "slides": [31],            # Folie 32
    },
    "song7": {
        "label": "Lied 7 (Schlusslied)",
        "slides": [35],            # Folie 36
    },
    # --- Liturgie ---
    "glaubensbekenntnis": {
        "label": "Glaubensbekenntnis",
        "slides": [6, 7],          # Folien 7-8
    },
    "kinderstunde": {
        "label": "Kinderstunde + Kinderlied",
        "slides": [13, 14, 15],    # Folien 14-16 (inkl. Song3)
    },
    "lesung": {
        "label": "Lesung",
        "slides": [11],            # Folie 12
    },
    "predigt1": {
        "label": "Predigt 1",
        "slides": [17, 18, 19],    # Folien 18-20
    },
    "gebetszeit": {
        "label": "Gebetszeit",
        "slides": [23],            # Folie 24
    },
    "predigt2": {
        "label": "Predigt 2",
        "slides": [27, 28, 29],    # Folien 28-30
    },
    # --- Abschluss ---
    "einladung": {
        "label": "Herzliche Einladung / Abkündigungen",
        "slides": [33],            # Folie 34
    },
}

# Standard-Reihenfolge der Abschnitte (für Drag & Drop Umordnung)
DEFAULT_SECTION_ORDER = [
    "begruessung",
    "song1",
    "glaubensbekenntnis",
    "song2",
    "lesung",
    "kinderstunde",      # enthält song3 (Kinderlied)
    "predigt1",
    "song4",
    "gebetszeit",
    "song5",
    "predigt2",
    "song6",
    "einladung",
    "song7",
]

# Jeder Block enthält seine Template-Folien + nachfolgende Thema-Folie
# (0-basierte Indizes)
SECTION_BLOCKS = {
    "begruessung": [2, 3],
    "song1": [4, 5],
    "glaubensbekenntnis": [6, 7, 8],
    "song2": [9, 10],
    "lesung": [11, 12],
    "kinderstunde": [13, 14, 15, 16],  # KS-Intro + Song3 + KS-Outro + Thema
    "predigt1": [17, 18, 19, 20],
    "song4": [21, 22],
    "gebetszeit": [23, 24],
    "song5": [25, 26],
    "predigt2": [27, 28, 29, 30],
    "song6": [31, 32],
    "einladung": [33, 34],
    "song7": [35, 36],
}

# Intro-Folien (immer am Anfang, nicht verschiebbar)
INTRO_SLIDES = [0, 1]

# Maximale Zeichen pro Folie für Bibeltexte (Cambria 32pt)
MAX_CHARS_PER_SLIDE = 500

# Song-Unterordner in Brüderrecords
SONG_DIRS = {
    "FJ1": "FJ1",
    "FJ2": "FJ2",
    "FJ3": "FJ3",
    "FJ4": "FJ4",
    "FJ5": "FJ5",
    "FJ6": "FJ6",
    "GLS": "GLS",
    "SUG": "SUG",
    "IWDD": "IWDD",
    "SGIDH": "SGIDH 2",
    "Loben": "Loben",
    "Kinderlieder": "Kinderlieder",
    "Sonstige Lieder": "Sonstige Lieder",
    "Lobpreisstrophen": "Lobpreisstrophen",
    "Einfach spitze": "Einfach spitze",
}

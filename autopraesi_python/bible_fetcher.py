"""Lädt Bibeltexte von die-bibel.de (Luther 1984)."""

import logging
import re

import requests
from bs4 import BeautifulSoup

log = logging.getLogger(__name__)

# Mapping deutscher Buchnamen auf die-bibel.de URL-Kürzel
BOOK_MAPPING = {
    # Altes Testament
    "1. mose": "GEN", "2. mose": "EXO", "3. mose": "LEV",
    "4. mose": "NUM", "5. mose": "DEU",
    "josua": "JOS", "richter": "JDG", "rut": "RUT",
    "1. samuel": "1SA", "2. samuel": "2SA", "1sam": "1SA", "2sam": "2SA",
    "1. könige": "1KI", "1. kön": "1KI", "1kön": "1KI",
    "2. könige": "2KI", "2. kön": "2KI", "2kön": "2KI",
    "1. chronik": "1CH", "2. chronik": "2CH", "1chr": "1CH", "2chr": "2CH",
    "esra": "EZR", "nehemia": "NEH", "ester": "EST",
    "hiob": "JOB", "psalm": "PSA", "psalmen": "PSA",
    "sprüche": "PRO", "prediger": "ECC",
    "hohelied": "SNG", "hohes lied": "SNG",
    "jesaja": "ISA", "jeremia": "JER",
    "klagelieder": "LAM", "hesekiel": "EZK",
    "daniel": "DAN", "hosea": "HOS", "joel": "JOL",
    "amos": "AMO", "obadja": "OBA", "jona": "JON",
    "micha": "MIC", "nahum": "NAM", "habakuk": "HAB",
    "zefanja": "ZEP", "haggai": "HAG",
    "sacharja": "ZEC", "maleachi": "MAL",
    # Neues Testament
    "matthäus": "MAT", "mt": "MAT", "matt": "MAT",
    "markus": "MRK", "mk": "MRK",
    "lukas": "LUK", "lk": "LUK",
    "johannes": "JHN", "joh": "JHN",
    "apostelgeschichte": "ACT", "apg": "ACT",
    "römer": "ROM", "röm": "ROM",
    "1. korinther": "1CO", "1. kor": "1CO", "1kor": "1CO",
    "2. korinther": "2CO", "2. kor": "2CO", "2kor": "2CO",
    "galater": "GAL", "gal": "GAL",
    "epheser": "EPH", "eph": "EPH",
    "philipper": "PHP", "phil": "PHP",
    "kolosser": "COL", "kol": "COL",
    "1. thessalonicher": "1TH", "1. thess": "1TH", "1thess": "1TH",
    "2. thessalonicher": "2TH", "2. thess": "2TH", "2thess": "2TH",
    "1. timotheus": "1TI", "1. tim": "1TI", "1tim": "1TI",
    "2. timotheus": "2TI", "2. tim": "2TI", "2tim": "2TI",
    "titus": "TIT", "philemon": "PHM",
    "hebräer": "HEB", "hebr": "HEB",
    "jakobus": "JAS", "jak": "JAS",
    "1. petrus": "1PE", "1. petr": "1PE", "1petr": "1PE",
    "2. petrus": "2PE", "2. petr": "2PE", "2petr": "2PE",
    "1. johannes": "1JN", "1. joh": "1JN", "1joh": "1JN",
    "2. johannes": "2JN", "2. joh": "2JN", "2joh": "2JN",
    "3. johannes": "3JN", "3. joh": "3JN", "3joh": "3JN",
    "judas": "JUD",
    "offenbarung": "REV", "offb": "REV",
}

BIBLE_URL = "https://www.die-bibel.de/bibel/LU84/{book}.{chapter}"

# die-bibel.de blockiert generische User-Agents (HTTP 403), daher voller Browser-UA
HTTP_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/126.0.0.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml",
    "Accept-Language": "de-DE,de;q=0.9",
}


def _lookup_book(book_de: str):
    """Schlägt ein Buch im Mapping nach, normalisiert '1 petrus' → '1. petrus' etc."""
    book_de = book_de.rstrip(".").strip()
    book_abbr = BOOK_MAPPING.get(book_de)
    if book_abbr:
        return book_abbr
    # Versuch mit Punkt nach Ziffer: "1 petrus" → "1. petrus"
    normalized = re.sub(r'^(\d)\s*', r'\1. ', book_de)
    book_abbr = BOOK_MAPPING.get(normalized)
    if book_abbr:
        return book_abbr
    # Fuzzy: Prefix-Match
    for key, val in BOOK_MAPPING.items():
        if book_de.startswith(key) or key.startswith(book_de):
            return val
    return None


# Alle Strich-Varianten, die als Vers-Trenner vorkommen können
# (Hyphen, En-Dash, Em-Dash, Figure-Dash, Minus-Sign, Horizontal-Bar)
_DASH_CHARS = "-‐‑‒–—―−"


def parse_reference(reference: str) -> list:
    """Parst eine deutsche Bibelstellen-Referenz.

    Beispiele:
        "Lukas 9, 57-62" → [("LUK", 9, 57, 62)]
        "Epheser 4,23" → [("EPH", 4, 23, 23)]
        "Lukas 18,35-19,10" → [("LUK", 18, 35, None), ("LUK", 19, 1, 10)]

    Returns:
        Liste von (book_abbr, chapter, verse_start, verse_end) Tupeln.
        verse_start=None: ganzes Kapitel. verse_end=None: bis Kapitelende.
    """
    results = []
    parts = reference.split(";")

    for part in parts:
        part = part.strip()
        if not part:
            continue

        part = re.sub(r'\([^)]*\)', '', part).strip()

        match = re.match(
            r'^((?:\d\.?\s*)?[A-Za-zäöüÄÖÜß]+\.?)\s+'
            r'(\d+)'
            r'[,.\s]\s*'
            r'(\d+)'
            rf'(?:\s*[{_DASH_CHARS}]\s*(?:(\d+)\s*[,.]\s*)?(\d+))?',
            part
        )

        if match:
            book_de = match.group(1).strip().lower()
            chapter = int(match.group(2))
            verse_start = int(match.group(3))
            end_chapter = int(match.group(4)) if match.group(4) else None
            verse_end = int(match.group(5)) if match.group(5) else verse_start

            book_abbr = _lookup_book(book_de)

            if not book_abbr:
                log.warning(f"Unbekanntes Buch: {book_de} in '{part}'")
                continue

            if end_chapter is None or end_chapter == chapter:
                results.append((book_abbr, chapter, verse_start, verse_end))
            elif end_chapter > chapter:
                # Kapitelübergreifend, z.B. "Lukas 18,35-19,10":
                # Startkapitel bis Ende, Zwischenkapitel ganz, Endkapitel ab Vers 1
                results.append((book_abbr, chapter, verse_start, None))
                for ch in range(chapter + 1, end_chapter):
                    results.append((book_abbr, ch, None, None))
                results.append((book_abbr, end_chapter, 1, verse_end))
            else:
                log.warning(f"Ungültiger Kapitelbereich: {part}")
            continue

        # Fallback: "Buch Kapitel" ohne Vers (z.B. "Psalm 133")
        match_chapter = re.match(
            r'^((?:\d\.?\s*)?[A-Za-zäöüÄÖÜß]+\.?)\s+(\d+)\s*$',
            part
        )
        if match_chapter:
            book_de = match_chapter.group(1).strip().lower()
            chapter = int(match_chapter.group(2))

            book_abbr = _lookup_book(book_de)

            if book_abbr:
                log.info(f"Ganzes Kapitel wird geladen: {book_de} {chapter}")
                results.append((book_abbr, chapter, None, None))  # None = ganzes Kapitel
            else:
                log.warning(f"Unbekanntes Buch: {book_de} in '{part}'")
        else:
            log.warning(f"Konnte Bibelstelle nicht parsen: {part}")

    return results


def fetch_bible_text(reference: str) -> list:
    """Lädt den Bibeltext für eine gegebene Referenz (Luther 1984).

    Args:
        reference: Deutsche Bibelstellen-Referenz, z.B. "Lukas 9, 57-62"

    Returns:
        Liste von (vers_nummer, vers_text) Tupeln, oder Platzhalter bei Fehler
    """
    if not reference or not reference.strip():
        return []

    refs = parse_reference(reference)
    if not refs:
        log.warning(f"Konnte Bibelstelle nicht parsen: {reference}")
        return [(0, f"[Bibeltext: {reference}]")]

    all_verses = []

    for book_abbr, chapter, verse_start, verse_end in refs:
        try:
            url = BIBLE_URL.format(book=book_abbr, chapter=chapter)
            if verse_start is None:
                log.info(f"Lade Bibeltext: {book_abbr} {chapter} (ganzes Kapitel)")
            elif verse_end is None:
                log.info(f"Lade Bibeltext: {book_abbr} {chapter},{verse_start} bis Kapitelende")
            else:
                log.info(f"Lade Bibeltext: {book_abbr} {chapter},{verse_start}-{verse_end}")

            resp = requests.get(url, timeout=15, headers=HTTP_HEADERS)
            if resp.status_code != 200:
                log.warning(f"HTTP {resp.status_code} für {url}")
                all_verses.append((0, f"[{reference}]"))
                continue

            soup = BeautifulSoup(resp.text, 'html.parser')

            # Alle Verse des Kapitels sammeln (Klassen LU84.{book}.{chapter}.N),
            # danach den gewünschten Bereich auswählen
            verse_dict = {}
            prefix = f"LU84.{book_abbr}.{chapter}."
            for span in soup.find_all(class_=True):
                for cls in span.get("class", []):
                    if cls.startswith(prefix):
                        try:
                            v_num = int(cls[len(prefix):])
                        except ValueError:
                            continue
                        text = re.sub(r'\s+', ' ', span.get_text(strip=False)).strip()
                        if text:
                            if v_num in verse_dict:
                                verse_dict[v_num] += " " + text
                            else:
                                verse_dict[v_num] = text

            for v_num in sorted(verse_dict.keys()):
                if verse_start is not None and v_num < verse_start:
                    continue
                if verse_end is not None and v_num > verse_end:
                    continue
                all_verses.append((v_num, verse_dict[v_num]))

        except requests.RequestException as e:
            log.warning(f"Netzwerkfehler beim Laden von {reference}: {e}")
            all_verses.append((0, f"[{reference}]"))

    if not all_verses:
        all_verses.append((0, f"[Bibeltext: {reference}]"))

    return all_verses


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    for ref in ["Lukas 9, 57-62", "Epheser 4,23", "Lukas 18,35-19,10"]:
        print(f"\nReferenz: {ref}")
        for num, text in fetch_bible_text(ref):
            print(f"  {num}: {text[:70]}")

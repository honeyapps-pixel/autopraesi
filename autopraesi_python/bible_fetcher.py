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


def _lookup_book(book_de: str):
    """Schlägt ein Buch im Mapping nach, normalisiert '1 petrus' → '1. petrus' etc."""
    book_abbr = BOOK_MAPPING.get(book_de)
    if book_abbr:
        return book_abbr
    # Versuch mit Punkt nach Ziffer: "1 petrus" → "1. petrus"
    normalized = re.sub(r'^(\d)\s+', r'\1. ', book_de)
    book_abbr = BOOK_MAPPING.get(normalized)
    if book_abbr:
        return book_abbr
    # Fuzzy: Prefix-Match
    for key, val in BOOK_MAPPING.items():
        if book_de.startswith(key) or key.startswith(book_de):
            return val
    return None


def parse_reference(reference: str) -> list:
    """Parst eine deutsche Bibelstellen-Referenz.

    Beispiele:
        "Lukas 9, 57-62" → [("LUK", 9, 57, 62)]
        "Epheser 4,23" → [("EPH", 4, 23, 23)]

    Returns:
        Liste von (book_abbr, chapter, verse_start, verse_end) Tupeln
    """
    results = []
    parts = reference.split(";")

    for part in parts:
        part = part.strip()
        if not part:
            continue

        part = re.sub(r'\([^)]*\)', '', part).strip()

        match = re.match(
            r'^((?:\d\.?\s*)?[A-Za-zäöüÄÖÜß]+)\s+'
            r'(\d+)'
            r'[,.\s]\s*'
            r'(\d+)'
            r'(?:\s*[–\-]\s*(\d+))?',
            part
        )

        if match:
            book_de = match.group(1).strip().lower()
            chapter = int(match.group(2))
            verse_start = int(match.group(3))
            verse_end = int(match.group(4)) if match.group(4) else verse_start

            book_abbr = _lookup_book(book_de)

            if book_abbr:
                results.append((book_abbr, chapter, verse_start, verse_end))
            else:
                log.warning(f"Unbekanntes Buch: {book_de} in '{part}'")
            continue

        # Fallback: "Buch Kapitel" ohne Vers (z.B. "Psalm 133")
        match_chapter = re.match(
            r'^((?:\d\.?\s*)?[A-Za-zäöüÄÖÜß]+)\s+(\d+)\s*$',
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
            else:
                log.info(f"Lade Bibeltext: {book_abbr} {chapter},{verse_start}-{verse_end}")

            resp = requests.get(url, timeout=15,
                                headers={"User-Agent": "Mozilla/5.0"})
            if resp.status_code != 200:
                log.warning(f"HTTP {resp.status_code} für {url}")
                all_verses.append((0, f"[{reference}]"))
                continue

            soup = BeautifulSoup(resp.text, 'html.parser')

            if verse_start is None:
                # Ganzes Kapitel: alle Verse mit Klasse LU84.{book}.{chapter}.N sammeln
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
                    all_verses.append((v_num, verse_dict[v_num]))
            else:
                for v_num in range(verse_start, verse_end + 1):
                    class_name = f"LU84.{book_abbr}.{chapter}.{v_num}"
                    spans = soup.find_all(class_=class_name)

                    verse_text = ""
                    for span in spans:
                        text = span.get_text(strip=False)
                        text = re.sub(r'\s+', ' ', text).strip()
                        if text:
                            verse_text += " " + text if verse_text else text

                    verse_text = verse_text.strip()
                    if verse_text:
                        all_verses.append((v_num, verse_text))

        except requests.RequestException as e:
            log.warning(f"Netzwerkfehler beim Laden von {reference}: {e}")
            all_verses.append((0, f"[{reference}]"))

    if not all_verses:
        all_verses.append((0, f"[Bibeltext: {reference}]"))

    return all_verses


def format_verses_plain(verses: list) -> str:
    """Formatiert Verse als einfachen Text (für Fallback)."""
    parts = []
    for v_num, text in verses:
        if v_num > 0:
            parts.append(f"{v_num}{text}")
        else:
            parts.append(text)
    return "\n".join(parts)


def split_text_for_slides(text: str, max_chars: int = 500) -> list:
    """Teilt einen langen Text in Abschnitte für mehrere Folien."""
    if len(text) <= max_chars:
        return [text]

    lines = text.split("\n")
    chunks = []
    current = []
    current_len = 0

    for line in lines:
        if current_len + len(line) + 1 > max_chars and current:
            chunks.append("\n".join(current))
            current = [line]
            current_len = len(line)
        else:
            current.append(line)
            current_len += len(line) + 1

    if current:
        chunks.append("\n".join(current))

    return chunks


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)

    ref = "Lukas 9, 57-62"
    print(f"Referenz: {ref}")
    verses = fetch_bible_text(ref)
    for num, text in verses:
        print(f"  {num}: {text}")

    print(f"\nAls Plaintext:\n{format_verses_plain(verses)}")

    print("\n" + "=" * 60)
    ref2 = "Epheser 4,23"
    print(f"Referenz: {ref2}")
    verses2 = fetch_bible_text(ref2)
    for num, text in verses2:
        print(f"  {num}: {text}")

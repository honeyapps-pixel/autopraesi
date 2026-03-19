"""Sucht Lieder in der Brüderrecords-Bibliothek."""
from __future__ import annotations

import logging
import os
import re

from config import SONG_LIBRARY, SONG_DIRS

log = logging.getLogger(__name__)


def build_song_index(library_path: str = SONG_LIBRARY) -> dict:
    """Indexiert alle PPTX-Dateien in der Lied-Bibliothek.

    Returns:
        Dict mit {unterordner_name: [(dateiname, voller_pfad), ...]}
    """
    index = {}
    if not os.path.exists(library_path):
        log.error(f"Lied-Bibliothek nicht gefunden: {library_path}")
        return index

    for entry in os.listdir(library_path):
        subdir = os.path.join(library_path, entry)
        if not os.path.isdir(subdir):
            continue

        files = []
        for f in os.listdir(subdir):
            if f.lower().endswith(".pptx") and not f.startswith("~$"):
                files.append((f, os.path.join(subdir, f)))

        index[entry] = files
        log.debug(f"Index: {entry} → {len(files)} Dateien")

    total = sum(len(v) for v in index.values())
    log.info(f"Song-Index erstellt: {total} Dateien in {len(index)} Ordnern")
    return index


def _normalize(text: str) -> str:
    """Normalisiert Text für Vergleich."""
    return re.sub(r'\s+', ' ', text.strip().lower())


def find_song(song, index: dict) -> str | None:
    """Sucht ein Lied in der Bibliothek.

    Args:
        song: SongEntry-Objekt
        index: Song-Index aus build_song_index()

    Returns:
        Voller Pfad zur PPTX-Datei oder None
    """
    if not song.raw:
        return None

    # 1. Bestimme primären Suchordner
    primary_dirs = _get_search_dirs(song)

    # 2. Suche in primären Ordnern
    for dir_name in primary_dirs:
        if dir_name not in index:
            continue
        result = _search_in_dir(song, index[dir_name])
        if result:
            log.info(f"Lied gefunden: {song.raw} → {os.path.basename(result)}")
            return result

    # 3. Fallback: Alle Ordner durchsuchen
    log.debug(f"Lied nicht in primären Ordnern, durchsuche alle: {song.raw}")
    for dir_name, files in index.items():
        if dir_name in primary_dirs:
            continue
        result = _search_in_dir(song, files)
        if result:
            log.info(f"Lied gefunden (Fallback): {song.raw} → {os.path.basename(result)}")
            return result

    log.warning(f"Lied NICHT gefunden: {song.raw}")
    return None


def _get_search_dirs(song) -> list:
    """Bestimmt die Suchordner basierend auf Kategorie und Buch."""
    dirs = []

    # Lobpreisstrophen haben eigenen Ordner
    if song.category == "Lobpreisstrophe":
        dirs.append("Lobpreisstrophen")

    # Kinderlieder
    if song.category == "Kinderlied":
        dirs.append("Kinderlieder")
        return dirs

    # Sonstige Lieder
    if song.category == "Sonstige Lieder":
        dirs.append("Sonstige Lieder")
        return dirs

    # Buch-basierte Zuordnung
    if song.book:
        book_upper = song.book.upper().strip()
        # "FJ" ohne Nummer → könnte FJ1-FJ6 sein
        if book_upper == "FJ":
            dirs.extend(["FJ1", "FJ2", "FJ3", "FJ4", "FJ5", "FJ6"])
        elif book_upper in SONG_DIRS:
            dirs.append(SONG_DIRS[book_upper])
        else:
            # Versuche direkten Ordnernamen
            for dir_name in SONG_DIRS.values():
                if dir_name.upper() == book_upper:
                    dirs.append(dir_name)

    return dirs


def _search_in_dir(song, files: list) -> str | None:
    """Sucht ein Lied in einer Liste von Dateien."""

    # Strategie 1: Nach Nummer suchen (nummerierte Lieder)
    if song.number:
        for fname, fpath in files:
            name_lower = fname.lower()
            num = song.number

            # Lobpreisstrophen: Dateiname enthält Buch + Nummer
            # z.B. "FJ1 239 - Immer mehr.pptx"
            if song.category == "Lobpreisstrophe" and song.book:
                prefix1 = f"{song.book} {num} - ".lower()
                prefix2 = f"{song.book} {num}- ".lower()
                prefix3 = f"{song.book} {num} -".lower()
                if (name_lower.startswith(prefix1) or
                    name_lower.startswith(prefix2) or
                    name_lower.startswith(prefix3)):
                    return fpath

            # Standard: Dateiname beginnt mit Nummer
            # z.B. "235 - Jesus dir nach.pptx" oder "235- Jesus.pptx"
            prefix1 = f"{num} - "
            prefix2 = f"{num}- "
            prefix3 = f"{num} -"
            prefix4 = f"{num}-"
            if (name_lower.startswith(prefix1.lower()) or
                name_lower.startswith(prefix2.lower()) or
                name_lower.startswith(prefix3.lower()) or
                (name_lower.startswith(prefix4.lower()) and
                 len(name_lower) > len(prefix4) and
                 not name_lower[len(prefix4)].isdigit())):
                return fpath

    # Strategie 2: Nach Titelwörtern suchen (Kinderlieder, Sonstige)
    if song.title_words:
        search_words = [w.lower().rstrip(",.:;!?") for w in song.title_words]
        for fname, fpath in files:
            fname_lower = _normalize(os.path.splitext(fname)[0])
            # Prüfe ob alle Suchworte im Dateinamen vorkommen
            if all(w in fname_lower for w in search_words):
                return fpath

    return None


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    from datetime import date
    from excel_reader import read_godi_plan

    data = read_godi_plan(date(2026, 3, 8))
    if data:
        index = build_song_index()
        print(f"\nSuche {len(data.songs)} Lieder:")
        for song in data.songs:
            path = find_song(song, index)
            if path:
                print(f"  ✓ {song.slot_key}: {os.path.basename(path)}")
            else:
                print(f"  ✗ {song.slot_key}: NICHT GEFUNDEN ({song.raw})")

"""Sucht Lieder in der Brüderrecords-Bibliothek."""
from __future__ import annotations

import logging
import os
import re
import unicodedata
from typing import Optional

import storage
from config import SONG_LIBRARY, SONG_DIRS

log = logging.getLogger(__name__)


def build_song_index(library_path: str = SONG_LIBRARY) -> dict:
    """Indexiert alle PPTX-Dateien in der Lied-Bibliothek (in Dropbox).

    Returns:
        Dict mit {unterordner_name: [(dateiname, dropbox_pfad), ...]}
    """
    from dropbox.files import FolderMetadata

    index = {}
    try:
        entries = storage.list_folder(library_path)
    except Exception as e:
        log.error(f"Lied-Bibliothek nicht lesbar ({library_path}): {e}")
        return index

    for entry in entries:
        if not isinstance(entry, FolderMetadata):
            continue
        files = storage.list_files(entry.path_lower, suffix=".pptx")
        index[entry.name] = files
        log.debug(f"Index: {entry.name} → {len(files)} Dateien")

    total = sum(len(v) for v in index.values())
    log.info(f"Song-Index erstellt: {total} Dateien in {len(index)} Ordnern")
    return index


def _normalize_book(book: str) -> str:
    """Normalisiert eine Buch-Angabe für den Ordner-Abgleich.

    Im GoDi-Plan wird das Buch mal mit, mal ohne Leerzeichen geschrieben
    ("FJ3", "FJ 3", "fj 3"). Alle Schreibweisen sollen denselben Ordner
    treffen, daher: Großschreibung + alle Leerzeichen entfernen → "FJ3".
    """
    return re.sub(r'\s+', '', (book or '').upper())


# Wörter (Buchstaben/Ziffern beliebiger Sprache, kein Unterstrich)
_WORD_RE = re.compile(r'[^\W_]+', re.UNICODE)


def _tokens(text: str) -> list:
    """Zerlegt Text in kleingeschriebene Wort-Tokens (ohne Satzzeichen).

    Umlaute werden auf NFC vereinheitlicht: Dateinamen aus Dropbox/macOS
    kommen teils zerlegt (NFD, "u" + Trema), Excel-Text ist zusammengesetzt –
    sonst würde "Zünde" nie "Zünde" treffen.
    """
    return _WORD_RE.findall(unicodedata.normalize("NFC", text).lower())


def _title_score(song, fname: str) -> tuple:
    """Vergleicht den Excel-Titel mit dem Dateinamen.

    Ganze Wörter werden verglichen (kein Substring-Match), damit z.B. "ich"
    nicht in "nicht" gefunden wird. Ab 4 Zeichen reicht ein Wortanfang
    ("herz" trifft "herzen"), damit leichte Abweichungen toleriert werden.

    Returns:
        (matched3, total3, matched_all, extra)
        matched3/total3: Treffer unter den ersten 3 Titelwörtern (tolerant,
            entscheidet ob eine Datei überhaupt passt)
        matched_all: Treffer über den kompletten Titel (unterscheidet z.B.
            "Lord you have my heart" von "… - deutsch")
        extra: Dateinamen-Wörter ohne Entsprechung im Titel (je weniger,
            desto genauer passt die Datei)
    """
    words3 = [t for w in song.title_words for t in _tokens(w)]
    if not words3:
        return (0, 0, 0, 0)
    words_all = _tokens(song.title) or words3
    ftoks = _tokens(os.path.splitext(fname)[0])

    def hit(w):
        return any(f == w or (len(w) >= 4 and f.startswith(w)) for f in ftoks)

    matched3 = sum(1 for w in words3 if hit(w))
    matched_all = sum(1 for w in words_all if hit(w))
    extra = sum(1 for f in ftoks
                if not any(f == w or (len(w) >= 4 and f.startswith(w)) for w in words_all))
    return (matched3, len(words3), matched_all, extra)


def _title_matches_fully(song, fname: str) -> bool:
    """True, wenn alle (ersten 3) Titelwörter im Dateinamen vorkommen."""
    matched, total, _, _ = _title_score(song, fname)
    return total > 0 and matched == total


def _number_regex(song) -> "re.Pattern":
    """Regex, der prüft, ob ein Dateiname mit der Liednummer beginnt.

    Deckt alle in der Bibliothek vorkommenden Schreibweisen ab:
    "235 - Titel", "235- Titel", "235 -Titel", "235-Titel".
    "12-3 …" zählt nicht als Nummer 12 (Ziffer direkt nach dem Bindestrich).
    Lobpreisstrophen sind mit Buch benannt ("FJ1 239 - Immer mehr"); dort
    darf zwischen Buch und Ziffer ein Leerzeichen stehen oder nicht. Das
    Buch-Präfix ist optional, damit ersatzweise auch das ganze Lied aus dem
    Buch-Ordner ("239 - Immer mehr") gefunden wird.
    """
    num = re.escape(song.number)
    if song.category == "Lobpreisstrophe" and song.book:
        book = _normalize_book(song.book)
        book_pat = r'\s*'.join(re.escape(c) for c in book)
        return re.compile(rf'^(?:{book_pat}\s+)?{num}\s*[-–—](?!\d)', re.IGNORECASE)
    return re.compile(rf'^{num}\s*[-–—](?!\d)')


class SongMatch:
    """Ergebnis einer Liedsuche inkl. Hinweis, falls der Treffer unsicher ist."""

    __slots__ = ("path", "warning")

    def __init__(self, path: Optional[str] = None, warning: str = ""):
        self.path = path
        self.warning = warning

    @property
    def file_name(self) -> str:
        return os.path.basename(self.path) if self.path else ""


def find_song(song, index: dict) -> Optional[str]:
    """Sucht ein Lied in der Bibliothek.

    Args:
        song: SongEntry-Objekt
        index: Song-Index aus build_song_index()

    Returns:
        Voller Pfad zur PPTX-Datei oder None
    """
    return find_song_detailed(song, index).path


def find_song_detailed(song, index: dict) -> SongMatch:
    """Sucht ein Lied und liefert zusätzlich einen Prüf-Hinweis.

    Ablauf:
      1. Nummern-Treffer sammeln. In den primären Ordnern (= Buch/Kategorie
         aus dem Excel) ist die Nummer die verlässliche Kennung; in allen
         anderen Ordnern zählt ein Nummern-Treffer nur, wenn auch der
         komplette Titel passt (jedes Liederbuch hat eine Nr. 87 – ohne
         Titelabgleich wäre der Fallback Zufall).
      2. Kandidaten nach Titel-Übereinstimmung sortieren, bei Gleichstand
         gewinnt der primäre Ordner (in der Reihenfolge von _get_search_dirs,
         z.B. Lobpreisstrophen vor FJ1), danach der vollständigere Titel-
         Treffer und der Dateiname mit den wenigsten überzähligen Wörtern.
         So wird z.B. bei "87 - Lord you have my heart.pptx" und
         "87 - … - deutsch.pptx" die zum Excel-Titel passende Datei gewählt.
      3. Ohne brauchbaren Nummern-Treffer: reine Titelsuche (Kinderlieder,
         Sonstige Lieder, …), erst primäre Ordner, dann alle.

    Ein Hinweis (warning) wird gesetzt, wenn der Treffer nicht eindeutig zum
    Excel-Eintrag passt – damit er im Frontend sichtbar geprüft werden kann.
    """
    if not song.raw:
        return SongMatch()

    primary_dirs = _get_search_dirs(song)
    primary_rank = {d: i for i, d in enumerate(primary_dirs)}
    # Reihenfolge: primäre Ordner zuerst, dann der Rest (deterministisch)
    ordered_dirs = [d for d in primary_dirs if d in index]
    ordered_dirs += sorted(d for d in index if d not in primary_rank)

    # --- Strategie 1: Nummer ---
    if song.number:
        num_re = _number_regex(song)
        candidates = []  # (sort_key, dir_name, fname, fpath, matched3, total3)
        for dir_name in ordered_dirs:
            is_primary = dir_name in primary_rank
            for fname, fpath in sorted(index[dir_name]):
                if not num_re.match(unicodedata.normalize("NFC", fname)):
                    continue
                matched3, total3, matched_all, extra = _title_score(song, fname)
                if not is_primary and not (total3 > 0 and matched3 == total3):
                    continue  # Fremder Ordner: nur mit vollem Titel-Treffer
                rank = primary_rank.get(dir_name, len(primary_rank))
                # Rangfolge: Titel passt → primärer Ordner → Ordner-Priorität
                # (z.B. Lobpreisstrophen vor FJ1) → vollständigerer Titel →
                # wenigste überzählige Wörter → Dateiname
                key = (-matched3, 0 if is_primary else 1, rank, -matched_all,
                       extra, fname.lower())
                candidates.append((key, dir_name, fname, fpath, matched3, total3))

        if candidates:
            candidates.sort(key=lambda c: c[0])
            _, dir_name, fname, fpath, matched3, total3 = candidates[0]
            warning = ""
            if dir_name not in primary_rank:
                warning = f"In anderem Ordner gefunden ({dir_name}) – bitte prüfen"
            elif total3 and matched3 == 0:
                # Kein einziges Titelwort passt – z.B. andere Sprachfassung
                warning = "Titel passt nicht zum Dateinamen – bitte prüfen"
            else:
                # Gleich gute Treffer (z.B. zwei Sprachfassungen mit gleichem
                # Titel)? Dann den gewählten Treffer sichtbar machen.
                same = [c for c in candidates[1:] if c[0][:4] == candidates[0][0][:4]]
                if same:
                    others = ", ".join(c[2] for c in same[:3])
                    warning = f"Mehrere passende Dateien (auch: {others})"
            if warning:
                log.warning(f"Lied {song.raw!r} → {fname}: {warning}")
            else:
                log.info(f"Lied gefunden: {song.raw} → {fname}")
            return SongMatch(fpath, warning)

    # --- Strategie 2: Titelwörter (alle müssen passen) ---
    if song.title_words:
        for dir_name in ordered_dirs:
            for fname, fpath in sorted(index[dir_name]):
                if _title_matches_fully(song, fname):
                    if dir_name in primary_rank:
                        log.info(f"Lied gefunden: {song.raw} → {fname}")
                        return SongMatch(fpath)
                    warning = f"In anderem Ordner gefunden ({dir_name}) – bitte prüfen"
                    log.warning(f"Lied {song.raw!r} → {fname}: {warning}")
                    return SongMatch(fpath, warning)

    log.warning(f"Lied NICHT gefunden: {song.raw}")
    return SongMatch()


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

    # Buch-basierte Zuordnung ("FJ 3" und "FJ3" treffen denselben Ordner)
    if song.book:
        book_key = _normalize_book(song.book)
        # "FJ" ohne Nummer → könnte FJ1-FJ6 sein
        if book_key == "FJ":
            dirs.extend(["FJ1", "FJ2", "FJ3", "FJ4", "FJ5", "FJ6"])
        else:
            for key, dir_name in SONG_DIRS.items():
                if book_key in (_normalize_book(key), _normalize_book(dir_name)):
                    if dir_name not in dirs:
                        dirs.append(dir_name)
                    break

    return dirs


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

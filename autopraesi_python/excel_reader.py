"""Liest den GoDi-Plan aus der Excel-Datei (Zugriff über die Dropbox-API)."""
from __future__ import annotations

import io
import logging
import re
from dataclasses import dataclass, field
from datetime import date, datetime, time
from typing import Optional

import openpyxl

import storage
from config import (GODI_PLAN_DIR, EXCEL_ROWS, ABKUENDIGUNGEN_ROWS,
                     EINLADUNG_ROW_START, EINLADUNG_ROW_END, EINLADUNG_COLS)

log = logging.getLogger(__name__)

# Farb-Codes für die automatische Erkennung. Verglichen wird nur der 6-stellige RGB-Teil
# (ohne Alpha), und Zellfarben werden – egal ob als rgb ODER indexed gespeichert – auf RGB
# aufgelöst. So werden Lieder auch erkannt, wenn die Excel-Zelle eine indizierte Farbe
# nutzt (z.B. Index 42 = CCFFCC) statt der Theme-Farbe C5E0B4.
SONG_COLORS    = {"C5E0B4", "CCFFCC", "C6EFCE", "99CC00"}  # Grüntöne → Lied (99CC00 = Sa-Layout)
PREDIGT_COLORS = {"FFA7A7", "FFC7CE"}            # Rottöne   → Predigt-Bibelstelle
LESUNG_COLORS  = {"F8CBAD", "FCE4D6"}            # Orangetöne → Lesung-Bibelstelle

# Rückwärtskompatible Einzelwerte (falls anderswo referenziert)
SONG_COLOR    = "FFC5E0B4"
PREDIGT_COLOR = "FFFFA7A7"
LESUNG_COLOR  = "FFF8CBAD"

# Bekannte Lied-Präfixe für die Inhaltserkennung
_SONG_BOOKS = re.compile(
    r'^(FJ\d*|GLS|SUG|IWDD|Loben|SGIDH)\s+\d+',
    re.IGNORECASE
)
_SONG_PREFIXES = ("lobpreisstrophe:", "kinderlied:", "sonstige lieder")

# Muster für eine einzelne Bibelstelle (kein Semikolon = keine Liste)
# ("Joel 3,1-2", "5. Mose 8,7", "2 Samuel 7, 17-26")
_BIBLE_REF = re.compile(
    r'^((?:\d\.?\s*)?[A-Za-zäöüÄÖÜß]+)\s+\d+',
)

# Spalte-B-Beschriftungen, deren Inhalt ein Gemeindelied ist – auch ohne grüne Zellfarbe
# (z.B. "Sonstige Lieder - …" mit Theme-Farbe). "Liedbeitrag" (Vortrag) zählt bewusst nicht.
_SONG_LABELS = ("gemeindelied", "lobpreisstrophe", "kinderlied", "schlusslied", "liedstrophe")

# Datumszelle in den Abkündigungen: "Di 13.10.26", "So 11.10.2026"
_ANNOUNCEMENT_DATE = re.compile(r'^\s*[A-Za-zÄÖÜäöü]{2,3}\.?\s+\d{1,2}\.\d{1,2}\.')

# Beginn eines Zitats (Begrüßungsvers steht in Anführungszeichen)
_QUOTE_CHARS = ('"', '„', '“', '”', '‚', "'", '»', '«')


@dataclass
class InvitationEvent:
    """Ein Termin für die 'Herzliche Einladung' Folie."""
    date_str: str = ""       # "Di 17.03.26"
    time_str: str = ""       # "19:00"
    event_name: str = ""     # "Gebetsstunde"
    note: str = ""           # "Abend-GoDi fällt aus!" (optional, aus Spalte F)


@dataclass
class SongEntry:
    """Ein Lied aus dem GoDi-Plan."""
    raw: str                    # Original-Zelltext
    category: str = ""          # "Gemeindelied", "Lobpreisstrophe", "Kinderlied", etc.
    book: str = ""              # "FJ1", "SUG", "GLS", etc.
    number: str = ""            # "235", "61", etc.
    title: str = ""             # Volltitel nach " - "
    title_words: list = field(default_factory=list)  # Erste 3 Wörter des Titels
    slot_key: str = ""          # "song1" bis "song7"


@dataclass
class GodiPlanData:
    """Alle Daten eines Gottesdienstes aus dem GoDi-Plan."""
    service_header: str = ""        # "Gottesdienst am 08.03.2026 (Okuli)"
    theme: str = ""                 # "Der Blick nach vorn"
    date_str: str = ""              # "08.03.2026"
    kirchenkalender: str = ""       # "Okuli"
    greeting_verse: str = ""        # Begrüßungsvers
    lesung_reference: str = ""      # "Lukas 9, 57-62"
    predigt1_reference: str = ""    # "Lukas 9, 57-62"
    predigt1_title: str = ""        # "Nicht zurückschauen – ..."
    predigt2_reference: str = ""    # "Epheser 4,23"
    predigt2_title: str = ""        # "Wie das Evangelium ..."
    is_abendmahl: bool = False
    songs: list = field(default_factory=list)  # 7 SongEntry-Objekte
    announcements: list = field(default_factory=list)  # Abkündigungen
    invitation_events: list = field(default_factory=list)  # List[InvitationEvent]
    layout: str = "sonntag"         # "sonntag" (Inhalte in Spalte D) | "samstag" (Spalte C)


def _norm_rgb6(fg) -> Optional[str]:
    """Löst eine Zellfarbe auf den 6-stelligen RGB-Teil (großgeschrieben) auf.

    Unterstützt sowohl ``rgb``-Farben als auch ``indexed``-Palettenfarben (z.B. Index 42
    = CCFFCC). Theme-Farben werden hier nicht aufgelöst (kein zuverlässiger Tint-Bezug);
    sie liefern None.
    """
    try:
        if fg is None:
            return None
        if fg.type == 'rgb' and isinstance(fg.rgb, str):
            return fg.rgb.upper()[-6:]
        if fg.type == 'indexed':
            from openpyxl.styles.colors import COLOR_INDEX
            return str(COLOR_INDEX[fg.indexed]).upper()[-6:]
    except Exception:
        return None
    return None


def _get_cell_color(cell) -> Optional[str]:
    """Gibt den 6-stelligen RGB-Wert der Zellfüllung zurück, oder None.

    Ignoriert leere/weiße/schwarze Füllungen. Löst rgb- UND indexed-Farben auf, damit die
    Lied-/Bibelstellen-Erkennung unabhängig davon funktioniert, wie die Farbe gespeichert ist.
    """
    try:
        fill = cell.fill
        if not fill:
            return None
        rgb6 = _norm_rgb6(fill.fgColor)
        if not rgb6 or rgb6 in ("000000", "FFFFFF"):
            return None
        return rgb6
    except Exception:
        return None


def _looks_like_song(val: str) -> bool:
    """Prüft ob ein Zellinhalt wie ein Lied aussieht."""
    if not val:
        return False
    low = val.lower()
    if any(low.startswith(p) for p in _SONG_PREFIXES):
        return True
    if " - " in val and _SONG_BOOKS.match(val):
        return True
    return False


def _looks_like_single_bible_ref(val: str) -> bool:
    """Prüft ob ein Zellinhalt eine einzelne Bibelstelle ist (kein Semikolon)."""
    if not val or ";" in val:
        return False
    return bool(_BIBLE_REF.match(val.strip()))


def _scan_by_color(ws, col: int = 4) -> tuple:
    """Scannt das Arbeitsblatt nach Farben und extrahiert Lieder und Bibelstellen.

    ``col`` ist die Inhaltsspalte (Sonntags-Layout: D, Samstags-Layout: C). Ein Lied wird
    auch ohne grüne Farbe erkannt, wenn Spalte B es als Lied beschriftet (z.B.
    "Gemeindelied") und der Text wie ein Lied aussieht.

    Returns:
        (song_raws, lesung_ref, predigt_refs)
        - song_raws: Liste von (row, raw_text, col_b) für alle grünen Lied-Zellen
        - lesung_ref: Erster orangefarbener Bibelstellen-Eintrag
        - predigt_refs: Liste roter Bibelstellen-Einträge (erste=Predigt1, zweite=Predigt2)
    """
    song_raws = []
    lesung_ref = ""
    predigt_refs = []

    for row in range(1, ws.max_row + 1):
        cell = ws.cell(row=row, column=col)
        color = _get_cell_color(cell)
        val = _text(cell.value)
        col_b = _text(ws.cell(row=row, column=2).value)
        song_label = col_b.lower().startswith(_SONG_LABELS)

        if (color in SONG_COLORS or song_label) and _looks_like_song(val):
            song_raws.append((row, val, col_b))
            log.debug(f"Farb-Scan: Lied in Zeile {row}: {val[:60]} (B={col_b})")

        elif color in LESUNG_COLORS and _looks_like_single_bible_ref(val) and not lesung_ref:
            lesung_ref = val
            log.debug(f"Farb-Scan: Lesung in Zeile {row}: {val}")

        elif color in PREDIGT_COLORS and _looks_like_single_bible_ref(val):
            predigt_refs.append(val)
            log.debug(f"Farb-Scan: Predigt-Ref in Zeile {row}: {val}")

    return song_raws, lesung_ref, predigt_refs


def _text(val) -> str:
    """Zellwert als bereinigter Text; Uhrzeiten/Datumswerte und leere Zellen → ""."""
    if val is None or isinstance(val, (time, datetime)):
        return ""
    return str(val).strip()


def _strip_placeholder(val) -> str:
    """Zellwert ohne Vorlagen-Präfix: "Text: Psalm 25,5" → "Psalm 25,5", "Titel: " → ""."""
    txt = _text(val)
    return re.sub(r'^(text|titel)\s*:\s*', '', txt, flags=re.IGNORECASE).strip()


def _fmt_time(val) -> str:
    """Uhrzeit-Zelle → "HH:MM" (datetime.time oder Text "19:00:00")."""
    if hasattr(val, "strftime"):
        return val.strftime("%H:%M")
    txt = str(val).strip() if val is not None else ""
    if re.match(r'^\d{1,2}:\d{2}:\d{2}$', txt):
        return txt[:-3]
    return txt


def _find_label_row(ws, prefix: str, col: int = 1, max_row: int = 15,
                    exact: bool = False) -> Optional[int]:
    """Erste Zeile, deren Beschriftung (Spalte ``col``) mit ``prefix`` beginnt.

    ``exact``: Beschriftung muss (ohne ":") genau ``prefix`` sein – unterscheidet
    "Predigttext:" (Sa-Layout) von "Predigttextvorschläge:" (So-Layout).
    """
    for row in range(1, min(ws.max_row, max_row) + 1):
        label = _text(ws.cell(row=row, column=col).value).lower().rstrip(":").strip()
        if (label == prefix) if exact else label.startswith(prefix):
            return row
    return None


def _detect_content_col(ws) -> int:
    """Inhaltsspalte des Sheets: D (4) im Sonntags-Layout, C (3) im Samstags-Layout.

    Das Samstags-Sheet ist schmaler: Thema/Leitung/Lieder stehen in Spalte C statt D.
    """
    for prefix in ("thema", "godi-leitung", "verkündigung"):
        row = _find_label_row(ws, prefix)
        if row:
            if _text(ws.cell(row=row, column=4).value):
                return 4
            if _text(ws.cell(row=row, column=3).value):
                return 3
    # Noch unbefülltes Sheet: Sa-Layout hat "Predigttext:", So-Layout "Predigttextvorschläge:"
    return 3 if _find_label_row(ws, "predigttext", exact=True) else 4


def _find_announcement_anchor(ws) -> Optional[int]:
    """Zeile mit "Abkündigungen" in Spalte B (Beginn des Termin-Blocks)."""
    for row in range(1, ws.max_row + 1):
        if _text(ws.cell(row=row, column=2).value).lower().startswith("abkündigungen"):
            return row
    return None


def _find_greeting(ws, col: int, end_row: int) -> str:
    """Begrüßungsvers: Inhalt der Zeile "Begrüßung" oder erstes Zitat vor der Predigt.

    Die Zeile wandert zwischen den Quartalen (16, 17, 18 …) und vor ihr kann z.B. ein
    "Liedbeitrag" stehen – daher nicht mehr über eine feste Zeile.
    """
    for row in range(10, end_row):
        label = _text(ws.cell(row=row, column=2).value).lower()
        val = _text(ws.cell(row=row, column=col).value)
        if label.startswith("begrüßung") and val and not _looks_like_song(val):
            return val
    for row in range(10, end_row):
        val = _text(ws.cell(row=row, column=col).value)
        if val.startswith(_QUOTE_CHARS):
            return val
    return ""


def _read_announcement_block(ws, anchor: int) -> tuple[list, list]:
    """Liest Abkündigungen + Einladungs-Termine ab der "Abkündigungen"-Zeile.

    Die Termin-Spalten werden je Zeile über die Datumszelle erkannt (So-Layout: B, Sa-Layout:
    C, dort ggf. schon in der Abkündigungen-Zeile selbst). Daneben stehen Uhrzeit,
    Veranstaltung und (So-Layout) in Spalte F ein Hinweis. Eine Zeile ohne Datum, aber mit
    Uhrzeit (zweiter Termin am selben Tag) übernimmt das Datum der Vorzeile.
    Der Block endet bei "Ausblick", "Danke an", dem Schlusslied oder "Nach dem GoDi".
    """
    announcements: list[str] = []
    events: list[InvitationEvent] = []
    date_col = None
    last_date = ""
    stop = ("ausblick", "danke an", "nach dem godi") + _SONG_LABELS
    for row in range(anchor, min(ws.max_row, anchor + 14) + 1):
        col_a = _text(ws.cell(row=row, column=1).value).lower()
        col_b = _text(ws.cell(row=row, column=2).value).lower()
        if row > anchor and (col_b.startswith(stop) or col_a.startswith("nach dem godi")):
            break
        for c in (2, 3):
            if _ANNOUNCEMENT_DATE.match(_text(ws.cell(row=row, column=c).value)):
                date_col = c
                break
        else:
            if date_col is None:
                continue
        datum = _text(ws.cell(row=row, column=date_col).value)
        if not _ANNOUNCEMENT_DATE.match(datum):
            datum = ""
        zeit = _fmt_time(ws.cell(row=row, column=date_col + 1).value)
        event = _text(ws.cell(row=row, column=date_col + 2).value)
        note = ""
        if date_col + 4 <= 6:  # Hinweis-Spalte F nur im So-Layout (Sa: dort Kürzel)
            note = _text(ws.cell(row=row, column=date_col + 4).value)
        if not event or not (datum or zeit):
            continue
        if datum:
            last_date = datum
        elif not last_date:
            continue
        announcements.append(" | ".join(p for p in (datum or last_date, zeit, event, note) if p))
        events.append(InvitationEvent(date_str=datum or last_date, time_str=zeit,
                                      event_name=event, note=note))
    return announcements, events


def parse_song_entry(raw: str, slot_key: str = "", col_b: str = "") -> SongEntry:
    """Parst einen Lied-Eintrag aus der Excel-Zelle.

    Args:
        raw: Zellinhalt aus Spalte D
        slot_key: Template-Slot (z.B. "song1") – wird ggf. später überschrieben
        col_b: Zellinhalt aus Spalte B für zusätzliche Kategorie-Erkennung
    """
    if not raw or not raw.strip():
        return SongEntry(raw="", slot_key=slot_key)

    song = SongEntry(raw=raw.strip(), slot_key=slot_key)

    text = raw.strip()
    # Doppelte Leerzeichen normalisieren (v1.4.1 Bugfix)
    text = re.sub(r'\s+', ' ', text)

    # Kategorie-Prefix erkennen und entfernen. Das Trennzeichen nach dem Präfix
    # variiert in der Praxis ("Sonstige Lieder - X", "Sonstige Lieder- X",
    # "Sonstige Lieder: X") – daher führende Trennzeichen robust abschneiden.
    _SEP = " -–—:\t"
    low = text.lower()
    if low.startswith("lobpreisstrophe"):
        song.category = "Lobpreisstrophe"
        text = text[len("lobpreisstrophe"):].lstrip(_SEP).strip()
    elif low.startswith("kinderlied"):
        song.category = "Kinderlied"
        text = text[len("kinderlied"):].lstrip(_SEP).strip()
    elif low.startswith("sonstige lieder"):
        song.category = "Sonstige Lieder"
        text = text[len("sonstige lieder"):].lstrip(_SEP).strip()
    else:
        song.category = "Gemeindelied"

    # Spalte B als zusätzliches Signal für Lobpreisstrophe
    # z.B. B="Lobpreisstrophe", D="FJ1 28 - In dir ist mein Leben"
    if song.category == "Gemeindelied" and col_b.lower().startswith("lobpreisstrophe"):
        song.category = "Lobpreisstrophe"

    # Buch und Nummer extrahieren: "FJ1 235 - Jesus, dir nach"
    # Oder nur Titel bei Kinderliedern: "Gottes große Liebe"
    if " - " in text:
        prefix, title = text.split(" - ", 1)
        song.title = title.strip()

        # Prefix parsen: "FJ1 235" oder "SUG 61" oder "FJ 29"
        parts = prefix.strip().split()
        if len(parts) >= 2:
            song.book = parts[0].strip()
            # Nummer: führendes "-" entfernen (v1.4.3 Bugfix)
            num = parts[-1].strip().lstrip("-").rstrip("-")
            # "FJ" ohne Nummer ist manchmal "FJ 29" → book=FJ, number=29
            # Prüfen ob der letzte Teil eine Nummer ist
            if num.replace(".", "").isdigit():
                song.number = num
                # Wenn mehr als 2 Teile, book könnte "FJ1" oder "SGIDH 2" sein
                if len(parts) > 2:
                    song.book = " ".join(parts[:-1])
            else:
                song.book = prefix.strip()
        elif len(parts) == 1:
            # Nur Buch, keine Nummer
            song.book = parts[0].strip()
    else:
        # Kein " - " → reiner Titel (Kinderlieder, etc.)
        song.title = text.strip()

    # Erste 3 Wörter des Titels für die Suche (v1.4.2)
    if song.title:
        words = song.title.split()
        song.title_words = words[:3]

    log.debug(f"Parsed song: {song}")
    return song


def _assign_songs_to_slots(song_raws: list) -> list:
    """Weist Lieder nach liturgischer Rolle den Template-Slots zu.

    Statt Lieder sequentiell song1–song7 zuzuweisen, werden sie nach
    Kategorie auf die richtigen Template-Positionen verteilt:
      - song3: Kinderlied (vor Kinderstunde)
      - song2, song6: Lobpreisstrophen
      - song1, song4, song5, song7: Gemeindelieder (song7 = Schlusslied)

    Args:
        song_raws: Liste von (row, raw_text, col_b) aus dem Farb-Scan
    """
    # Alle Lieder parsen (ohne feste Slot-Zuweisung)
    all_songs = []
    for row, raw, col_b in song_raws:
        song = parse_song_entry(raw, col_b=col_b)
        all_songs.append(song)

    # Nach Kategorie aufteilen
    kinderlieder = [s for s in all_songs if s.category == "Kinderlied"]
    lobpreis = [s for s in all_songs if s.category == "Lobpreisstrophe"]
    gemeinde = [s for s in all_songs
                if s.category in ("Gemeindelied", "Sonstige Lieder")]

    slots = {}

    # Kinderlied → song3
    if kinderlieder:
        slots["song3"] = kinderlieder[0]
        if len(kinderlieder) > 1:
            log.warning(f"{len(kinderlieder)} Kinderlieder gefunden, "
                        f"nur das erste wird verwendet")

    # Lobpreisstrophen → song2, song6
    if len(lobpreis) >= 1:
        slots["song2"] = lobpreis[0]
    if len(lobpreis) >= 2:
        slots["song6"] = lobpreis[1]
    if len(lobpreis) > 2:
        log.warning(f"{len(lobpreis)} Lobpreisstrophen gefunden, "
                    f"nur die ersten 2 werden verwendet")

    # Gemeindelieder → song1 (Eröffnung), song4, song5, song7 (Schlusslied)
    # Überzählige → song_extra1, song_extra2, ...
    if len(gemeinde) >= 2:
        # Erstes Gemeindelied → song1, letztes → song7 (Schlusslied)
        slots["song1"] = gemeinde[0]
        slots["song7"] = gemeinde[-1]
        # Mittlere Gemeindelieder → song4, song5, dann Extras
        middle = gemeinde[1:-1]
        middle_slots = ["song4", "song5"]
        for i, song in enumerate(middle):
            if i < len(middle_slots):
                slots[middle_slots[i]] = song
            else:
                extra_key = f"song_extra{i - len(middle_slots) + 1}"
                slots[extra_key] = song
                log.info(f"Extra-Lied: {song.raw[:40]} → {extra_key}")
    elif len(gemeinde) == 1:
        slots["song1"] = gemeinde[0]
    else:
        log.warning("Keine Gemeindelieder im Farb-Scan gefunden")

    # Ergebnis-Liste aufbauen
    all_slots = ["song1", "song2", "song3", "song4", "song5", "song6", "song7"]
    # Extra-Slots hinzufügen
    extra_slots = sorted([k for k in slots if k.startswith("song_extra")])
    all_slots.extend(extra_slots)

    result = []
    for slot_key in all_slots:
        if slot_key in slots:
            song = slots[slot_key]
            song.slot_key = slot_key
            result.append(song)
        else:
            result.append(SongEntry(raw="", slot_key=slot_key))
            log.warning(f"Slot {slot_key}: kein passendes Lied gefunden")

    # Zusammenfassung loggen
    assigned = [s for s in all_songs if any(s is slots.get(k) for k in all_slots)]
    skipped = [s for s in all_songs if s not in assigned]
    if skipped:
        log.warning(f"{len(skipped)} Lied(er) nicht zugewiesen: "
                    f"{[s.raw[:40] for s in skipped]}")

    return result


def _parse_header(header: str, excel_path: str = "") -> tuple:
    """Extrahiert Datum und Kirchenkalendername aus dem Header.
    'Gottesdienst am 08.03.2026 (Okuli)' → ('08.03.2026', 'Okuli')

    Fehlt im Header das Jahr ('Gottesdienst am Samstag 03.10.'), wird es aus dem
    Dateinamen des Plans ergänzt ('GoDi-Plan 2026_Q4.xlsx'; Januar/Februar in einem
    Q4-Plan gehören ins Folgejahr).
    """
    date_str = ""
    kirche = ""

    if not header:
        return date_str, kirche

    # Datum extrahieren: DD.MM.YYYY (auch einstellige Tage/Monate)
    m = re.search(r'(\d{1,2})\.(\d{1,2})\.(\d{4})?', header)
    if m:
        day, month, year = int(m.group(1)), int(m.group(2)), m.group(3)
        if not year:
            fm = re.search(r'(20\d{2})_(Q?\d)', excel_path)
            if fm:
                y = int(fm.group(1))
                if fm.group(2).upper().lstrip("Q") == "4" and month <= 2:
                    y += 1
                year = str(y)
        if year and 1 <= day <= 31 and 1 <= month <= 12:
            date_str = f"{day:02d}.{month:02d}.{year}"

    # Kirchenkalender extrahieren: (Name)
    m = re.search(r'\(([^)]+)\)', header)
    if m:
        kirche = m.group(1).strip()

    return date_str, kirche


def _is_conflict_copy(name: str) -> bool:
    """Erkennt Dropbox-Konfliktkopien (sollen ignoriert werden)."""
    low = name.lower()
    return "in konflikt stehende kopie" in low or "conflicted copy" in low


def _is_godi_plan(name: str) -> bool:
    """Prüft, ob ein Dateiname eine echte GoDi-Plan Excel-Datei ist.

    Dropbox-Konfliktkopien werden ausgeschlossen, damit sie weder im
    Präsentations- noch im GoDi-Plan-Reiter als Auswahl auftauchen.
    """
    if _is_conflict_copy(name):
        return False
    return name.startswith("GoDi-Plan") and name.lower().endswith(".xlsx")


def find_godi_plan_excel(sunday: date) -> Optional[str]:
    """Findet die richtige GoDi-Plan Excel-Datei im Dropbox-Ordner."""
    sheet_name = f"So {sunday.day}.{sunday.month}"

    for name, path in storage.list_files(GODI_PLAN_DIR, suffix=".xlsx"):
        if not _is_godi_plan(name):
            continue
        try:
            wb = openpyxl.load_workbook(
                io.BytesIO(storage.download_bytes(path)),
                read_only=True, data_only=True)
            # Prüfe ob Sheet existiert (auch mit trailing space)
            for sheet in wb.sheetnames:
                if sheet.strip() == sheet_name:
                    wb.close()
                    log.info(f"GoDi-Plan gefunden: {path}, Sheet: {sheet}")
                    return path
            wb.close()
        except Exception as e:
            log.warning(f"Fehler beim Öffnen von {path}: {e}")

    return None


def list_all_sheets() -> list[tuple[str, str]]:
    """Listet alle Sheets aus allen GoDi-Plan Excel-Dateien (in Dropbox).

    Returns:
        Liste von (sheet_name, excel_path) Tupeln, wobei excel_path ein
        Dropbox-Pfad ist. Hilfs-Sheets (Überblick, Vorlage, etc.) werden gefiltert.
    """
    skip = {"Überblick", "GoDi-Vorlage", "Nächsten GoDis im nächsten Plan"}
    results = []
    for name, path in storage.list_files(GODI_PLAN_DIR, suffix=".xlsx"):
        if not _is_godi_plan(name):
            continue
        try:
            wb = openpyxl.load_workbook(
                io.BytesIO(storage.download_bytes(path)),
                read_only=True, data_only=True)
            for sheet in wb.sheetnames:
                stripped = sheet.strip()
                if stripped and stripped not in skip:
                    results.append((stripped, path))
            wb.close()
        except Exception as e:
            log.warning(f"Fehler beim Öffnen von {path}: {e}")
    return results


def _read_fixed_announcements(ws, data: GodiPlanData) -> None:
    """Alter Weg für Sheets ohne "Abkündigungen"-Beschriftung: feste Zeilen aus config."""
    for row_num in ABKUENDIGUNGEN_ROWS:
        parts = [_fmt_time(v) for v in (ws.cell(row=row_num, column=c).value for c in range(2, 6)) if v]
        if parts:
            data.announcements.append(" | ".join(parts))
    for row_num in range(EINLADUNG_ROW_START, EINLADUNG_ROW_END + 1):
        datum_str = _text(ws.cell(row=row_num, column=EINLADUNG_COLS["datum"]).value)
        event_str = _text(ws.cell(row=row_num, column=EINLADUNG_COLS["veranstaltung"]).value)
        if datum_str and event_str:
            data.invitation_events.append(InvitationEvent(
                date_str=datum_str,
                time_str=_fmt_time(ws.cell(row=row_num, column=EINLADUNG_COLS["uhrzeit"]).value),
                event_name=event_str,
                note=_text(ws.cell(row=row_num, column=EINLADUNG_COLS["hinweis"]).value),
            ))


def read_godi_plan_by_sheet(sheet_name: str, excel_path: str) -> GodiPlanData | None:
    """Liest den GoDi-Plan für ein beliebiges Sheet.

    Args:
        sheet_name: Name des Sheets (z.B. 'So 15.3', 'Pa 18.3')
        excel_path: Dropbox-Pfad zur Excel-Datei
    """
    wb = openpyxl.load_workbook(
        io.BytesIO(storage.download_bytes(excel_path)), data_only=True)

    # Sheet finden (mit trailing space Fallback)
    ws = None
    for name in wb.sheetnames:
        if name.strip() == sheet_name:
            ws = wb[name]
            break

    if ws is None:
        log.error(f"Sheet '{sheet_name}' nicht gefunden in {excel_path}")
        wb.close()
        return None

    def cell(key: str) -> str:
        """Liest eine Zelle anhand des Schlüssels aus EXCEL_ROWS.

        Leere Vorlagen-Platzhalter ("Text:", "Titel:") zählen als leer, ebenso Zellen in
        einer Zeile, die Spalte B als "Abendmahl" ausweist (Q4: statt "Predigt 2").
        """
        row, col = EXCEL_ROWS[key]
        if _text(ws.cell(row=row, column=2).value).lower().startswith("abendmahl"):
            return ""
        return _strip_placeholder(ws.cell(row=row, column=col).value)

    data = GodiPlanData()

    # Layout erkennen: Sonntags-Sheets haben die Inhalte in Spalte D, die schmaleren
    # Samstags-Sheets in Spalte C. Zeilen verschieben sich zwischen den Quartalen,
    # daher wird überall zuerst über Beschriftungen/Farben gesucht und nur im
    # Sonntags-Layout (validiert) auf die festen Zeilen aus EXCEL_ROWS zurückgefallen.
    col = _detect_content_col(ws)
    compact = col == 3
    if compact:
        data.layout = "samstag"
        log.info("Samstags-Layout erkannt (Inhalte in Spalte C)")

    def label_value(prefix: str, *, exact: bool = False) -> str:
        row = _find_label_row(ws, prefix, exact=exact)
        return _text(ws.cell(row=row, column=col).value) if row else ""

    anchor = _find_announcement_anchor(ws)
    body_end = anchor or min(ws.max_row, 40)

    # Header und Metadaten
    data.service_header = cell("header")
    data.theme = label_value("thema") or ("" if compact else cell("theme"))
    data.date_str, data.kirchenkalender = _parse_header(data.service_header, excel_path)

    # Begrüßungsvers (Zeile wandert je nach Plan → über Beschriftung/Zitat suchen)
    data.greeting_verse = _find_greeting(ws, col, body_end)

    # === Farb-basierter Scan für Lieder und Bibelstellen ===
    song_raws, lesung_ref_color, predigt_refs_color = _scan_by_color(ws, col)

    # Bibelstellen: Farb-Scan bevorzugt, Fallback auf Beschriftung bzw. feste Zeilen
    if lesung_ref_color:
        data.lesung_reference = lesung_ref_color
        log.info(f"Lesung (Farbe erkannt): {lesung_ref_color}")
    else:
        lesung = label_value("lesungstext")
        data.lesung_reference = lesung if _looks_like_single_bible_ref(lesung) else ""
        if not data.lesung_reference:
            log.warning("Lesung: keine Bibelstelle gefunden")

    if compact:
        # Sa-Layout: eine Predigt, Bibelstelle in der Kopfzeile "Predigttext:"
        ref = label_value("predigttext", exact=True)
        data.predigt1_reference = ref if _looks_like_single_bible_ref(ref) else (
            predigt_refs_color[0] if predigt_refs_color else "")
    else:
        if len(predigt_refs_color) >= 1:
            data.predigt1_reference = predigt_refs_color[0]
            log.info(f"Predigt1 (Farbe erkannt): {predigt_refs_color[0]}")
        elif _looks_like_single_bible_ref(cell("predigt1_referenz")):
            data.predigt1_reference = cell("predigt1_referenz")
            log.warning("Predigt1: Farb-Erkennung fehlgeschlagen, nutze feste Zeile")

        if len(predigt_refs_color) >= 2:
            data.predigt2_reference = predigt_refs_color[1]
            log.info(f"Predigt2 (Farbe erkannt): {predigt_refs_color[1]}")
        elif _looks_like_single_bible_ref(cell("predigt2_referenz")):
            data.predigt2_reference = cell("predigt2_referenz")
            log.warning("Predigt2: Farb-Erkennung fehlgeschlagen, nutze feste Zeile")

    # Predigt-Titel und -Referenz dynamisch über Spalte B suchen
    # (feste Zeilen funktionieren nicht bei variablen Gottesdienst-Layouts)
    predigt1_found = False
    predigt2_found = False
    for row in range(1, ws.max_row + 1):
        col_b = ws.cell(row=row, column=2).value
        if not col_b or not isinstance(col_b, str):
            continue
        col_b_lower = col_b.strip().lower()

        if "predigt 1" in col_b_lower or "predigt1" in col_b_lower:
            ref = _strip_placeholder(ws.cell(row=row, column=4).value)
            title = _strip_placeholder(ws.cell(row=row, column=5).value)
            if ref:
                data.predigt1_reference = ref
            if title:
                data.predigt1_title = title
            predigt1_found = True
            log.info(f"Predigt1 (Spalte B Zeile {row}): ref=\"{data.predigt1_reference}\" titel=\"{data.predigt1_title}\"")

        elif compact and col_b_lower.startswith("predigt") and not predigt1_found:
            # Sa-Layout: Zeile "Predigt" trägt in Spalte C den Titel
            title = _text(ws.cell(row=row, column=col).value)
            if title:
                data.predigt1_title = title
            predigt1_found = True
            log.info(f"Predigt (Sa-Layout, Zeile {row}): titel=\"{data.predigt1_title}\"")

        elif "predigt 2" in col_b_lower or "predigt2" in col_b_lower:
            ref = _strip_placeholder(ws.cell(row=row, column=4).value)
            title = _strip_placeholder(ws.cell(row=row, column=5).value)
            if ref:
                data.predigt2_reference = ref
            if title:
                data.predigt2_title = title
            predigt2_found = True
            log.info(f"Predigt2 (Spalte B Zeile {row}): ref=\"{data.predigt2_reference}\" titel=\"{data.predigt2_title}\"")

    if not compact:
        # Fallback auf feste Zeilen (nur Sonntags-Layout)
        if not predigt1_found and not data.predigt1_title:
            data.predigt1_title = cell("predigt1_titel")
        if not predigt2_found and not data.predigt2_title:
            data.predigt2_title = cell("predigt2_titel")

    # Abendmahl-Erkennung: Beschriftung "Abendmahl" vor den Abkündigungen (Q4: Spalte B
    # statt "Predigt 2", oder als Verkündigungs-Teil in Spalte F) bzw. Thema.
    data.is_abendmahl = data.theme.lower().startswith("abendmahl") or any(
        _text(ws.cell(row=r, column=c).value).lower().startswith("abendmahl")
        for r in range(2, body_end) for c in range(2, 7))

    # === Lieder: Farb-Scan ===
    if len(song_raws) >= 2 or compact:
        log.info(f"Farb-Scan: {len(song_raws)} Lieder gefunden")
        data.songs = _assign_songs_to_slots(song_raws)
    else:
        # Fallback: feste Zeilen-Mappings – nur Zellen übernehmen, die wie ein Lied aussehen
        log.warning(f"Farb-Scan lieferte nur {len(song_raws)} Lieder – "
                    f"Fallback auf feste Zeilen-Mappings")
        song_keys = ["song1", "song2", "song3", "song4", "song5", "song6", "song7"]
        for key in song_keys:
            raw = cell(key)
            song = parse_song_entry(raw if _looks_like_song(raw) else "", key)
            data.songs.append(song)

    # Abkündigungen + Einladungs-Events (Herzliche Einladung Folie)
    if anchor:
        data.announcements, data.invitation_events = _read_announcement_block(ws, anchor)
    else:
        _read_fixed_announcements(ws, data)

    log.info(f"Einladungs-Events: {len(data.invitation_events)} gefunden")

    wb.close()

    log.info(f"GoDi-Plan gelesen: {data.service_header}")
    log.info(f"Thema: {data.theme}")
    log.info(f"Lieder: {[s.raw for s in data.songs]}")

    return data


def read_godi_plan(sunday: date) -> GodiPlanData | None:
    """Liest den GoDi-Plan für den gegebenen Sonntag (Wrapper für Automation)."""
    excel_path = find_godi_plan_excel(sunday)
    if not excel_path:
        log.error(f"Keine GoDi-Plan Datei gefunden für {sunday}")
        return None
    sheet_name = f"So {sunday.day}.{sunday.month}"
    return read_godi_plan_by_sheet(sheet_name, excel_path)


if __name__ == "__main__":
    logging.basicConfig(level=logging.DEBUG)
    # Test mit So 15.3.2026
    result = read_godi_plan(date(2026, 3, 15))
    if result:
        print(f"\nHeader: {result.service_header}")
        print(f"Thema: {result.theme}")
        print(f"Datum: {result.date_str}, Kirche: {result.kirchenkalender}")
        print(f"Lesung: {result.lesung_reference}")
        print(f"Predigt 1: {result.predigt1_reference} - {result.predigt1_title}")
        print(f"Predigt 2: {result.predigt2_reference} - {result.predigt2_title}")
        print(f"Abendmahl: {result.is_abendmahl}")
        print(f"\nLieder:")
        for s in result.songs:
            print(f"  {s.slot_key}: book={s.book}, num={s.number}, "
                  f"title={s.title}, cat={s.category}")

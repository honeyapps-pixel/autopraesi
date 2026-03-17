"""AutoPräsi – Automatische Gottesdienst-Präsentation.

Erstellt jeden Donnerstag die Präsentation für den kommenden Sonntag.
"""

import logging
import os
import sys
import time
from datetime import date, timedelta

from config import LOG_DIR, IMAGE_DIR
from excel_reader import read_godi_plan
from song_finder import build_song_index, find_song
from presentation_builder import build_presentation
from status_reporter import report_success, report_run

log = logging.getLogger("autopraesi")


def setup_logging():
    """Richtet Logging in Datei und Konsole ein."""
    os.makedirs(LOG_DIR, exist_ok=True)
    log_file = os.path.join(LOG_DIR, f"autopraesi_{date.today().isoformat()}.log")

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )


def next_sunday(from_date: date = None) -> date:
    """Berechnet den nächsten Sonntag (oder heute, falls Sonntag)."""
    if from_date is None:
        from_date = date.today()
    days_ahead = 6 - from_date.weekday()  # Montag=0, Sonntag=6
    if days_ahead < 0:
        days_ahead += 7
    return from_date + timedelta(days=days_ahead)


def _find_image(date_str: str):
    """Sucht das Hintergrundbild für den Gottesdienst (z.B. 'Bild 08.03..jpg')."""
    if not date_str:
        return None
    # date_str ist z.B. "08.03.2026" → wir brauchen "08.03."
    parts = date_str.split(".")
    if len(parts) < 2:
        return None
    name = f"Bild {parts[0]}.{parts[1]}..jpg"
    path = os.path.join(IMAGE_DIR, name)
    if os.path.exists(path):
        log.info(f"Hintergrundbild gefunden: {name}")
        return path
    log.warning(f"Hintergrundbild nicht gefunden: {path}")
    return None


def run(sunday: date = None):
    """Hauptablauf: Liest Plan, sucht Lieder, baut Präsentation."""
    if sunday is None:
        sunday = next_sunday()

    start_time = time.time()
    log.info(f"=== AutoPräsi für Sonntag {sunday.strftime('%d.%m.%Y')} ===")

    # 1. GoDi-Plan lesen
    data = read_godi_plan(sunday)
    if not data:
        log.error("GoDi-Plan nicht gefunden – Abbruch.")
        report_run(sunday, "error", error_message="GoDi-Plan nicht gefunden")
        return None

    log.info(f"Thema: {data.theme}")
    log.info(f"Datum: {data.date_str}, Kirchenkalender: {data.kirchenkalender}")

    # 2. Lied-Index aufbauen und Lieder suchen
    index = build_song_index()
    song_paths = {}
    missing = []
    PFLICHT_SLOTS = {"song1", "song4", "song5", "song7"}

    for song in data.songs:
        path = find_song(song, index)
        if path:
            song_paths[song.slot_key] = path
        elif song.raw:
            missing.append(f"{song.slot_key}: {song.raw}")

    # Song-Zusammenfassung als Tabelle
    log.info("--- Lied-Übersicht ---")
    for song in data.songs:
        slot = song.slot_key
        status = "OK" if slot in song_paths else ("FEHLT" if song.raw else "leer")
        datei = os.path.basename(song_paths[slot]) if slot in song_paths else "-"
        log.info(f"  {slot} [{status:5s}] {song.raw[:45] if song.raw else '(kein Eintrag)':<45} → {datei}")
        if status == "FEHLT" and slot in PFLICHT_SLOTS:
            log.warning(f"  *** PFLICHT-SLOT {slot} fehlt: {song.raw} ***")
    log.info("----------------------")

    if missing:
        log.warning(f"{len(missing)} Lied(er) nicht gefunden: {missing}")

    # 3. Hintergrundbild suchen
    image_path = _find_image(data.date_str)

    # 4. Präsentation bauen
    try:
        output = build_presentation(data, song_paths, image_path=image_path,
                                    fetch_bible=True)
    except Exception as e:
        log.error(f"Fehler beim Erstellen: {e}")
        report_run(sunday, "error", data=data,
                   error_message=str(e),
                   duration_seconds=time.time() - start_time)
        raise

    log.info(f"Fertig: {output}")
    duration = time.time() - start_time

    # 5. Status an Dashboard melden
    from pptx import Presentation
    slide_count = len(Presentation(output).slides)

    report_success(
        sunday_date=sunday,
        data=data,
        song_paths=song_paths,
        missing_songs=missing,
        output_file=os.path.basename(output),
        slide_count=slide_count,
        image_found=image_path is not None,
        duration_seconds=round(duration, 1),
    )

    if missing:
        log.warning("ACHTUNG: Fehlende Lieder – Präsentation manuell prüfen!")

    return output


if __name__ == "__main__":
    setup_logging()

    # Optional: Datum als Argument (YYYY-MM-DD)
    if len(sys.argv) > 1:
        try:
            parts = sys.argv[1].split("-")
            sunday = date(int(parts[0]), int(parts[1]), int(parts[2]))
        except (ValueError, IndexError):
            print(f"Ungültiges Datum: {sys.argv[1]} (Format: YYYY-MM-DD)")
            sys.exit(1)
    else:
        sunday = next_sunday()

    result = run(sunday)
    if result:
        print(f"\nPräsentation erstellt: {result}")
    else:
        print("\nFehler beim Erstellen der Präsentation.")
        sys.exit(1)

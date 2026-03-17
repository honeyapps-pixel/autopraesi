"""Meldet den Status jedes Laufs an Supabase."""

import json
import logging

import requests

log = logging.getLogger(__name__)

SUPABASE_URL = "https://fnzcspteoazchhszeqtf.supabase.co"
SUPABASE_ANON_KEY = (
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9."
    "eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZuemNzcHRlb2F6Y2hoc3plcXRmIiwi"
    "cm9sZSI6ImFub24iLCJpYXQiOjE3NzI5Njg2MzUsImV4cCI6MjA4ODU0NDYzNX0."
    "O6xg-yYQScpL8fglasn8R1PkyqUNGweTbCBrHf2Ea-U"
)


def report_run(sunday_date, status, data=None, error_message=None,
               duration_seconds=None):
    """Schreibt den Lauf-Status in die Supabase-Datenbank.

    Args:
        sunday_date: date-Objekt des Sonntags
        status: 'success', 'partial' oder 'error'
        data: GodiPlanData-Objekt (optional)
        error_message: Fehlermeldung (optional)
        duration_seconds: Laufzeit in Sekunden (optional)
    """
    row = {
        "sunday_date": sunday_date.isoformat(),
        "status": status,
        "error_message": error_message,
        "duration_seconds": duration_seconds,
    }

    if data:
        row["theme"] = data.theme
        row["kirchenkalender"] = data.kirchenkalender
        row["date_str"] = data.date_str

    headers = {
        "apikey": SUPABASE_ANON_KEY,
        "Authorization": f"Bearer {SUPABASE_ANON_KEY}",
        "Content-Type": "application/json",
    }

    try:
        resp = requests.post(
            f"{SUPABASE_URL}/rest/v1/runs",
            headers=headers,
            json=row,
            timeout=10,
        )
        if resp.status_code in (200, 201):
            log.info("Status an Dashboard gemeldet")
        else:
            log.warning(f"Dashboard-Meldung fehlgeschlagen: {resp.status_code} {resp.text}")
    except requests.RequestException as e:
        log.warning(f"Dashboard-Meldung fehlgeschlagen: {e}")


def report_success(sunday_date, data, song_paths, missing_songs,
                   output_file, slide_count, image_found, duration_seconds):
    """Meldet einen erfolgreichen Lauf."""
    songs = []
    for song in data.songs:
        songs.append({
            "slot": song.slot_key,
            "name": song.raw,
            "found": song.slot_key in song_paths,
        })

    status = "partial" if missing_songs else "success"

    row = {
        "sunday_date": sunday_date.isoformat(),
        "status": status,
        "output_file": output_file,
        "slide_count": slide_count,
        "theme": data.theme,
        "kirchenkalender": data.kirchenkalender,
        "date_str": data.date_str,
        "songs": json.dumps(songs),
        "missing_songs": json.dumps(missing_songs),
        "image_found": image_found,
        "duration_seconds": duration_seconds,
        "lesung_reference": data.lesung_reference,
        "predigt1_reference": data.predigt1_reference,
        "predigt1_title": data.predigt1_title,
        "predigt2_reference": data.predigt2_reference,
        "predigt2_title": data.predigt2_title,
        "announcements": json.dumps(data.announcements if data.announcements else []),
    }

    headers = {
        "apikey": SUPABASE_ANON_KEY,
        "Authorization": f"Bearer {SUPABASE_ANON_KEY}",
        "Content-Type": "application/json",
    }

    try:
        resp = requests.post(
            f"{SUPABASE_URL}/rest/v1/runs",
            headers=headers,
            json=row,
            timeout=10,
        )
        if resp.status_code in (200, 201):
            log.info("Status an Dashboard gemeldet")
        else:
            log.warning(f"Dashboard-Meldung fehlgeschlagen: {resp.status_code} {resp.text}")
    except requests.RequestException as e:
        log.warning(f"Dashboard-Meldung fehlgeschlagen: {e}")

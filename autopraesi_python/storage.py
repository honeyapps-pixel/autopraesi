"""Dropbox-Speicher-Abstraktion für AutoPräsi.

Ersetzt den lokalen Dateizugriff (früher: Dropbox-Sync-Ordner auf dem Mac) durch
Dropbox-API-Aufrufe, damit das Backend in der Cloud läuft. Authentifizierung über
einen langlebigen Refresh-Token; die kurzlebigen Access-Tokens erneuert das SDK
selbständig.

Erwartete Umgebungsvariablen:
    DROPBOX_APP_KEY, DROPBOX_APP_SECRET, DROPBOX_REFRESH_TOKEN
"""
from __future__ import annotations

import logging
import os
import tempfile
import time
from typing import Optional

import dropbox
from dropbox.exceptions import ApiError, RateLimitError
from dropbox.files import WriteMode

log = logging.getLogger(__name__)

# Modul-globaler Client (lazy erstellt, damit der Import ohne Env-Vars nicht fehlschlägt)
_dbx: Optional[dropbox.Dropbox] = None

# Maximale Wiederholungen bei Drosselung (HTTP 429)
_MAX_RETRIES = 4


def get_client() -> dropbox.Dropbox:
    """Baut bzw. liefert den zwischengespeicherten Dropbox-Client.

    Nutzt den Refresh-Token-Flow: app_key + app_secret + oauth2_refresh_token.
    Das SDK holt sich daraus automatisch gültige Access-Tokens.
    """
    global _dbx
    if _dbx is None:
        try:
            app_key = os.environ["DROPBOX_APP_KEY"]
            app_secret = os.environ["DROPBOX_APP_SECRET"]
            refresh_token = os.environ["DROPBOX_REFRESH_TOKEN"]
        except KeyError as e:
            raise RuntimeError(
                f"Dropbox-Zugangsdaten fehlen: Umgebungsvariable {e} nicht gesetzt. "
                "Erforderlich: DROPBOX_APP_KEY, DROPBOX_APP_SECRET, DROPBOX_REFRESH_TOKEN."
            ) from e
        _dbx = dropbox.Dropbox(
            oauth2_refresh_token=refresh_token,
            app_key=app_key,
            app_secret=app_secret,
        )
        log.info("Dropbox-Client initialisiert (Refresh-Token-Flow)")
    return _dbx


def _with_retry(fn, *args, **kwargs):
    """Führt einen Dropbox-Aufruf aus und wiederholt ihn bei Drosselung (429)."""
    for attempt in range(_MAX_RETRIES):
        try:
            return fn(*args, **kwargs)
        except RateLimitError as e:
            wait = getattr(e, "backoff", None) or getattr(e.error, "retry_after", None) or 2
            log.warning(f"Dropbox-Drosselung – warte {wait}s (Versuch {attempt + 1}/{_MAX_RETRIES})")
            time.sleep(wait)
    # Letzter Versuch ohne Abfangen, damit der Fehler durchschlägt
    return fn(*args, **kwargs)


def list_folder(path: str, *, recursive: bool = False) -> list:
    """Listet alle Einträge eines Dropbox-Ordners (mit Paging über has_more).

    Returns:
        Liste von dropbox.files.Metadata (FileMetadata / FolderMetadata).
    """
    dbx = get_client()
    entries = []
    res = _with_retry(dbx.files_list_folder, path, recursive=recursive)
    entries.extend(res.entries)
    while res.has_more:
        res = _with_retry(dbx.files_list_folder_continue, res.cursor)
        entries.extend(res.entries)
    return entries


def list_files(path: str, suffix: Optional[str] = None) -> list[tuple[str, str]]:
    """Listet nur Dateien eines Ordners als (name, path_lower)-Tupel.

    Args:
        path: Dropbox-Ordnerpfad (z.B. "/Gemeinde").
        suffix: Optionaler Endungsfilter (z.B. ".pptx", ".xlsx"), case-insensitiv.

    Überspringt Office-Lock-Dateien ("~$...").
    """
    from dropbox.files import FileMetadata

    result = []
    for entry in list_folder(path):
        if not isinstance(entry, FileMetadata):
            continue
        if entry.name.startswith("~$"):
            continue
        if suffix and not entry.name.lower().endswith(suffix.lower()):
            continue
        result.append((entry.name, entry.path_lower))
    return result


def download_bytes(path: str) -> bytes:
    """Lädt eine Datei aus Dropbox und gibt ihren Inhalt als Bytes zurück.

    Geeignet für openpyxl (via io.BytesIO).
    """
    dbx = get_client()
    _metadata, res = _with_retry(dbx.files_download, path)
    return res.content


def download_to_temp(path: str, suffix: str = "") -> str:
    """Lädt eine Datei in eine lokale Temp-Datei und gibt deren Pfad zurück.

    Geeignet für python-pptx, das echte Dateipfade benötigt. Der Aufrufer ist
    für das spätere Aufräumen zuständig (delete=False).
    """
    dbx = get_client()
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    try:
        _metadata, res = _with_retry(dbx.files_download, path)
        tmp.write(res.content)
    finally:
        tmp.close()
    return tmp.name


def upload_file(local_path: str, dropbox_path: str) -> None:
    """Lädt eine lokale Datei nach Dropbox hoch (überschreibt vorhandene)."""
    with open(local_path, "rb") as f:
        upload_bytes(f.read(), dropbox_path)


def upload_bytes(data: bytes, dropbox_path: str) -> None:
    """Lädt Bytes nach Dropbox hoch (überschreibt vorhandene Datei).

    Single-shot-Upload (für Dateien < 150 MB ausreichend; Ausgaben hier ~30 MB).
    """
    dbx = get_client()
    _with_retry(dbx.files_upload, data, dropbox_path, mode=WriteMode.overwrite)
    log.info(f"Nach Dropbox hochgeladen: {dropbox_path} ({len(data)} Bytes)")


def file_exists(path: str) -> bool:
    """Prüft, ob eine Datei/ein Ordner unter dem Dropbox-Pfad existiert."""
    dbx = get_client()
    try:
        _with_retry(dbx.files_get_metadata, path)
        return True
    except ApiError:
        return False

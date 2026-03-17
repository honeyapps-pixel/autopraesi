"""AutoPräsi API – FastAPI Backend für die Web-UI."""

import logging
import os
import tempfile
from dataclasses import asdict

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

from config import IMAGE_DIR, OUTPUT_DIR_DESKTOP, TOGGLEABLE_SECTIONS
from excel_reader import list_all_sheets, read_godi_plan_by_sheet, GodiPlanData, parse_song_entry
from song_finder import build_song_index, find_song
from presentation_builder import build_presentation

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
log = logging.getLogger("autopraesi.api")

app = FastAPI(title="AutoPräsi API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Song-Index einmalig bauen und cachen
_song_index = None


def _get_song_index():
    global _song_index
    if _song_index is None:
        _song_index = build_song_index()
    return _song_index


def _find_image(date_str: str) -> str | None:
    """Sucht das Hintergrundbild (z.B. 'Bild 08.03..jpg')."""
    if not date_str:
        return None
    parts = date_str.split(".")
    if len(parts) < 2:
        return None
    name = f"Bild {parts[0]}.{parts[1]}..jpg"
    path = os.path.join(IMAGE_DIR, name)
    return path if os.path.exists(path) else None


# --- Models ---

class SheetInfo(BaseModel):
    name: str
    excel_path: str


class SongStatus(BaseModel):
    slot_key: str
    raw: str
    category: str
    book: str
    number: str
    title: str
    found: bool
    file_name: str


class SectionInfo(BaseModel):
    key: str
    label: str
    default_enabled: bool


class GenerateRequest(BaseModel):
    sheet_name: str
    excel_path: str
    overrides: dict | None = None
    fetch_bible: bool = True
    disabled_sections: list[str] | None = None  # z.B. ["glaubensbekenntnis", "kinderstunde"]


# --- Endpoints ---

@app.get("/api/sheets", response_model=list[SheetInfo])
def get_sheets():
    """Alle verfügbaren Sheets aus den GoDi-Plan Excel-Dateien."""
    sheets = list_all_sheets()
    return [SheetInfo(name=name, excel_path=path) for name, path in sheets]


@app.get("/api/sections", response_model=list[SectionInfo])
def get_sections():
    """Alle togglebaren Abschnitte im Template."""
    return [
        SectionInfo(key=key, label=sec["label"], default_enabled=True)
        for key, sec in TOGGLEABLE_SECTIONS.items()
    ]


@app.get("/api/sheet/{sheet_name}")
def get_sheet_data(sheet_name: str, excel_path: str):
    """Liest die Daten eines Sheets und gibt sie als JSON zurück."""
    data = read_godi_plan_by_sheet(sheet_name, excel_path, skip_dropbox_sync=True)
    if not data:
        raise HTTPException(404, f"Sheet '{sheet_name}' nicht gefunden")

    index = _get_song_index()
    songs = []
    for song in data.songs:
        path = find_song(song, index)
        songs.append(SongStatus(
            slot_key=song.slot_key,
            raw=song.raw,
            category=song.category,
            book=song.book,
            number=song.number,
            title=song.title,
            found=path is not None,
            file_name=os.path.basename(path) if path else "",
        ))

    image_path = _find_image(data.date_str)

    return {
        "service_header": data.service_header,
        "theme": data.theme,
        "date_str": data.date_str,
        "kirchenkalender": data.kirchenkalender,
        "greeting_verse": data.greeting_verse,
        "lesung_reference": data.lesung_reference,
        "predigt1_reference": data.predigt1_reference,
        "predigt1_title": data.predigt1_title,
        "predigt2_reference": data.predigt2_reference,
        "predigt2_title": data.predigt2_title,
        "is_abendmahl": data.is_abendmahl,
        "songs": [s.model_dump() for s in songs],
        "announcements": data.announcements,
        "invitation_events": [asdict(e) for e in data.invitation_events],
        "image_found": image_path is not None,
        "image_path": image_path,
    }


@app.post("/api/upload-image")
async def upload_image(file: UploadFile = File(...)):
    """Lädt ein Bild hoch und gibt den temporären Pfad zurück."""
    suffix = os.path.splitext(file.filename)[1] if file.filename else ".jpg"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix, dir=tempfile.gettempdir())
    content = await file.read()
    tmp.write(content)
    tmp.close()
    return {"path": tmp.name, "filename": file.filename}


@app.post("/api/generate")
def generate_presentation(req: GenerateRequest):
    """Generiert die Präsentation."""
    data = read_godi_plan_by_sheet(req.sheet_name, req.excel_path, skip_dropbox_sync=True)
    if not data:
        raise HTTPException(404, f"Sheet '{req.sheet_name}' nicht gefunden")

    # Overrides anwenden
    if req.overrides:
        o = req.overrides
        if "theme" in o and o["theme"]:
            data.theme = o["theme"]
        if "greeting_verse" in o and o["greeting_verse"]:
            data.greeting_verse = o["greeting_verse"]
        if "lesung_reference" in o and o["lesung_reference"]:
            data.lesung_reference = o["lesung_reference"]
        if "predigt1_reference" in o and o["predigt1_reference"]:
            data.predigt1_reference = o["predigt1_reference"]
        if "predigt1_title" in o and o["predigt1_title"]:
            data.predigt1_title = o["predigt1_title"]
        if "predigt2_reference" in o and o["predigt2_reference"]:
            data.predigt2_reference = o["predigt2_reference"]
        if "predigt2_title" in o and o["predigt2_title"]:
            data.predigt2_title = o["predigt2_title"]
        if "announcements" in o:
            data.announcements = o["announcements"]

        # Song-Overrides
        if "songs" in o and o["songs"]:
            for slot_key, raw_text in o["songs"].items():
                for i, song in enumerate(data.songs):
                    if song.slot_key == slot_key:
                        data.songs[i] = parse_song_entry(raw_text, slot_key)
                        break

    # Skip-Slides berechnen aus disabled_sections
    skip_slides = set()
    if req.disabled_sections:
        for section_key in req.disabled_sections:
            if section_key in TOGGLEABLE_SECTIONS:
                skip_slides.update(TOGGLEABLE_SECTIONS[section_key]["slides"])
                log.info(f"Abschnitt deaktiviert: {section_key} "
                         f"(Folien {TOGGLEABLE_SECTIONS[section_key]['slides']})")

    # Songs suchen
    index = _get_song_index()
    song_paths = {}
    missing = []
    for song in data.songs:
        path = find_song(song, index)
        if path:
            song_paths[song.slot_key] = path
        elif song.raw:
            missing.append(song.slot_key)

    # Bild
    image_path = req.overrides.get("image_path") if req.overrides else None
    if not image_path:
        image_path = _find_image(data.date_str)

    # Output-Name aus Sheet-Name ableiten
    output_name = f"{req.sheet_name}_ungeprüft.pptx"

    try:
        output = build_presentation(
            data, song_paths,
            image_path=image_path,
            fetch_bible=req.fetch_bible,
            output_name=output_name,
            skip_slides=skip_slides if skip_slides else None,
        )
    except Exception as e:
        log.error(f"Fehler beim Generieren: {e}", exc_info=True)
        raise HTTPException(500, f"Fehler beim Generieren: {e}")

    return {
        "success": True,
        "output_path": output,
        "output_name": os.path.basename(output),
        "missing_songs": missing,
    }


@app.get("/api/download/{filename}")
def download_file(filename: str):
    """Generierte Präsentation herunterladen."""
    path = os.path.join(OUTPUT_DIR_DESKTOP, filename)
    if not os.path.exists(path):
        raise HTTPException(404, f"Datei nicht gefunden: {filename}")
    return FileResponse(path, filename=filename,
                        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")


@app.post("/api/refresh-songs")
def refresh_song_index():
    """Song-Index neu aufbauen."""
    global _song_index
    _song_index = build_song_index()
    count = sum(len(v) for v in _song_index.values())
    return {"success": True, "total_songs": count}

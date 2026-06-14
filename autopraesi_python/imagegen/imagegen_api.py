"""Lokaler Bild-Generator für AutoPräsi (läuft NUR auf dem Mac Mini).

Erzeugt christliche Gottesdienst-Hintergrundbilder per MFLUX (z-image-turbo, MLX-nativ,
schnellst). Dieser Dienst ist bewusst von ``api.py`` (Cloud/Render) getrennt: Render kann
MFLUX nicht ausführen, also läuft die Generierung lokal und wird über einen Cloudflared-Tunnel
für die Web-App erreichbar gemacht.

Der Dienst GENERIERT nur und hält die Kandidaten in ``_work/``. Den finalen Dropbox-Upload
macht ausschließlich das Cloud-Backend (``/api/save-sunday-image``) – so braucht dieser Dienst
keine Dropbox-Zugangsdaten.

Start:
    ~/.mflux-venv/bin/uvicorn imagegen_api:app --host 127.0.0.1 --port 8189
(siehe start-imagegen.sh – startet zusätzlich den Tunnel)
"""
from __future__ import annotations

import logging
import os
import threading
import time
import uuid
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
log = logging.getLogger("autopraesi.imagegen")

# --- Pfade ---
HERE = Path(__file__).resolve().parent
WORK_DIR = HERE / "_work"
WORK_DIR.mkdir(exist_ok=True)

# Bildgröße = Folienformat 4:3 (PowerPoint 9144000×6858000 EMU). 1024×768 ist 4:3,
# durch 16 teilbar und ein guter Kompromiss aus Tempo und Qualität für einen
# Folien-Hintergrund (Beamer-Auflösung). Größer = deutlich langsamer.
WIDTH = int(os.environ.get("IMG_WIDTH", "1024"))
HEIGHT = int(os.environ.get("IMG_HEIGHT", "768"))
STEPS = int(os.environ.get("IMG_STEPS", "8"))
# quantize 4 ist auf Apple Silicon spürbar schneller als 8/bf16. None = bf16.
_q = os.environ.get("IMG_QUANTIZE", "4")
QUANTIZE: Optional[int] = None if _q.lower() in ("", "none", "0") else int(_q)

# Kandidaten älter als N Stunden beim Start aufräumen.
MAX_AGE_HOURS = int(os.environ.get("IMG_MAX_AGE_HOURS", "48"))

# --- Prompt-Stil: christliches Gottesdienst-Hintergrundbild --------------------
# Fester Sakral-Stil, damit der Nutzer kein Prompt-Wissen braucht. Bewusst
# atmosphärisch/symbolisch statt figürlich (KI-Gesichter/Jesus-Darstellungen
# wirken oft unschön). Ruhige Komposition mit Freiraum, da Titel/Datum später
# über das Bild gelegt werden.
# WICHTIG: keine Wörter wie "slide"/"text"/"title" im POSITIV-Prompt – sie verleiten das
# Modell dazu, Schrift ins Bild zu zeichnen. Der Folientitel/Datum/Vers wird später von der
# Präsentation ÜBER das Bild gelegt, gehört also NICHT ins generierte Bild.
#
# Stil an den vorhandenen Gemeinde-Hintergrundbildern ausgerichtet (Ordner
# /Gemeinde/Brüderrecords°/Hintergrundbilder): durchweg fotografische, wallpaper-artige
# Motive – ruhige Natur (Felder, Wälder, Wasser, Himmel, Sonnenuntergänge), dezente
# christliche Symbolik (Kreuz-Silhouette, aufgeschlagene Bibel, Kerzenlicht, betende Hände
# von hinten), warmes Bokeh – fast immer ruhig komponiert mit freier Fläche für Text,
# warmes/weiches Licht, emotional aber nicht kitschig.
STYLE = (
    "A professional, high-quality photographic background image for a Christian worship "
    "setting, in the style of a clean cinematic widescreen wallpaper. Reverent, hopeful and "
    "emotionally warm, with beautiful natural or divine light — golden hour, soft god rays, "
    "gentle glow. Calm, uncluttered and well-balanced composition with a simple area of "
    "negative space, so it works well as a background. Tasteful subjects such as serene "
    "landscapes, golden fields, misty forests, calm water, dramatic skies and sunsets, or "
    "quiet Christian symbolism like a distant cross silhouette, an open Bible or candlelight. "
    "Fine-art photography, shallow depth of field and soft bokeh, rich but natural color, "
    "dignified, evocative, photorealistic."
)
# Schrift hart ausschließen (Negation funktioniert im Negativ-Prompt zuverlässiger als im Positiv).
NEGATIVE = (
    "text, words, letters, captions, title, heading, subtitle, typography, font, writing, "
    "handwriting, inscription, signage, sign, banner, poster, label, watermark, logo, "
    "signature, numbers, frame, border, "
    "human face, portrait, people close-up, deformed hands, kitsch, cartoon, anime, "
    "low quality, blurry, cluttered, busy composition, oversaturated"
)


def _build_prompt(theme: str, wochenspruch: str, freitext: str) -> str:
    """Baut den Bildprompt – reine SZENEN-Beschreibung aus Freitext + Stil.

    Thema und Wochenspruch werden BEWUSST NICHT in den Bildprompt übernommen: kurze, slogan-
    artige Phrasen (z.B. ein Thema) malt das Modell sonst manchmal als Titel ins Bild. Sie
    dienen im Formular nur als Inspiration für den Freitext. Der eigentliche Text (Titel,
    Datum, Vers) liegt später als Overlay auf der Folie – nicht im generierten Bild.
    """
    parts: list[str] = []
    if freitext and freitext.strip():
        parts.append(freitext.strip())
    parts.append(STYLE)
    # Bewusst KEIN "no text" im Positiv-Prompt – Negationen werden dort schlecht verstanden
    # und führen eher zu Schrift. Textausschluss passiert über NEGATIVE.
    return ". ".join(parts)


def _cleanup_old() -> None:
    cutoff = time.time() - MAX_AGE_HOURS * 3600
    for p in WORK_DIR.glob("*.png"):
        try:
            if p.stat().st_mtime < cutoff:
                p.unlink()
        except OSError:
            pass


# --- Persistentes Modell (lädt einmal, bleibt im Speicher) ---------------------
# Auf dieser Hardware ist die Inferenz der Kostentreiber, nicht das Laden. Den
# Subprozess-pro-Request-Weg vermeiden wir, weil dabei jedes Mal die MLX-Kernel neu
# kompiliert würden (erster Step ~2× so teuer). Eine Lock serialisiert die GPU-Nutzung,
# da FastAPI sync-Endpunkte im Threadpool laufen.
_model = None


def _get_model():
    global _model
    if _model is None:
        from mflux.models.common.config import ModelConfig
        from mflux.models.z_image.variants.z_image import ZImage
        log.info("Lade Modell z-image-turbo (quantize=%s) …", QUANTIZE)
        t0 = time.time()
        _model = ZImage(model_config=ModelConfig.z_image_turbo(), quantize=QUANTIZE)
        log.info("Modell geladen in %.1fs", time.time() - t0)
    return _model


# --- Job-Queue (asynchron) -----------------------------------------------------
# Ein Bild dauert auf dieser Hardware ~1 Min. Über den Cloudflared-Tunnel würde ein
# so lange gehaltener HTTP-Request am Edge (Cloudflare 524 nach ~100s) abbrechen.
# Daher: /generate legt Jobs an und kehrt SOFORT zurück; ein einzelner Worker-Thread
# erzeugt sie seriell (GPU-sicher); das Frontend pollt /status. Jeder Request ist kurz.
import queue

_jobs: dict[str, dict] = {}          # id -> {status, seed, error, prompt}
_jobs_lock = threading.Lock()
_job_queue: "queue.Queue[str]" = queue.Queue()


def _worker() -> None:
    while True:
        job_id = _job_queue.get()
        try:
            with _jobs_lock:
                job = _jobs.get(job_id)
            if not job or job["status"] == "deleted":
                continue
            t0 = time.time()
            model = _get_model()  # lädt einmal, bleibt im Speicher
            image = model.generate_image(
                seed=job["seed"], prompt=job["prompt"],
                width=WIDTH, height=HEIGHT,
                num_inference_steps=STEPS, negative_prompt=NEGATIVE,
            )
            image.save(path=str(WORK_DIR / f"{job_id}.png"))
            with _jobs_lock:
                cur = _jobs.get(job_id)
                if not cur or cur["status"] == "deleted":
                    (WORK_DIR / f"{job_id}.png").unlink(missing_ok=True)
                else:
                    cur["status"] = "done"
            log.info("Bild %s fertig in %.1fs (seed %s)", job_id[:8], time.time() - t0, job["seed"])
        except Exception as e:  # noqa: BLE001 – Job-Fehler isolieren, Worker läuft weiter
            log.error("Job %s fehlgeschlagen: %s", job_id[:8], e, exc_info=True)
            with _jobs_lock:
                if job_id in _jobs:
                    _jobs[job_id]["status"] = "error"
                    _jobs[job_id]["error"] = str(e)
        finally:
            _job_queue.task_done()


threading.Thread(target=_worker, daemon=True, name="imagegen-worker").start()


def _enqueue(prompt: str, count: int, seed: Optional[int]) -> list[dict]:
    """Legt ``count`` Jobs an und gibt [{id, seed}] (Status: pending) zurück."""
    if seed is not None:
        seeds = [seed + i for i in range(count)]
    else:
        base = int(time.time() * 1000) % 1_000_000
        seeds = [base + i * 7919 for i in range(count)]
    out: list[dict] = []
    for s in seeds:
        job_id = uuid.uuid4().hex
        with _jobs_lock:
            _jobs[job_id] = {"status": "pending", "seed": s, "error": None, "prompt": prompt}
        _job_queue.put(job_id)
        out.append({"id": job_id, "seed": s})
    return out


# --- App -----------------------------------------------------------------------

app = FastAPI(title="AutoPräsi Bild-Generator (lokal)")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

_cleanup_old()


class GenerateRequest(BaseModel):
    theme: str = ""
    wochenspruch: str = ""
    freitext: str = ""
    count: int = Field(default=1, ge=1, le=3)
    seed: Optional[int] = None


class RegenerateRequest(BaseModel):
    theme: str = ""
    wochenspruch: str = ""
    freitext: str = ""
    seed: Optional[int] = None


def _status_of(job_id: str) -> str:
    with _jobs_lock:
        job = _jobs.get(job_id)
    return job["status"] if job else "unknown"


def _image_payload(item: dict) -> dict:
    return {
        "id": item["id"],
        "seed": item.get("seed"),
        "url": f"/image/{item['id']}",
        "status": "pending",
    }


@app.get("/health")
def health():
    with _jobs_lock:
        pending = sum(1 for j in _jobs.values() if j["status"] == "pending")
    return {"ok": True, "model": "z-image-turbo", "loaded": _model is not None,
            "size": f"{WIDTH}x{HEIGHT}", "steps": STEPS, "pending": pending}


@app.post("/generate")
def generate(req: GenerateRequest):
    """Legt 1–3 Bild-Jobs an und kehrt sofort zurück (Status: pending)."""
    prompt = _build_prompt(req.theme, req.wochenspruch, req.freitext)
    items = _enqueue(prompt, req.count, req.seed)
    return {"images": [_image_payload(it) for it in items], "prompt": prompt}


@app.post("/regenerate")
def regenerate(req: RegenerateRequest):
    """Legt EINEN neuen Bild-Job an (geänderter Prompt/Seed) = „Bearbeiten"."""
    prompt = _build_prompt(req.theme, req.wochenspruch, req.freitext)
    items = _enqueue(prompt, 1, req.seed)
    return {**_image_payload(items[0]), "prompt": prompt}


@app.get("/status")
def status(ids: str):
    """Status mehrerer Jobs. ``ids`` = kommaseparierte Job-IDs."""
    out = []
    for jid in ids.split(","):
        jid = jid.strip()
        if not jid:
            continue
        with _jobs_lock:
            job = _jobs.get(jid)
        out.append({
            "id": jid,
            "status": job["status"] if job else "unknown",
            "error": job.get("error") if job else None,
        })
    return out


@app.get("/image/{img_id}")
def get_image(img_id: str):
    # Pfad-Traversal verhindern: nur reine Hex-IDs zulassen.
    if not img_id.isalnum():
        raise HTTPException(400, "Ungültige ID")
    path = WORK_DIR / f"{img_id}.png"
    if not path.exists():
        raise HTTPException(404, "Bild noch nicht fertig oder nicht vorhanden")
    return FileResponse(str(path), media_type="image/png")


@app.delete("/image/{img_id}")
def delete_image(img_id: str):
    if not img_id.isalnum():
        raise HTTPException(400, "Ungültige ID")
    # Job markieren (falls noch in der Queue/Worker), Datei entfernen.
    with _jobs_lock:
        job = _jobs.get(img_id)
        if job:
            job["status"] = "deleted"
    path = WORK_DIR / f"{img_id}.png"
    existed = path.exists()
    if existed:
        path.unlink()
    return {"deleted": existed or job is not None}

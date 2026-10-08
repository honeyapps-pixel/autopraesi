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

import json
import logging
import os
import threading
import time
import urllib.request
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

# --- KI-Szenenfindung (LOKAL via Ollama) --------------------------------------
# Thema + Wochenspruch werden von einem LOKALEN Sprachmodell (Ollama, läuft auf dem
# Mac Mini – KEIN API-Key, keine Cloud) in konkrete, text-freie Bildmotive übersetzt.
# Damit gehen die Bilder WIRKLICH auf Thema/Wochenspruch ein, sind untereinander
# deutlich verschieden und kommen bei jeder Generierung neu heraus.
# Fällt Ollama aus (Dienst nicht erreichbar, Modell fehlt), greift still die feste
# ``VARIATIONS``-Liste unten – der Dienst läuft normal weiter.
OLLAMA_URL = os.environ.get("IMG_OLLAMA_URL", "http://localhost:11434").rstrip("/")
SCENE_MODEL = os.environ.get("IMG_SCENE_MODEL", "qwen2.5:32b")
# Sprachmodell nach der Szenenfindung SOFORT entladen ("0"): qwen2.5:32b belegt ~20 GB und
# würde sonst neben dem Bildmodell im Speicher bleiben → der Mac swappt und die Bild-
# generierung bleibt minutenlang hängen. Neu laden kostet nur ~10–20 s pro Generierung.
SCENE_KEEP_ALIVE = os.environ.get("IMG_SCENE_KEEP_ALIVE", "0")
SCENE_TIMEOUT = int(os.environ.get("IMG_SCENE_TIMEOUT", "90"))    # s; kalt lädt Modell erst
# Gedächtnis gegen Wiederholung: zuletzt benutzte Motive, die das Modell meiden soll.
HISTORY_FILE = WORK_DIR / "recent_scenes.json"
HISTORY_KEEP = 30          # so viele jüngste Motive aufheben
HISTORY_AVOID = 24         # so viele davon dem Modell als „nicht wiederholen" mitgeben

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


# FALLBACK-Bildvarianten (nur ohne KI-Szenenfindung): feste Motive, die rotieren,
# damit bei 2–3 Bildern nicht 3× fast dasselbe entsteht. Auf eure Hintergrundbilder
# abgestimmt (ruhige Natur + dezente christliche Symbolik). Gehen NICHT aufs Thema ein –
# der reguläre Weg ist die KI-Szenenfindung weiter unten.
VARIATIONS = [
    "an open golden field under a vast sunrise sky, soft warm light",
    "a misty forest with sunbeams streaming through the trees",
    "calm reflective water at dawn with a soft glowing sky and gentle mist",
    "a distant cross silhouette on a hill against a dramatic golden sunset",
    "rolling hills under dramatic clouds with strong god rays",
    "a close, intimate nature detail with shallow depth of field and warm bokeh",
    "a quiet path leading toward a luminous horizon at golden hour",
    "soft candlelight near an open Bible in warm gentle focus",
    "a vast starry night sky over a calm silent landscape",
    "a coastal horizon at dawn with soft pastel colors and calm sea",
]


def _build_prompt(theme: str, wochenspruch: str, freitext: str, variation: str = "") -> str:
    """Baut den finalen Bildprompt: reine SZENEN-Beschreibung (+ optional Freitext) + Stil.

    ``variation`` ist das eigentliche Motiv – entweder eine KI-Szene (aus Thema/Wochenspruch
    abgeleitet, s. ``_derive_scenes``) oder ein festes Fallback-Motiv. Thema und Wochenspruch
    selbst stehen hier NICHT als Klartext im Prompt: kurze Slogans malt das Modell sonst als
    Titel ins Bild; ihr Sinn fließt stattdessen über die KI-Szene ein. Der eigentliche Text
    (Titel, Datum, Vers) liegt später als Overlay auf der Folie – nicht im generierten Bild.
    """
    parts: list[str] = []
    if freitext and freitext.strip():
        parts.append(freitext.strip())
    if variation:
        parts.append(variation)
    parts.append(STYLE)
    # Bewusst KEIN "no text" im Positiv-Prompt – Negationen werden dort schlecht verstanden
    # und führen eher zu Schrift. Textausschluss passiert über NEGATIVE.
    return ". ".join(parts)


# --- KI-Szenenfindung ----------------------------------------------------------
# Wandelt Thema + Wochenspruch (+ Freitext) in konkrete, text-freie Bildmotive um –
# symbolisch passend, untereinander verschieden, und mit einem Gedächtnis, das schon
# benutzte Motive meidet (gegen „war letzte Woche schon"). Läuft lokal über Ollama.
_history_lock = threading.Lock()

SCENE_SYSTEM = (
    "You are an art director choosing the photographic background image for the slides of a "
    "German Protestant church service (Gottesdienst). You are given the sermon theme (Thema) and "
    "the weekly Bible verse (Wochenspruch) in German, plus an optional free-text idea. Invent "
    "vivid background MOTIFS for a text-to-image model.\n\n"
    "Rules:\n"
    "- Output ENGLISH descriptions only.\n"
    "- Each motif is ONE concise phrase of roughly 12-28 words: subject, setting, light and mood "
    "as comma-separated visual descriptors (no full sentences about people doing things).\n"
    "- Connect to the MEANING and imagery of the theme and verse symbolically and atmospherically; "
    "evoke the feeling, never illustrate the literal words.\n"
    "- Absolutely NO text, letters, words, captions, signage or writing anywhere in the image.\n"
    "- No close-up human faces and no depiction of Jesus or identifiable holy figures (AI faces "
    "look wrong); people only as distant silhouettes, seen from behind, or just hands.\n"
    "- Photographic, wallpaper-like, reverent, hopeful, with warm natural or divine light, and a "
    "calm composition that leaves empty negative space so a title can be overlaid.\n"
    "- Stay in the congregation's house style: quiet nature (fields, forests, water, skies, "
    "sunrises and sunsets) and gentle Christian symbolism (a distant cross silhouette, an open "
    "Bible, candlelight, bread and a cup for Holy Communion, a single lit candle).\n"
    "- Make the requested motifs clearly DIFFERENT from one another in subject, palette, time of "
    "day and composition, and different from any 'recently used' motifs provided.\n\n"
    'Respond with ONLY a JSON object of the form {"scenes": ["motif one", "motif two"]} — '
    "no markdown code fences and no text outside the JSON."
)


def _ollama_post(path: str, payload: dict, timeout: int) -> dict:
    """POST an die lokale Ollama-API; gibt das JSON zurück (wirft bei Fehler)."""
    req = urllib.request.Request(
        f"{OLLAMA_URL}{path}",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8"))


def _scene_ai_ready() -> bool:
    """True, wenn die lokale Ollama-API erreichbar ist (für /health). Kurzer Probe-Call."""
    try:
        with urllib.request.urlopen(f"{OLLAMA_URL}/api/tags", timeout=3) as r:
            return r.status == 200
    except Exception:  # noqa: BLE001 – Ollama aus -> Fallback-Motive
        return False


def _load_recent() -> list[str]:
    with _history_lock:
        try:
            data = json.loads(HISTORY_FILE.read_text("utf-8"))
            return [s for s in data if isinstance(s, str)][-HISTORY_AVOID:]
        except Exception:  # noqa: BLE001 – fehlende/kaputte Datei = leere Historie
            return []


def _remember(scenes: list[str]) -> None:
    clean = [s.strip() for s in scenes if s and s.strip()]
    if not clean:
        return
    with _history_lock:
        try:
            old = [s for s in json.loads(HISTORY_FILE.read_text("utf-8")) if isinstance(s, str)]
        except Exception:  # noqa: BLE001
            old = []
        merged = (old + clean)[-HISTORY_KEEP:]
        try:
            HISTORY_FILE.write_text(json.dumps(merged, ensure_ascii=False), "utf-8")
        except OSError as e:
            log.warning("Szenen-Historie nicht speicherbar: %s", e)


def _parse_scenes(text: str) -> list[str]:
    """Tolerant JSON aus der Modellantwort lesen (auch wenn ```-Fences o.ä. drumherum sind)."""
    import re
    t = text.strip()
    if t.startswith("```"):
        t = re.sub(r"^```[a-zA-Z]*\s*", "", t).rstrip("`").strip()
    try:
        obj = json.loads(t)
    except Exception:  # noqa: BLE001 – erstes {...} herausgreifen
        m = re.search(r"\{.*\}", t, re.S)
        if not m:
            return []
        try:
            obj = json.loads(m.group(0))
        except Exception:  # noqa: BLE001
            return []
    scenes = obj.get("scenes") if isinstance(obj, dict) else obj
    if not isinstance(scenes, list):
        return []
    return [s.strip() for s in scenes if isinstance(s, str) and s.strip()]


def _derive_scenes(theme: str, wochenspruch: str, freitext: str,
                   count: int, avoid: list[str]) -> Optional[list[str]]:
    """Übersetzt Thema+Wochenspruch(+Freitext) in ``count`` text-freie Bildmotive.

    Gibt None zurück, wenn das lokale Modell (Ollama) nicht erreichbar ist oder der
    Aufruf scheitert – die Aufrufer fallen dann auf die festen ``VARIATIONS`` zurück.
    """
    user = (
        f"Thema: {theme.strip() or '(keins)'}\n"
        f"Wochenspruch: {wochenspruch.strip() or '(keiner)'}\n"
        f"Freitext-Bildidee: {freitext.strip() or '(keine)'}\n\n"
        f"Gib genau {count} deutlich unterschiedliche Motive zurück."
    )
    if avoid:
        user += "\n\nRecently used (avoid resembling these):\n- " + "\n- ".join(avoid)
    try:
        # format="json" zwingt Ollama zu sauberem JSON.
        resp = _ollama_post("/api/chat", {
            "model": SCENE_MODEL,
            "stream": False,
            "format": "json",
            "keep_alive": SCENE_KEEP_ALIVE,  # "0" = nach der Antwort entladen
            "options": {"temperature": 1.0},  # hoch = mehr Abwechslung über Generierungen
            "messages": [
                {"role": "system", "content": SCENE_SYSTEM},
                {"role": "user", "content": user},
            ],
        }, timeout=SCENE_TIMEOUT)
        scenes = _parse_scenes(resp.get("message", {}).get("content", ""))
    except Exception as e:  # noqa: BLE001 – jeder Fehler (Ollama aus/Timeout) -> Fallback
        log.warning("KI-Szenenfindung (Ollama) fehlgeschlagen (%s) – nutze feste Varianten.", e)
        return None
    if not scenes:
        return None
    # Genau ``count`` zurückgeben (zyklisch auffüllen, falls das Modell zu wenige lieferte).
    if len(scenes) < count:
        base = list(scenes)
        while len(scenes) < count:
            scenes.append(base[len(scenes) % len(base)])
    return scenes[:count]


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


def _enqueue(theme: str, wochenspruch: str, freitext: str, count: int,
             seed: Optional[int]) -> list[dict]:
    """Legt ``count`` Jobs an – jeder mit eigenem, themenbezogenem Motiv – und gibt [{id, seed}].

    Bevorzugt KI-Motive (passend zu Thema/Wochenspruch, immer neu, untereinander deutlich
    verschieden, mit Gedächtnis gegen Wiederholung). Ohne KI: rotierende feste ``VARIATIONS``
    (alter Weg). Jedes Bild bekommt zusätzlich einen eigenen Seed.
    """
    if seed is not None:
        seeds = [seed + i * 9973 for i in range(count)]
    else:
        base = int(time.time() * 1000) % 1_000_000
        seeds = [base + i * 7919 for i in range(count)]

    scenes = _derive_scenes(theme, wochenspruch, freitext, count, _load_recent())
    ki = scenes is not None
    if not ki:
        # Fallback: feste Varianten rotieren (zeitbasiert, damit nicht immer dieselben).
        offset = int(time.time())
        scenes = [VARIATIONS[(offset + i) % len(VARIATIONS)] for i in range(count)]

    out: list[dict] = []
    for i, s in enumerate(seeds):
        scene = scenes[i] if i < len(scenes) else scenes[i % len(scenes)]
        # KI-Motiv enthält den Freitext schon; im Fallback Freitext separat voranstellen.
        prompt = _build_prompt(theme, wochenspruch, "" if ki else freitext, scene)
        job_id = uuid.uuid4().hex
        with _jobs_lock:
            _jobs[job_id] = {"status": "pending", "seed": s, "error": None, "prompt": prompt}
        _job_queue.put(job_id)
        out.append({"id": job_id, "seed": s})

    if ki:
        _remember(scenes)  # schon benutzte Motive merken -> nächste Generierung meidet sie
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
            "size": f"{WIDTH}x{HEIGHT}", "steps": STEPS, "pending": pending,
            "scene_ai": _scene_ai_ready(), "scene_model": SCENE_MODEL}


@app.post("/generate")
def generate(req: GenerateRequest):
    """Legt 1–3 Bild-Jobs an (je eigene Variante) und kehrt sofort zurück (Status: pending)."""
    items = _enqueue(req.theme, req.wochenspruch, req.freitext, req.count, req.seed)
    return {"images": [_image_payload(it) for it in items]}


@app.post("/regenerate")
def regenerate(req: RegenerateRequest):
    """Legt EINEN neuen Bild-Job an (frische Variante/Seed) = „Bearbeiten"."""
    items = _enqueue(req.theme, req.wochenspruch, req.freitext, 1, req.seed)
    return _image_payload(items[0])


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

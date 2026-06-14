# AutoPräsi – Lokaler Bild-Generator

Erzeugt christliche Gottesdienst-Hintergrundbilder per **MFLUX `z-image-turbo`** (MLX-nativ,
schnellst) lokal auf dem **Mac Mini**. Wird vom Reiter „Bilder generieren" der Web-App genutzt.

Warum lokal? Das Cloud-Backend (Render) kann MFLUX nicht ausführen. Dieser Dienst läuft auf dem
Mac und wird über einen **ngrok-Tunnel mit fester Domain** für die Vercel-App erreichbar. Er
**generiert nur**; den finalen Dropbox-Upload macht das Cloud-Backend (`/api/save-sunday-image`)
– dieser Dienst braucht daher keine Dropbox-Zugangsdaten.

Das Modell wird **in-process** geladen (mflux als Python-Bibliothek) und bleibt im Speicher –
deshalb läuft der Dienst zwingend im `~/.mflux-venv`. So wird das MLX-Kernel-Kompilieren nur
einmal bezahlt; jedes weitere Bild ist schneller. Auf dieser Hardware ist die Inferenz selbst
der Kostentreiber (~1 Min/Bild bei 1024×768/8 Steps), nicht das Laden.

## Einmalig einrichten

```bash
~/.mflux-venv/bin/pip install fastapi "uvicorn[standard]"
# MFLUX ist bereits in ~/.mflux-venv installiert; ngrok mit Authtoken konfiguriert.
```

## Starten

```bash
cd autopraesi_python/imagegen
./start-imagegen.sh
```

Startet uvicorn (Port 8189) + ngrok auf der **festen Domain**
`https://tawniest-uxorially-boyce.ngrok-free.dev`. Diese Domain ist in der Web-App als
`NEXT_PUBLIC_IMAGEGEN_URL` (Vercel) hinterlegt – solange `start-imagegen.sh` dieselbe Domain
nutzt, **funktioniert der Reiter ohne weitere Eingabe**. Im Reiter gibt es nur einen Button
„Generator prüfen", der Erreichbarkeit meldet.

> ngrok-Free zeigt Browser-Anfragen sonst eine Warnseite – die App umgeht das per Header
> `ngrok-skip-browser-warning`; Vorschaubilder werden als Blob geladen.

## Konfiguration (ENV)

| Variable | Default | Zweck |
|---|---|---|
| `IMG_PORT` | `8189` | Port des Dienstes |
| `IMAGEGEN_DOMAIN` | `tawniest-uxorially-boyce.ngrok-free.dev` | feste ngrok-Domain |
| `IMG_WIDTH`/`IMG_HEIGHT` | `1024`/`768` | Bildgröße (4:3 = Folienformat) |
| `IMG_STEPS` | `8` | Inferenz-Schritte (weniger = schneller) |
| `IMG_QUANTIZE` | `4` | Quantisierung (`4` schnellste, `8`/`none` höhere Treue) |
| `IMG_MAX_AGE_HOURS` | `48` | Kandidaten älter als … beim Start löschen |

## Endpunkte

- `GET /health`
- `POST /generate` `{ theme, wochenspruch, freitext, count(1–3), seed? }`
- `POST /regenerate` `{ theme, wochenspruch, freitext, seed? }` → ein Bild
- `GET /status?ids=` · `GET /image/{id}` · `DELETE /image/{id}`

Kandidaten liegen in `_work/` (nicht in Git).

#!/usr/bin/env bash
# Startet den lokalen Bild-Generator (MFLUX) + einen Cloudflared-Tunnel, damit die
# Vercel-App ihn von überall erreicht. Läuft auf dem Mac Mini.
#
#   ./start-imagegen.sh                 # Quick-Tunnel (zufällige URL, in der App eintragen)
#   ./start-imagegen.sh <tunnel-name>   # benannter Cloudflared-Tunnel (stabile URL, empfohlen)
#
# Der Dienst nutzt das mflux-venv (Python + FastAPI + uvicorn müssen dort installiert sein):
#   ~/.mflux-venv/bin/pip install fastapi "uvicorn[standard]"
set -euo pipefail
cd "$(dirname "$0")"

PORT="${IMG_PORT:-8189}"
VENV="${MFLUX_VENV:-$HOME/.mflux-venv}"
UVICORN="$VENV/bin/uvicorn"

[ -x "$UVICORN" ] || { echo "❌ uvicorn nicht im venv: $UVICORN  (pip install fastapi 'uvicorn[standard]')" >&2; exit 1; }

echo "▶ Starte Generator-Dienst auf 127.0.0.1:$PORT …"
"$UVICORN" imagegen_api:app --host 127.0.0.1 --port "$PORT" &
API_PID=$!
trap 'kill "$API_PID" 2>/dev/null || true' EXIT

# Kurz warten, bis der Dienst hochgekommen ist.
sleep 2

echo "▶ Starte Cloudflared-Tunnel …"
if [ "${1:-}" != "" ]; then
  # Benannter Tunnel: stabile URL (vorher via `cloudflared tunnel create <name>` + DNS-Route).
  cloudflared tunnel run --url "http://127.0.0.1:$PORT" "$1"
else
  # Quick-Tunnel: gibt eine zufällige https://*.trycloudflare.com-URL aus.
  # Diese URL im Reiter „Bilder generieren" als Generator-URL eintragen.
  cloudflared tunnel --url "http://127.0.0.1:$PORT"
fi

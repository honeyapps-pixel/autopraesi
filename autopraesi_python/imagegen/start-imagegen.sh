#!/usr/bin/env bash
# Startet den lokalen Bild-Generator (MFLUX) + einen ngrok-Tunnel mit FESTER Domain,
# damit die Vercel-App ihn ohne weitere Konfiguration erreicht. Läuft auf dem Mac Mini.
#
#   ./start-imagegen.sh
#
# Die feste Domain ist in der App als NEXT_PUBLIC_IMAGEGEN_URL hinterlegt – solange
# hier dieselbe Domain genutzt wird, "funktioniert es einfach" nach jedem Neustart.
# Voraussetzungen: ngrok mit Authtoken konfiguriert; fastapi+uvicorn im mflux-venv:
#   ~/.mflux-venv/bin/pip install fastapi "uvicorn[standard]"
set -euo pipefail
cd "$(dirname "$0")"

PORT="${IMG_PORT:-8189}"
DOMAIN="${IMAGEGEN_DOMAIN:-tawniest-uxorially-boyce.ngrok-free.dev}"
VENV="${MFLUX_VENV:-$HOME/.mflux-venv}"
UVICORN="$VENV/bin/uvicorn"

[ -x "$UVICORN" ] || { echo "❌ uvicorn nicht im venv: $UVICORN  (pip install fastapi 'uvicorn[standard]')" >&2; exit 1; }
command -v ngrok >/dev/null || { echo "❌ ngrok nicht gefunden (brew install ngrok)" >&2; exit 1; }

# Altprozesse sauber beenden (verhindert Port-/ngrok-Sitzungskonflikte beim Neustart).
pkill -f "uvicorn imagegen_api" 2>/dev/null || true
pkill -f "ngrok http $PORT" 2>/dev/null || true
sleep 1

echo "▶ Starte Generator-Dienst auf 127.0.0.1:$PORT …"
"$UVICORN" imagegen_api:app --host 127.0.0.1 --port "$PORT" &
API_PID=$!
trap 'kill "$API_PID" 2>/dev/null || true' EXIT

sleep 2

echo "▶ Starte ngrok-Tunnel auf https://$DOMAIN …"
ngrok http "$PORT" --url="https://$DOMAIN"

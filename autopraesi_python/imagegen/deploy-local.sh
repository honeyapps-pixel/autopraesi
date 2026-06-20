#!/usr/bin/env bash
# Übernimmt Code-Änderungen am lokalen Bild-Generator in die laufende Auto-Start-Kopie
# und startet den LaunchAgent neu.
#
# Warum eine Kopie? macOS (TCC) verbietet Hintergrund-Diensten (launchd) den Zugriff auf
# ~/Desktop/~/Documents/~/Downloads. Das Repo liegt unter ~/Desktop, also läuft der Dienst
# aus einer Kopie unter ~/Library/Application Support (nicht geschützt). Quelle bleibt das Repo.
#
#   ./deploy-local.sh            # nach Änderungen an imagegen_api.py / start-imagegen.sh
#   ./deploy-local.sh --install  # zusätzlich LaunchAgent (neu) installieren
set -euo pipefail
cd "$(dirname "$0")"

DST="$HOME/Library/Application Support/AutoPraesi/imagegen"
PLIST="$HOME/Library/LaunchAgents/com.autopraesi.imagegen.plist"
LABEL="com.autopraesi.imagegen"

mkdir -p "$DST/_work"
cp imagegen_api.py start-imagegen.sh "$DST/"
chmod +x "$DST/start-imagegen.sh"
echo "✅ Kopiert nach $DST"

if [ "${1:-}" = "--install" ]; then
  cp com.autopraesi.imagegen.plist "$PLIST"
  echo "✅ LaunchAgent installiert: $PLIST"
fi

UID_N=$(id -u)
if [ -f "$PLIST" ]; then
  launchctl bootout "gui/$UID_N/$LABEL" 2>/dev/null || true
  launchctl bootstrap "gui/$UID_N" "$PLIST"
  launchctl kickstart -k "gui/$UID_N/$LABEL"
  echo "✅ Agent neu gestartet. Health in ~10s:"
  sleep 10
  curl -s -m 12 -H "ngrok-skip-browser-warning: true" \
    https://tawniest-uxorially-boyce.ngrok-free.dev/health && echo " <- öffentlich OK" || echo "noch nicht erreichbar – Logs: /tmp/imagegen-agent.err"
else
  echo "ℹ️  LaunchAgent nicht installiert. Mit --install einrichten."
fi

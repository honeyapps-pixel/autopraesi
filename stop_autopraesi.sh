#!/bin/bash
# AutoPräsi stoppen
LOG_DIR="$HOME/Desktop/Projekte/Autopräsi/autopraesi_github/logs"

kill $(cat "$LOG_DIR/backend.pid" 2>/dev/null) 2>/dev/null && echo "Backend gestoppt" || echo "Backend war nicht aktiv"
kill $(cat "$LOG_DIR/ngrok.pid" 2>/dev/null) 2>/dev/null && echo "ngrok gestoppt" || echo "ngrok war nicht aktiv"
kill $(lsof -ti :8000) 2>/dev/null || true
kill $(pgrep -f ngrok) 2>/dev/null || true

echo "AutoPräsi gestoppt."

#!/bin/bash
# AutoPräsi — Startet Backend + ngrok und updated Vercel automatisch
set -e

PROJECT_DIR="$HOME/Desktop/Projekte/Autopräsi/autopraesi_github"
PYTHON_DIR="$PROJECT_DIR/autopraesi_python"
WEB_DIR="$PROJECT_DIR/autopraesi-web"
LOG_DIR="$PROJECT_DIR/logs"
mkdir -p "$LOG_DIR"

echo "=== AutoPräsi Start $(date) ==="

# 1. Alte Prozesse beenden
kill $(lsof -ti :8000) 2>/dev/null || true
kill $(pgrep -f ngrok) 2>/dev/null || true
sleep 2

# 2. Backend starten
echo "Starting backend..."
cd "$PYTHON_DIR"
source venv/bin/activate
uvicorn api:app --host 0.0.0.0 --port 8000 > "$LOG_DIR/backend.log" 2>&1 &
BACKEND_PID=$!
echo "Backend PID: $BACKEND_PID"
sleep 3

# 3. ngrok starten
echo "Starting ngrok..."
ngrok http 8000 --log=stdout > "$LOG_DIR/ngrok.log" 2>&1 &
NGROK_PID=$!
echo "ngrok PID: $NGROK_PID"
sleep 5

# 4. ngrok URL auslesen
NGROK_URL=$(curl -s http://localhost:4040/api/tunnels | python3 -c "import sys,json; d=json.load(sys.stdin); print(d['tunnels'][0]['public_url'])" 2>/dev/null)

if [ -z "$NGROK_URL" ]; then
    echo "ERROR: ngrok URL konnte nicht gelesen werden"
    exit 1
fi

echo "ngrok URL: $NGROK_URL"

# 5. .env.local updaten und Vercel deployen
echo "NEXT_PUBLIC_API_URL=$NGROK_URL" > "$WEB_DIR/.env.local"

# api.ts Fallback updaten
python3 -c "
import re, sys
path = sys.argv[1]
url = sys.argv[2]
with open(path) as f:
    content = f.read()
content = re.sub(
    r'const API = process\.env\.NEXT_PUBLIC_API_URL \|\| \"[^\"]*\"',
    f'const API = process.env.NEXT_PUBLIC_API_URL || \"{url}\"',
    content
)
with open(path, 'w') as f:
    f.write(content)
" "$WEB_DIR/lib/api.ts" "$NGROK_URL"

echo "Deploying to Vercel..."
cd "$WEB_DIR"
vercel deploy --prod --yes > "$LOG_DIR/vercel.log" 2>&1

echo "=== AutoPräsi läuft ==="
echo "Backend: http://localhost:8000"
echo "Tunnel:  $NGROK_URL"
echo "Web:     https://autopraesi-web.vercel.app"
echo ""
echo "Logs: $LOG_DIR/"
echo "Stoppen: kill $BACKEND_PID $NGROK_PID"

# PIDs speichern zum späteren Stoppen
echo "$BACKEND_PID" > "$LOG_DIR/backend.pid"
echo "$NGROK_PID" > "$LOG_DIR/ngrok.pid"

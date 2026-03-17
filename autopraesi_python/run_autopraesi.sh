#!/bin/bash
# AutoPräsi – Wöchentlicher Lauf (via launchd)
# Erstellt die Gottesdienst-Präsentation für den kommenden Sonntag.

cd "$(dirname "$0")"
./venv/bin/python3 main.py

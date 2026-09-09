#!/bin/bash
cd "$(dirname "$0")"
echo "→ Installation des dépendances..."
pip install --break-system-packages -q -r requirements.txt
echo "→ Lancement sur http://localhost:5000"
echo "  Ctrl+C pour arrêter"
python3 app.py

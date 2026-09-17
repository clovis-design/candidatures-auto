#!/usr/bin/env bash
set -euo pipefail
ROOT="$(dirname "$(realpath "$0")")"
if [ ! -x "$ROOT/.venv/bin/python" ]; then
  echo "→ Création de l'environnement Python local..."
  python3 -m venv "$ROOT/.venv"
fi
echo "→ Vérification des dépendances..."
"$ROOT/.venv/bin/python" -m pip install -q -r "$ROOT/requirements.txt"
echo "→ Lancement sur http://${HOST:-127.0.0.1}:${PORT:-5000}"
echo "  Ctrl+C pour arrêter"
exec "$ROOT/.venv/bin/python" "$ROOT/app.py"

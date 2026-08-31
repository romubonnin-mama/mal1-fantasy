#!/bin/sh
# Démarrage du serveur admin sur Render. Ignore le code copié au build par Docker et
# clone le dépôt à neuf à chaque démarrage du conteneur : la seule source de vérité est
# GitHub (comportement voulu — un conteneur Render est jetable, jamais persistant).
set -e

if [ -z "$GITHUB_TOKEN" ]; then
  echo "ERREUR : la variable d'environnement GITHUB_TOKEN n'est pas définie." >&2
  exit 1
fi

REPO_URL="https://x-access-token:${GITHUB_TOKEN}@github.com/romubonnin-mama/mal1-fantasy.git"
CLONE_DIR="/app/live-repo"

rm -rf "$CLONE_DIR"
git clone "$REPO_URL" "$CLONE_DIR"
cd "$CLONE_DIR"

git config user.email "admin-cloud@mal1-fantasy.local"
git config user.name "Ma L1 Admin (cloud)"

exec python scripts/admin_server.py

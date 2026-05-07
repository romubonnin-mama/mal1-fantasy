"""
Gèle la saison courante dans data-archive-2025-26.json.
À lancer depuis la racine du projet après avoir calculé J34 :
    python scripts/freeze_season.py
"""
import shutil
import json
from datetime import datetime

src = 'data.json'
dst = 'data-archive-2025-26.json'

with open(src, encoding='utf-8') as f:
    d = json.load(f)

nb_journees = d.get('derniere_journee', '?')
champion = d['classement'][0]['nom'] if d.get('classement') else '?'

shutil.copy(src, dst)
print(f"[{datetime.now():%Y-%m-%d %H:%M}] Saison 2025-26 figée.")
print(f"  → {dst}")
print(f"  → {nb_journees} journées · Champion : {champion}")

"""
Télécharge les logos des clubs depuis api-sports.io vers le dossier logos/.
URL publique (sans clé API) : https://media.api-sports.io/football/teams/{id}.png

Usage : python scripts/download_logos.py
"""

import json
import sys
from pathlib import Path

try:
    import requests
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "requests", "-q"])
    import requests

BASE_DIR  = Path(__file__).parent.parent
LOGOS_DIR = BASE_DIR / "logos"
LOGOS_DIR.mkdir(exist_ok=True)

with open(BASE_DIR / "clubs.json", encoding="utf-8") as f:
    clubs = json.load(f)

club_ids = sorted(set(clubs.values()))
print(f"{len(club_ids)} clubs à télécharger vers {LOGOS_DIR}\n")

for club_id in club_ids:
    dest = LOGOS_DIR / f"{club_id}.png"
    if dest.exists():
        print(f"  {club_id}.png  déjà présent, skip")
        continue
    url = f"https://media.api-sports.io/football/teams/{club_id}.png"
    try:
        r = requests.get(url, timeout=10)
        if r.status_code == 200 and r.headers.get("content-type", "").startswith("image"):
            dest.write_bytes(r.content)
            print(f"  {club_id}.png  OK ({len(r.content)//1024} Ko)")
        else:
            print(f"  {club_id}.png  ERREUR HTTP {r.status_code}")
    except Exception as e:
        print(f"  {club_id}.png  ERREUR : {e}")

print("\nTerminé.")

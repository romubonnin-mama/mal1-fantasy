"""
Reconstruit clubs.json (mapping joueur -> id club API-Sports) depuis les effectifs
acquis aux enchères dans Firestore (collection 'managers', champ 'joueurs').

A lancer une fois les enchères ET le mercato terminés (le club de chaque joueur
enchéri peut encore bouger pendant le mercato, cf. décision utilisateur 2026-08-06).

Usage : python scripts/rebuild_clubs_from_firestore.py
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

BASE_DIR = Path(__file__).parent.parent
CLUBS_PATH = BASE_DIR / "clubs.json"

PROJECT_ID = "coach-ligue1"

# Doit rester synchronisé avec CLUB_IDS_A dans admin.html / CLUB_IDS dans enchere.html
CLUB_IDS = {
    'PSG': 85, 'PARIS': 85, 'PARIS SG': 85, 'PARIS SAINT-GERMAIN': 85,
    'OM': 81, 'MARSEILLE': 81, 'OLYMPIQUE DE MARSEILLE': 81,
    'MONACO': 91, 'ASM': 91, 'AS MONACO': 91,
    'OL': 80, 'LYON': 80, 'OLYMPIQUE LYONNAIS': 80,
    'LOSC': 79, 'LILLE': 79, 'LILLE OSC': 79,
    'NICE': 84, 'OGC NICE': 84, 'OGCN': 84,
    'LENS': 116, 'RCL': 116, 'RC LENS': 116,
    'RENNES': 94, 'SRFC': 94, 'STADE RENNAIS': 94,
    'STRASBOURG': 95, 'RCSA': 95, 'RC STRASBOURG': 95,
    'AUXERRE': 108, 'AJA': 108, 'AJ AUXERRE': 108,
    'LE HAVRE': 111, 'HAC': 111, 'HAVRE': 111,
    'TOULOUSE': 96, 'TFC': 96, 'FC TOULOUSE': 96,
    'NANTES': 83, 'FCN': 83, 'FC NANTES': 83,
    'SAINT-ETIENNE': 97, 'ASSE': 97, 'ST ETIENNE': 97, 'AS SAINT-ETIENNE': 97,
    'ANGERS': 77, 'SCO': 77, 'ANGERS SCO': 77,
    'BREST': 106, 'SB29': 106, 'STADE BRESTOIS': 106,
    'METZ': 112, 'FC METZ': 112,
    'LORIENT': 97, 'FCL': 97, 'FC LORIENT': 97,
    'PFC': 114, 'PARIS FC': 114,
    'TROYES': 110, 'ESTAC': 110, 'ESTAC TROYES': 110,
    'LE MANS': 1298, 'LMFC': 1298, 'LE MANS FC': 1298,
}


def _fs_value(v: dict):
    if "stringValue" in v:
        return v["stringValue"]
    if "integerValue" in v:
        return int(v["integerValue"])
    if "doubleValue" in v:
        return v["doubleValue"]
    if "booleanValue" in v:
        return v["booleanValue"]
    if "nullValue" in v:
        return None
    if "arrayValue" in v:
        return [_fs_value(x) for x in v["arrayValue"].get("values", [])]
    if "mapValue" in v:
        return _fs_fields(v["mapValue"].get("fields", {}))
    return None


def _fs_fields(fields: dict) -> dict:
    return {k: _fs_value(v) for k, v in fields.items()}


def fetch_managers() -> dict:
    url = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/(default)/documents/managers"
    managers = {}
    page_token = None
    while True:
        params = {"pageToken": page_token} if page_token else {}
        r = requests.get(url, params=params, timeout=15)
        r.raise_for_status()
        payload = r.json()
        for doc in payload.get("documents", []):
            nom = doc["name"].rsplit("/", 1)[-1]
            managers[nom] = _fs_fields(doc.get("fields", {}))
        page_token = payload.get("nextPageToken")
        if not page_token:
            break
    return managers


def main():
    print(f"Lecture Firestore ({PROJECT_ID}/managers)...")
    managers = fetch_managers()
    if not managers:
        print("Aucun manager trouvé, abandon.")
        return

    clubs = {}
    conflicts = []
    unknown_clubs = set()

    for mgr, data in managers.items():
        for j in data.get("joueurs") or []:
            joueur_nom = (j.get("nom") or "").strip().upper()
            club_nom = (j.get("club") or "").strip().upper()
            if not joueur_nom or not club_nom:
                continue
            club_id = CLUB_IDS.get(club_nom)
            if club_id is None:
                unknown_clubs.add(club_nom)
                continue
            club_id = str(club_id)
            if joueur_nom in clubs and clubs[joueur_nom] != club_id:
                conflicts.append((joueur_nom, clubs[joueur_nom], club_id, mgr))
            else:
                clubs[joueur_nom] = club_id

    with open(CLUBS_PATH, "w", encoding="utf-8") as f:
        json.dump(clubs, f, ensure_ascii=False, indent=2)
        f.write("\n")

    print(f"\n{CLUBS_PATH} mis à jour : {len(clubs)} joueurs")

    if unknown_clubs:
        print("\n⚠️  Clubs non reconnus (à ajouter dans CLUB_IDS) — joueurs ignorés :")
        for c in sorted(unknown_clubs):
            print(f"   - {c}")

    if conflicts:
        print("\n⚠️  Conflits club détectés pour un même joueur (dernière valeur ignorée, vérifier manuellement) :")
        for nom, old, new, mgr in conflicts:
            print(f"   - {nom}: {old} vs {new} (vu chez {mgr})")


if __name__ == "__main__":
    main()

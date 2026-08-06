"""
Synchronise data/roster.json depuis les résultats d'enchères validés dans Firestore
(collection 'managers', champ 'joueurs' rempli par "Compiler le tour" dans admin.html).

L'Excel n'a rien à faire manuellement : export_excel.py lit roster.json et remplit
la grille automatiquement dès qu'une journée est calculée. Ce script est donc le seul
maillon manquant entre "un tour est compilé" et "l'effectif est à jour partout".

Usage : python scripts/sync_roster_from_firestore.py
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
DATA_DIR = BASE_DIR / "data"
ROSTER_PATH = DATA_DIR / "roster.json"

PROJECT_ID = "coach-ligue1"
POSTE_ORDER = ["G", "D", "M", "A"]


def _fs_value(v: dict):
    """Convertit une valeur typée Firestore REST en valeur Python native."""
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


def build_roster(managers: dict) -> dict:
    roster = {}
    for nom, data in managers.items():
        joueurs = data.get("joueurs") or []
        buckets = {p: [] for p in POSTE_ORDER}
        seen = set()
        for j in joueurs:
            poste = j.get("poste")
            joueur_nom = (j.get("nom") or "").strip()
            club = (j.get("club") or "").strip().lower()
            if poste not in buckets or not joueur_nom:
                continue
            key = (joueur_nom.lower(), club)
            if key in seen:
                continue
            seen.add(key)
            buckets[poste].append(joueur_nom)
        roster[nom] = buckets
    return roster


def main():
    print(f"Lecture Firestore ({PROJECT_ID}/managers)...")
    managers = fetch_managers()
    if not managers:
        print("Aucun manager trouvé, abandon.")
        return

    roster = build_roster(managers)

    with open(ROSTER_PATH, "w", encoding="utf-8") as f:
        json.dump(roster, f, ensure_ascii=False, indent=2)
        f.write("\n")

    print(f"\n{ROSTER_PATH} mis à jour :\n")
    for nom in sorted(roster):
        counts = " / ".join(f"{p}:{len(roster[nom][p])}" for p in POSTE_ORDER)
        total = sum(len(roster[nom][p]) for p in POSTE_ORDER)
        budget_restant = managers[nom].get("budget_restant")
        print(f"  {nom:<10} {total:2d} joueurs ({counts})  budget restant: {budget_restant}")


if __name__ == "__main__":
    main()

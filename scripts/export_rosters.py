"""
Exporte les compositions d'équipe de chaque manager (joueur, poste, club)
depuis data/roster.json + clubs.json.

Usage:
  python scripts/export_rosters.py

Génère :
  data/rosters_export.csv   (Manager, Poste, Joueur, Club) — à ouvrir dans Excel/Sheets
  data/rosters_export.json  (même contenu, structuré par manager/poste)
"""

import csv
import json
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "data"

POSTES = ["G", "D", "M", "A"]

# ID club (numérique, cf. logos/<id>.png) -> nom lisible.
# Vérifié visuellement contre les logos du dossier logos/ (les alias de
# CLUB_IDS_A dans admin.html contiennent une erreur : l'id 97 y est mappé à
# la fois à LORIENT et à SAINT-ETIENNE, alors que le logo confirme id 97 = Lorient).
CLUB_NAMES = {
    "77":   "Angers SCO",
    "79":   "LOSC Lille",
    "80":   "Olympique Lyonnais",
    "81":   "Olympique de Marseille",
    "83":   "FC Nantes",
    "84":   "OGC Nice",
    "85":   "Paris Saint-Germain",
    "91":   "AS Monaco",
    "94":   "Stade Rennais",
    "95":   "RC Strasbourg",
    "96":   "Toulouse FC",
    "97":   "FC Lorient",
    "106":  "Stade Brestois",
    "108":  "AJ Auxerre",
    "110":  "ESTAC Troyes",
    "111":  "Le Havre AC",
    "112":  "FC Metz",
    "114":  "Paris FC",
    "116":  "RC Lens",
    "1298": "Le Mans FC",
}


def main():
    with open(DATA_DIR / "roster.json", encoding="utf-8") as f:
        roster = json.load(f)
    with open(BASE_DIR / "clubs.json", encoding="utf-8") as f:
        clubs = json.load(f)

    rows = []
    export = {}
    for manager, postes in roster.items():
        export[manager] = {}
        for poste in POSTES:
            joueurs = postes.get(poste, [])
            export[manager][poste] = []
            for nom in joueurs:
                club_id = str(clubs.get(nom, "")).strip()
                club_nom = CLUB_NAMES.get(club_id, f"Club inconnu (id {club_id})" if club_id else "Club inconnu")
                rows.append({"manager": manager, "poste": poste, "nom": nom, "club": club_nom})
                export[manager][poste].append({"nom": nom, "club": club_nom})

    csv_path = DATA_DIR / "rosters_export.csv"
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["manager", "poste", "nom", "club"],
                                 extrasaction="ignore")
        writer.writerow({"manager": "Manager", "poste": "Poste", "nom": "Joueur", "club": "Club"})
        for row in rows:
            writer.writerow(row)

    json_path = DATA_DIR / "rosters_export.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(export, f, ensure_ascii=False, indent=2)

    inconnus = sorted({r["nom"] for r in rows if "inconnu" in r["club"].lower()})
    print(f"✅ {len(rows)} joueurs exportés pour {len(roster)} managers.")
    print(f"  → {csv_path}")
    print(f"  → {json_path}")
    if inconnus:
        print(f"⚠️  Club introuvable pour : {', '.join(inconnus)}")


if __name__ == "__main__":
    main()

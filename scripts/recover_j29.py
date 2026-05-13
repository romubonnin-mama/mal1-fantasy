"""
Récupère les stats J29 depuis le commit git 867a043 (état correct avant écrasement)
et reconstruit manual_stats.json["29"], en conservant les stats déjà saisies pour
les joueurs des matchs en retard (présents dans le manual_stats.json actuel).

Usage : python scripts/recover_j29.py
"""

import json
import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "data"

GOOD_COMMIT = "867a043"
JOURNEE = "29"


def load_detail_from_git(commit: str) -> dict:
    r = subprocess.run(
        ["git", "show", f"{commit}:data.json"],
        cwd=BASE_DIR, capture_output=True, text=True, encoding="utf-8"
    )
    if r.returncode != 0:
        print(f"ERREUR git show : {r.stderr}")
        sys.exit(1)
    d = json.loads(r.stdout)
    detail = d.get("detail_journees", {}).get(JOURNEE)
    if not detail:
        print(f"Aucune donnée J{JOURNEE} dans {commit}:data.json")
        sys.exit(1)
    return detail


def reconstruct_manual_stats(detail: dict) -> dict:
    """
    Convertit le format detail_journees → manual_stats pour tous les titulaires.
    """
    result = {}
    for manager, postes in detail.items():
        m_stats = {}
        for poste, players in postes.items():
            for p in players:
                nom    = p["nom"]
                statut = p.get("statut", "")

                if statut == "r":
                    continue  # remplaçant → pas dans manual_stats

                if statut == "A":
                    m_stats[nom] = {"absent": True}
                    continue

                # Titulaire : reconstruction depuis les champs calculés
                tj   = p.get("tj", "0")
                bm   = (p.get("bm") or {}).get("val", 0) or 0
                pd   = (p.get("pd") or {}).get("val", 0) or 0
                cs   = (p.get("cs") or {}).get("val", 0) or 0
                be   = (p.get("be") or {}).get("val", 0) or 0  # nb buts encaissés
                bcsc = (p.get("bcsc") or {}).get("val", 0) or 0
                pm   = (p.get("pm") or {}).get("val", 0) or 0
                pma  = (p.get("pma") or {}).get("val", 0) or 0
                cj   = (p.get("cj") or {}).get("val", 0) or 0
                cr   = (p.get("cr") or {}).get("val", 0) or 0  # -1 si carton rouge

                s = {}

                tj_str = str(tj)
                if tj_str == "M":
                    s["full_match"] = True
                elif "-" in tj_str:
                    # format "entre-sort" ex: "70-90"
                    parts = tj_str.split("-")
                    try:
                        s["entre_a"] = int(parts[0])
                        s["fin_a"]   = int(parts[1])
                    except ValueError:
                        pass
                elif tj_str != "0":
                    try:
                        v = int(tj_str)
                        if v > 0:
                            s["sort_a"] = v
                    except ValueError:
                        pass
                # tj == "0" → aucune clé temps (0 minutes joués)

                if bm   > 0: s["goals"]       = bm
                if pd   > 0: s["assists"]      = pd
                if cs   > 0: s["cs"]           = True
                if be   >= 3: s["be_malus"]    = True
                if bcsc > 0: s["own_goals"]    = bcsc
                if pm   > 0: s["pen_scored"]   = pm
                if pma  > 0: s["pen_mm_saved"] = pma
                if cj   > 0: s["yellow_cards"] = cj
                if cr   < 0: s["red_card"]     = True

                m_stats[nom] = s

        if m_stats:
            result[manager] = m_stats
    return result


def main():
    print(f"Lecture de data.json au commit {GOOD_COMMIT}...")
    detail = load_detail_from_git(GOOD_COMMIT)
    print(f"  Managers trouvés : {list(detail.keys())}")

    print("Reconstruction manual_stats J29...")
    reconstructed = reconstruct_manual_stats(detail)

    # Charger l'état actuel pour récupérer les stats des matchs en retard
    ms_path = DATA_DIR / "manual_stats.json"
    with open(ms_path, encoding="utf-8") as f:
        manual_stats = json.load(f)

    current_j29 = manual_stats.get(JOURNEE, {})

    # Fusion : reconstruits comme base, stats actuelles (matchs en retard) par-dessus
    merged = {}
    all_managers = set(list(reconstructed.keys()) + list(current_j29.keys()))
    for manager in all_managers:
        base    = reconstructed.get(manager, {})
        overlay = current_j29.get(manager, {})
        merged[manager] = {**base, **overlay}  # overlay écrase uniquement les joueurs communs

    manual_stats[JOURNEE] = merged
    with open(ms_path, "w", encoding="utf-8") as f:
        json.dump(manual_stats, f, ensure_ascii=False, indent=2)

    print(f"\nmanual_stats.json J{JOURNEE} reconstruit :")
    for manager, players in merged.items():
        nb_overlay = len([n for n in players if n in current_j29.get(manager, {})])
        print(f"  {manager} : {len(players)} joueurs ({nb_overlay} depuis matchs en retard)")

    print("\nEtape suivante :")
    print("  1. Lance l'admin  →  python scripts/admin_server.py")
    print("  2. Stats J  →  J29  →  verifie que les stats sont bien la")
    print("  3. Ajoute/complete les joueurs des matchs de ce soir si besoin")
    print("  4. Calculer  →  Publier GitHub")


if __name__ == "__main__":
    main()

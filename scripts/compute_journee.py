"""
Calcule les points d'une journée depuis les stats manuelles et met à jour data.json.
Appelé par admin_server.py via POST /api/compute/<journee>.
"""

import json
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "data"
sys.path.insert(0, str(Path(__file__).parent))

from scoring import calcul_joueur, appliquer_capitaine

POSTES = ["G", "D", "M", "A"]


def compute(journee: int) -> dict:
    with open(DATA_DIR / "roster.json",       encoding="utf-8") as f: roster      = json.load(f)
    with open(DATA_DIR / "lineups.json",      encoding="utf-8") as f: lineups     = json.load(f)
    with open(DATA_DIR / "manual_stats.json", encoding="utf-8") as f: manual      = json.load(f)
    with open(DATA_DIR / "corrections.json",  encoding="utf-8") as f: corrections = json.load(f)
    with open(BASE_DIR / "data.json",         encoding="utf-8") as f: data        = json.load(f)

    j_lineups    = lineups.get(str(journee), {})
    j_manual     = manual.get(str(journee), {})
    j_corrections = corrections.get(str(journee), {})

    if not j_lineups:
        raise ValueError(f"Aucune composition définie pour J{journee}.")

    detail_journee = {}
    scores_journee = {}

    for manager, postes in roster.items():
        lineup    = j_lineups.get(manager, {})
        titulaires = set(lineup.get("titulaires", []))
        capitaine  = lineup.get("capitaine")
        coeff      = int(lineup.get("coeff", 1))

        equipe_result = {p: [] for p in POSTES}
        total = 0

        for poste in POSTES:
            for nom in postes.get(poste, []):
                is_titu = nom in titulaires
                s = j_manual.get(manager, {}).get(nom, {})

                minutes    = int(s.get("minutes", 0) or 0)
                full_match = bool(s.get("full_match", False))
                red_card   = bool(s.get("red_card", False))

                player_corrections = j_corrections.get(manager, {}).get(nom, {})
                result = calcul_joueur(
                    poste       = poste,
                    minutes     = minutes,
                    is_full_match = full_match,
                    goals_scored  = int(s.get("goals", 0)),
                    assists       = int(s.get("assists", 0)),
                    goals_conceded = int(s.get("goals_conceded", 0)),
                    penalties_scored = int(s.get("pen_scored", 0)),
                    penalties_missed          = int(s.get("pen_mm_saved", 0)) if poste != "G" else 0,
                    penalties_saved_or_opp_missed = int(s.get("pen_mm_saved", 0)) if poste == "G" else 0,
                    own_goals     = int(s.get("own_goals", 0)),
                    yellow_cards  = int(s.get("yellow_cards", 0)),
                    red_card      = red_card,
                    corrections   = player_corrections,
                )

                cap_str = ""
                if is_titu and nom == capitaine:
                    pts_cap = appliquer_capitaine(result["pts"], coeff)
                    cap_str = str(coeff)
                    result["pts"] = pts_cap
                elif not is_titu:
                    result["pts"] = 0

                if is_titu:
                    total += result["pts"]

                equipe_result[poste].append({
                    "nom":    nom,
                    "statut": "" if is_titu else "r",
                    "cap":    cap_str,
                    "tj":     result["tj"],    "tj_pts": result["tj_pts"],
                    "bm":     result["bm"],    "be":     result["be"],
                    "bcsc":   result["bcsc"],  "cs":     result["cs"],
                    "pm":     result["pm"],    "pma":    result["pma"],
                    "pd":     result["pd"],    "cj":     result["cj"],
                    "cr":     result["cr"],    "pts":    result["pts"],
                })

        detail_journee[manager] = equipe_result
        scores_journee[manager] = total

    # Mettre à jour data.json
    data["historique"][str(journee)]      = scores_journee
    data["detail_journees"][str(journee)] = detail_journee
    data["scores_journee"]                = scores_journee
    data["derniere_journee"]              = max(data.get("derniere_journee", 0), journee)

    noms   = [j["nom"] for j in data["classement"]]
    cumul  = {n: 0 for n in noms}
    evo    = {n: [] for n in noms}
    for jj in range(1, data["derniere_journee"] + 1):
        if str(jj) in data["historique"]:
            for n in noms:
                cumul[n] += data["historique"][str(jj)].get(n, 0)
            for n in noms:
                evo[n].append({"j": jj, "pts": cumul[n]})
    data["evolution"] = evo
    data["classement"] = sorted(
        [{"rang": 0, "nom": n, "pts": cumul[n]} for n in noms],
        key=lambda x: -x["pts"],
    )
    for i, j in enumerate(data["classement"]):
        j["rang"] = i + 1

    with open(BASE_DIR / "data.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return {"ok": True, "scores": scores_journee, "classement": data["classement"]}


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python compute_journee.py <journee>")
        sys.exit(1)
    r = compute(int(sys.argv[1]))
    print("Scores:", r["scores"])
    print("Classement:", [(j["rang"], j["nom"], j["pts"]) for j in r["classement"]])

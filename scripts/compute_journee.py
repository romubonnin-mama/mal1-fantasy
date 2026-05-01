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

from scoring import calcul_joueur, appliquer_capitaine, CS_PTS

POSTES = ["G", "D", "M", "A"]

from scoring import BM_PTS, PMA_PTS, CS_PTS, BE_PTS


def _stat_pts(stat: str, poste: str, old_val: int, new_val: int, player: dict):
    """Retourne (old_pts, new_pts) pour un changement de valeur brute d'une stat."""
    if stat == "bm":
        return old_val * BM_PTS.get(poste, 0), new_val * BM_PTS.get(poste, 0)
    if stat == "pd":
        return old_val * 2, new_val * 2
    if stat == "pm":
        return old_val * 2, new_val * 2
    if stat == "pma":
        return old_val * PMA_PTS.get(poste, 0), new_val * PMA_PTS.get(poste, 0)
    if stat == "bcsc":
        return old_val * (-2), new_val * (-2)
    if stat == "cs":
        return old_val * CS_PTS.get(poste, 0), new_val * CS_PTS.get(poste, 0)
    if stat == "be":
        be = BE_PTS.get(poste, 0)
        return (be if old_val >= 3 else 0), (be if new_val >= 3 else 0)
    if stat == "cj":
        red_card = (player.get("cr", {}).get("val", 0) != 0)
        return (0, 0) if red_card else (old_val * (-1), new_val * (-1))
    return 0, 0


def _apply_corrections_past(journee: int, j_corrections: dict, data: dict) -> dict:
    """
    Applique des corrections à une journée passée (sans lineup disponible).
    Utilise le detail_journees existant dans data.json comme base.
    Vide l'entrée corrections.json de la journée après application.
    """
    detail_journee = data["detail_journees"].get(str(journee))
    if not detail_journee:
        raise ValueError(f"Aucune donnée existante pour J{journee} dans data.json.")

    scores_journee = {}

    for manager, equipe in detail_journee.items():
        manager_corrections = j_corrections.get(manager, {})
        total = 0

        for poste, players in equipe.items():
            for player in players:
                nom      = player["nom"]
                is_titu  = (player.get("statut", "") != "r")
                cap      = player.get("cap", "")
                coeff    = int(cap) if cap and str(cap).isdigit() else 0
                multiplier = (1 + coeff) if (cap and is_titu) else 1

                if is_titu and nom in manager_corrections:
                    nom_corr = manager_corrections[nom]

                    # ABS : force le joueur absent, ignore les autres corrections
                    if int((nom_corr.get("abs") or {}).get("val", 0) or 0):
                        player["pts"] = 0
                        player["statut"] = "A"
                    else:
                        # full_match : tj passe à "M" (4 pts)
                        if int((nom_corr.get("full_match") or {}).get("val", 0) or 0):
                            tj_entry = player.get("tj_pts")
                            old_tj_pts = tj_entry.get("pts", 0) if isinstance(tj_entry, dict) else 0
                            player["tj"]     = "M"
                            player["tj_pts"] = {"val": "M", "pts": 4}
                            player["pts"]   += (4 - old_tj_pts) * multiplier

                        for stat, corr in nom_corr.items():
                            if stat in ("abs", "full_match"):
                                continue
                            delta_val = int((corr.get("val") or 0) or 0)
                            if not delta_val or stat not in player:
                                continue
                            old_val = player[stat]["val"] if isinstance(player[stat], dict) else 0
                            new_val = old_val + delta_val
                            old_pts, new_pts = _stat_pts(stat, poste, old_val, new_val, player)
                            player[stat] = {"val": new_val, "pts": new_pts}
                            player["pts"] += (new_pts - old_pts) * multiplier

                if is_titu:
                    total += player["pts"]

        scores_journee[manager] = total

    # Mettre à jour data.json
    data["historique"][str(journee)]      = scores_journee
    data["detail_journees"][str(journee)] = detail_journee
    data["scores_journee"]                = scores_journee
    data["derniere_journee"]              = max(data.get("derniere_journee", 0), journee)

    noms  = [j["nom"] for j in data["classement"]]
    cumul = {n: 0 for n in noms}
    evo   = {n: [] for n in noms}
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
    for i, entry in enumerate(data["classement"]):
        entry["rang"] = i + 1

    with open(BASE_DIR / "data.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    # Vider l'entrée corrections.json pour cette journée (déjà appliquée)
    corrections_path = DATA_DIR / "corrections.json"
    with open(corrections_path, encoding="utf-8") as f:
        all_corrections = json.load(f)
    all_corrections.pop(str(journee), None)
    with open(corrections_path, "w", encoding="utf-8") as f:
        json.dump(all_corrections, f, ensure_ascii=False, indent=2)

    return {"ok": True, "scores": scores_journee, "classement": data["classement"]}


def _minutes(s: dict) -> int:
    sort_a  = int(s.get("sort_a",  0) or 0)
    entre_a = int(s.get("entre_a", 0) or 0)
    fin_a   = int(s.get("fin_a",   0) or 0)
    if entre_a > 0:
        end = sort_a if sort_a > 0 else (fin_a if fin_a > 0 else 90)
        return end - entre_a
    if sort_a > 0:
        return sort_a
    return int(s.get("minutes", 0) or 0)


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
        if not j_corrections:
            raise ValueError(f"Aucune composition définie pour J{journee}. Ajoutez des corrections d'abord.")
        return _apply_corrections_past(journee, j_corrections, data)

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

                absent     = bool(s.get("absent", False))
                minutes    = _minutes(s)
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

                # Absent : 0 point, aucun bonus
                if absent:
                    result["pts"] = 0

                # Override CS manuel (case à cocher dans l'admin)
                if s.get("cs") and not red_card and not absent:
                    cs_pts = CS_PTS.get(poste, 0)
                    if cs_pts != result["cs"]["pts"]:
                        result["pts"] += cs_pts - result["cs"]["pts"]
                        result["cs"] = {"val": 1, "pts": cs_pts}

                cap_str = ""
                if is_titu and not absent and nom == capitaine:
                    pts_cap = appliquer_capitaine(result["pts"], coeff)
                    cap_str = str(coeff)
                    result["pts"] = pts_cap
                elif not is_titu:
                    result["pts"] = 0

                if is_titu:
                    total += result["pts"]

                equipe_result[poste].append({
                    "nom":    nom,
                    "statut": "A" if (is_titu and absent) else ("" if is_titu else "r"),
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

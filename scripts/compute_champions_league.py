"""
Calcule les résultats de la Champions League Ma L1 (compétition parallèle au championnat
principal, voir Champions League.docx) à partir des données déjà calculées pour le
championnat (data.json::detail_journees). Appelé par admin_server.py via
POST /api/compute-cl/<journee>.

Règles clés (voir Champions League.docx) :
- Mêmes compositions/stats que le championnat, mais le bonus capitaine ne compte pas.
- Phase de poule (J4,J6,J8,J10,J12,J14,J15) : round-robin à 8, 3/1/0 pts.
- Phase finale (quarts J23/J26, demies J29/J32, finale J34) : aller-retour sauf la finale.
- Règle de report : gérée manuellement par l'admin (case à cocher par joueur/journée dans
  report_overrides) — substitue la moyenne de points MaL1 (sans capitaine) du joueur sur la
  saison en cours, arrondie. Pas de détection automatique (pas de données de calendrier L1
  réel dans ce projet).
"""

import json
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent
DATA_PATH = BASE_DIR / "data.json"
CL_PATH   = BASE_DIR / "champions_league.json"

POULE_JOURNEES = (4, 6, 8, 10, 12, 14, 15)

# clé -> (poids) : True = le plus grand gagne, False = le plus petit gagne
CRITERES_DEPARTAGE = (("bm", True), ("pd", True), ("cs", True), ("cr", False), ("cj", False))


# ─── Points "sans capitaine" ────────────────────────────────────────────────────

def pts_sans_capitaine_joueur(p: dict) -> int:
    """
    Points d'un joueur sans le bonus capitaine. Les champs composants (bm, be, bcsc, cs,
    pm, pma, pd, cj, cr, tj_pts) ne sont JAMAIS multipliés par le coefficient capitaine
    dans compute_journee.py (seul le total `pts` l'est, uniquement pour le capitaine) :
    donc pour un non-capitaine, `pts` est déjà la bonne valeur ; pour le capitaine, on
    reconstruit le total brut depuis les composantes.
    """
    if not p.get("cap"):
        return p.get("pts", 0) or 0
    tj = p.get("tj_pts", 0)
    tj_pts = tj if isinstance(tj, int) else (tj.get("pts", 0) if isinstance(tj, dict) else 0)
    total = tj_pts
    for k in ("bm", "be", "bcsc", "cs", "pm", "pma", "pd", "cj", "cr"):
        v = p.get(k)
        if isinstance(v, dict):
            total += v.get("pts", 0) or 0
    return total


def historique_joueur(nom: str, data: dict, avant_journee: int) -> list:
    """Liste des points (sans capitaine) du joueur `nom` sur les journées < avant_journee
    déjà calculées, tous managers confondus (peu importe qui le possédait), titulaire
    uniquement (statut != 'r')."""
    pts_list = []
    for j_str, par_manager in data.get("detail_journees", {}).items():
        if int(j_str) >= avant_journee:
            continue
        for manager, postes in par_manager.items():
            for poste, players in postes.items():
                for p in players:
                    if p.get("nom") == nom and p.get("statut") != "r":
                        pts_list.append(pts_sans_capitaine_joueur(p))
    return pts_list


def moyenne_joueur(nom: str, data: dict, avant_journee: int) -> int:
    """Moyenne (arrondie) de points MaL1 (sans capitaine) du joueur sur la saison en
    cours avant `avant_journee`. 0 si aucun match disputé (règle de report)."""
    pts_list = historique_joueur(nom, data, avant_journee)
    if not pts_list:
        return 0
    return round(sum(pts_list) / len(pts_list))


def score_manager_cl(manager: str, journee: int, data: dict, report_players: set) -> int:
    """Score CL d'un manager pour une journée : somme des titulaires, sans bonus
    capitaine, avec substitution moyenne pour les joueurs signalés en report."""
    equipe = data.get("detail_journees", {}).get(str(journee), {}).get(manager, {})
    total = 0
    for poste, players in equipe.items():
        for p in players:
            if p.get("statut") == "r":
                continue
            nom = p.get("nom")
            if nom in report_players:
                total += moyenne_joueur(nom, data, journee)
            else:
                total += pts_sans_capitaine_joueur(p)
    return total


def _stats_reelles(manager: str, journee: int, data: dict) -> dict:
    """Stats réelles (buts, passes, clean sheets, cartons) des titulaires d'un manager
    pour une journée — utilisées pour les critères de départage."""
    equipe = data.get("detail_journees", {}).get(str(journee), {}).get(manager, {})
    out = {"bm": 0, "pd": 0, "cs": 0, "cr": 0, "cj": 0}
    for poste, players in equipe.items():
        for p in players:
            if p.get("statut") == "r":
                continue
            out["bm"] += (p.get("bm") or {}).get("val", 0) or 0
            out["pd"] += (p.get("pd") or {}).get("val", 0) or 0
            out["cs"] += 1 if (p.get("cs") or {}).get("val", 0) else 0
            out["cr"] += 1 if (p.get("cr") or {}).get("val", 0) else 0
            out["cj"] += (p.get("cj") or {}).get("val", 0) or 0
    return out


def _departage(a: str, b: str, pts_a: int, pts_b: int, data: dict, cl_data: dict, journees: list):
    """Vainqueur d'un H2H de phase finale à égalité : critères successifs, puis meilleur
    classement de poule. Retourne None si égalité totale (cas limite non tranché)."""
    if pts_a > pts_b:
        return a
    if pts_b > pts_a:
        return b

    stat_a = {"bm": 0, "pd": 0, "cs": 0, "cr": 0, "cj": 0}
    stat_b = dict(stat_a)
    for j in journees:
        for k, v in _stats_reelles(a, j, data).items():
            stat_a[k] += v
        for k, v in _stats_reelles(b, j, data).items():
            stat_b[k] += v

    for key, higher_better in CRITERES_DEPARTAGE:
        va, vb = stat_a[key], stat_b[key]
        if va == vb:
            continue
        if higher_better:
            return a if va > vb else b
        return a if va < vb else b

    rang = {c["nom"]: i for i, c in enumerate(cl_data.get("classement_poule", []))}
    ra, rb = rang.get(a, 999), rang.get(b, 999)
    if ra != rb:
        return a if ra < rb else b
    return None


# ─── Phase de poule ──────────────────────────────────────────────────────────────

def _recompute_classement_poule(cl_data: dict, data: dict) -> None:
    participants = cl_data["participants"]
    table = {p: {"nom": p, "pts": 0, "v": 0, "n": 0, "d": 0, "diff_h2h": 0,
                 "bm": 0, "pd": 0, "cs": 0, "cr": 0, "cj": 0} for p in participants}

    for j_str, res in cl_data.get("resultats_poule", {}).items():
        journee = int(j_str)
        for m in res["matches"]:
            a, b = m["a"], m["b"]
            diff = m["pts_a"] - m["pts_b"]
            if m["vainqueur"] == a:
                table[a]["pts"] += 3
                table[a]["v"]   += 1
                table[b]["d"]   += 1
            elif m["vainqueur"] == b:
                table[b]["pts"] += 3
                table[b]["v"]   += 1
                table[a]["d"]   += 1
            else:
                table[a]["pts"] += 1
                table[b]["pts"] += 1
                table[a]["n"]   += 1
                table[b]["n"]   += 1
            table[a]["diff_h2h"] += diff
            table[b]["diff_h2h"] -= diff

            for k, v in _stats_reelles(a, journee, data).items():
                table[a][k] += v
            for k, v in _stats_reelles(b, journee, data).items():
                table[b][k] += v

    classement_actuel = {c["nom"]: c["rang"] for c in data.get("classement", [])}

    def sort_key(p):
        t = table[p]
        return (-t["pts"], -t["diff_h2h"], -t["bm"], -t["pd"], -t["cs"],
                t["cr"], t["cj"], classement_actuel.get(p, 999))

    ranked = sorted(participants, key=sort_key)
    cl_data["classement_poule"] = [table[p] for p in ranked]


def _seed_bracket(cl_data: dict) -> None:
    ranked = [c["nom"] for c in cl_data["classement_poule"]]
    pairs = {"A": (ranked[0], ranked[7]), "B": (ranked[3], ranked[4]),
             "C": (ranked[1], ranked[6]), "D": (ranked[2], ranked[5])}
    for k, (a, b) in pairs.items():
        cl_data["bracket"]["quarts"][k]["paire"] = [a, b]
    cl_data["phase"] = "quarts"


def _compute_poule_journee(journee: int, data: dict, cl_data: dict) -> dict:
    pairs = cl_data.get("calendrier_poule", {}).get(str(journee))
    if not pairs:
        raise ValueError(f"Calendrier de la phase de poule non renseigné pour J{journee}.")
    if str(journee) not in data.get("detail_journees", {}):
        raise ValueError(f"La journée MaL1 J{journee} n'a pas encore été calculée.")

    report_players = set(cl_data.get("report_overrides", {}).get(str(journee), []))
    matches = []
    for a, b in pairs:
        pts_a = score_manager_cl(a, journee, data, report_players)
        pts_b = score_manager_cl(b, journee, data, report_players)
        if pts_a > pts_b:
            vainqueur = a
        elif pts_b > pts_a:
            vainqueur = b
        else:
            vainqueur = None
        matches.append({"a": a, "b": b, "pts_a": pts_a, "pts_b": pts_b, "vainqueur": vainqueur})

    cl_data["resultats_poule"][str(journee)] = {"matches": matches}
    _recompute_classement_poule(cl_data, data)

    if cl_data["phase"] == "poule" and len(cl_data["resultats_poule"]) == len(POULE_JOURNEES):
        _seed_bracket(cl_data)

    return {"ok": True, "phase": journee, "matches": matches, "classement_poule": cl_data["classement_poule"]}


# ─── Phase finale ────────────────────────────────────────────────────────────────

def _find_bracket_match(cl_data: dict, journee: int):
    b = cl_data["bracket"]
    for k, m in b["quarts"].items():
        if m["aller"]["j"] == journee:
            return "quarts", k, "aller", m
        if m["retour"]["j"] == journee:
            return "quarts", k, "retour", m
    for k, m in b["demies"].items():
        if m["aller"]["j"] == journee:
            return "demies", k, "aller", m
        if m["retour"]["j"] == journee:
            return "demies", k, "retour", m
    if b["finale"]["j"] == journee:
        return "finale", None, None, b["finale"]
    return None


def _avancer_tour(cl_data: dict, round_name: str, match_id: str, vainqueur: str) -> None:
    if vainqueur is None:
        return
    if round_name == "quarts":
        mapping = {"A": ("1", 0), "B": ("1", 1), "C": ("2", 0), "D": ("2", 1)}
        demi_id, idx = mapping[match_id]
        cl_data["bracket"]["demies"][demi_id]["paire"][idx] = vainqueur
    elif round_name == "demies":
        idx = 0 if match_id == "1" else 1
        cl_data["bracket"]["finale"]["paire"][idx] = vainqueur


def _compute_bracket_journee(journee: int, data: dict, cl_data: dict) -> dict:
    found = _find_bracket_match(cl_data, journee)
    if not found:
        raise ValueError(f"J{journee} ne correspond à aucun match programmé de la phase finale CL.")
    round_name, match_id, leg, match = found

    a, b = match["paire"]
    if a is None or b is None:
        raise ValueError(f"La paire du match J{journee} n'est pas encore déterminée (tour précédent pas terminé).")
    if str(journee) not in data.get("detail_journees", {}):
        raise ValueError(f"La journée MaL1 J{journee} n'a pas encore été calculée.")

    report_players = set(cl_data.get("report_overrides", {}).get(str(journee), []))
    pts_a = score_manager_cl(a, journee, data, report_players)
    pts_b = score_manager_cl(b, journee, data, report_players)

    if round_name == "finale":
        match["pts_a"], match["pts_b"] = pts_a, pts_b
        vainqueur = _departage(a, b, pts_a, pts_b, data, cl_data, [journee])
        match["vainqueur"] = vainqueur
        if vainqueur is not None:
            cl_data["phase"] = "terminee"
        return {"ok": True, "round": "finale", "pts_a": pts_a, "pts_b": pts_b, "vainqueur": vainqueur}

    leg_dict = match[leg]
    leg_dict["pts_a"], leg_dict["pts_b"] = pts_a, pts_b
    vainqueur = None

    if leg == "retour":
        if match["aller"]["pts_a"] is None:
            raise ValueError(f"Le match aller (J{match['aller']['j']}) n'a pas encore été calculé pour ce match.")
        agg_a = match["aller"]["pts_a"] + pts_a
        agg_b = match["aller"]["pts_b"] + pts_b
        journees = [match["aller"]["j"], match["retour"]["j"]]
        vainqueur = _departage(a, b, agg_a, agg_b, data, cl_data, journees)
        match["vainqueur"] = vainqueur
        _avancer_tour(cl_data, round_name, match_id, vainqueur)

    return {"ok": True, "round": round_name, "match": match_id, "leg": leg,
            "pts_a": pts_a, "pts_b": pts_b, "vainqueur": vainqueur}


# ─── Point d'entrée ──────────────────────────────────────────────────────────────

def compute(journee: int) -> dict:
    with open(DATA_PATH, encoding="utf-8") as f:
        data = json.load(f)
    with open(CL_PATH, encoding="utf-8") as f:
        cl_data = json.load(f)

    if journee in POULE_JOURNEES:
        result = _compute_poule_journee(journee, data, cl_data)
    else:
        result = _compute_bracket_journee(journee, data, cl_data)

    with open(CL_PATH, "w", encoding="utf-8") as f:
        json.dump(cl_data, f, ensure_ascii=False, indent=2)

    return result


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python compute_champions_league.py <journee>")
        sys.exit(1)
    r = compute(int(sys.argv[1]))
    print(json.dumps(r, ensure_ascii=False, indent=2))

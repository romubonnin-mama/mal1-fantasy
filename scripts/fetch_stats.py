"""
Récupère les stats de la journée depuis API-Football et génère data.json.

Usage:
  python scripts/fetch_stats.py <journee>        # ex: python scripts/fetch_stats.py 32
  python scripts/fetch_stats.py --build-ids      # construit le mapping player_ids.json

Prérequis:
  - Clé API dans la variable d'environnement APISPORTS_KEY
    ou dans le fichier scripts/config.json : {"api_key": "..."}
  - data/lineups.json renseigné pour la journée
  - data/player_ids.json renseigné pour les joueurs concernés
"""

import json
import os
import sys
import time
import subprocess
from pathlib import Path

try:
    import requests
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "requests", "-q"])
    import requests

# ─── Configuration ─────────────────────────────────────────────────────────────

BASE_DIR    = Path(__file__).parent.parent
DATA_DIR    = BASE_DIR / "data"
SCRIPTS_DIR = BASE_DIR / "scripts"

LEAGUE_ID = 61    # Ligue 1 McDonald's
SEASON    = 2025  # saison 2025-2026

BASE_URL = "https://v3.football.api-sports.io"


def load_api_key() -> str:
    key = os.environ.get("APISPORTS_KEY", "")
    if not key:
        cfg_path = SCRIPTS_DIR / "config.json"
        if cfg_path.exists():
            with open(cfg_path, encoding="utf-8") as f:
                key = json.load(f).get("api_key", "")
    if not key:
        print("ERREUR : clé API introuvable.")
        print("  → Crée scripts/config.json avec {\"api_key\": \"TA_CLE\"}")
        print("  → ou définis la variable d'environnement APISPORTS_KEY")
        sys.exit(1)
    return key


def api_get(endpoint: str, params: dict, headers: dict) -> list:
    url = f"{BASE_URL}/{endpoint}"
    r = requests.get(url, params=params, headers=headers, timeout=15)
    r.raise_for_status()
    data = r.json()
    if data.get("errors"):
        raise RuntimeError(f"API error: {data['errors']}")
    return data.get("response", [])


# ─── Fixtures ──────────────────────────────────────────────────────────────────

def get_fixtures(journee: int, headers: dict) -> list:
    """Retourne tous les matchs de la journée."""
    round_str = f"Regular Season - {journee}"
    return api_get("fixtures", {"league": LEAGUE_ID, "season": SEASON, "round": round_str}, headers)


def get_fixture_players(fixture_id: int, headers: dict) -> list:
    """Retourne les stats joueurs pour un match donné."""
    time.sleep(0.3)  # respect rate limit
    return api_get("fixtures/players", {"fixture": fixture_id}, headers)


# ─── Mapping nom → API-Football ID ─────────────────────────────────────────────

def search_player(name: str, headers: dict, team_id: int = None) -> list:
    params = {"search": name, "league": LEAGUE_ID, "season": SEASON}
    if team_id:
        params["team"] = team_id
    time.sleep(0.5)
    return api_get("players", params, headers)


def build_player_ids(headers: dict):
    """
    Parcourt roster.json et cherche chaque joueur sur API-Football.
    Lance un dialogue interactif pour confirmer chaque ID.
    Sauvegarde dans data/player_ids.json.
    """
    roster_path = DATA_DIR / "roster.json"
    ids_path    = DATA_DIR / "player_ids.json"

    with open(roster_path, encoding="utf-8") as f:
        roster = json.load(f)
    with open(ids_path, encoding="utf-8") as f:
        player_ids = json.load(f)

    for manager, postes in roster.items():
        for poste, joueurs in postes.items():
            for nom in joueurs:
                if nom in player_ids:
                    continue
                print(f"\n🔍 {manager} {poste} : {nom}")
                results = search_player(nom, headers)
                if not results:
                    print(f"  ⚠️  Aucun résultat pour '{nom}'")
                    player_ids[nom] = None
                    continue
                for i, entry in enumerate(results[:5]):
                    p = entry["player"]
                    team = entry["statistics"][0]["team"]["name"] if entry.get("statistics") else "?"
                    print(f"  [{i}] ID={p['id']}  {p['name']}  ({p.get('nationality','?')})  Équipe: {team}")
                choice = input("  Choix (0-4, s=skip, n=null) : ").strip().lower()
                if choice == "n":
                    player_ids[nom] = None
                elif choice == "s":
                    continue
                elif choice.isdigit() and int(choice) < len(results):
                    player_ids[nom] = results[int(choice)]["player"]["id"]
                    print(f"  ✅ {nom} → ID {player_ids[nom]}")
                with open(ids_path, "w", encoding="utf-8") as f:
                    json.dump(player_ids, f, ensure_ascii=False, indent=2)

    print("\n✅ player_ids.json mis à jour.")


# ─── Calcul des stats joueur ────────────────────────────────────────────────────

def parse_player_stats(api_stats: dict, poste: str, goals_conceded_match: int) -> dict:
    """
    Transforme la réponse API-Football en paramètres pour scoring.calcul_joueur().
    api_stats : élément de response[].statistics[]
    """
    from scoring import calcul_joueur

    games   = api_stats.get("games", {})
    goals   = api_stats.get("goals", {})
    cards   = api_stats.get("cards", {})
    penalty = api_stats.get("penalty", {})

    minutes      = games.get("minutes") or 0
    substitute   = games.get("substitute", False)
    is_full_match = (not substitute) and minutes >= 90

    goals_scored   = goals.get("total") or 0
    assists        = goals.get("assists") or 0
    own_goals      = 0  # API-Football ne donne pas les CSC directement ici
    saves          = goals.get("saves") or 0
    yellow_cards   = cards.get("yellow") or 0
    red_card       = bool(cards.get("red"))
    pen_scored     = penalty.get("scored") or 0
    pen_missed     = penalty.get("missed") or 0
    pen_saved      = penalty.get("saved") or 0

    # Pour G : pma = penalties saved
    # Pour D/M/A : pma = penalties missed by the player
    if poste == "G":
        pma_val = pen_saved
    else:
        pma_val = pen_missed

    return calcul_joueur(
        poste=poste,
        minutes=minutes,
        is_full_match=is_full_match,
        goals_scored=goals_scored,
        assists=assists,
        goals_conceded=goals_conceded_match,
        penalties_scored=pen_scored,
        penalties_missed=pen_missed,
        penalties_saved_or_opp_missed=pen_saved,
        own_goals=own_goals,
        yellow_cards=yellow_cards,
        red_card=red_card,
    )


# ─── Pipeline principal ─────────────────────────────────────────────────────────

def fetch_journee(journee: int, headers: dict):
    """Récupère les stats de la journée et met à jour data.json."""
    from scoring import appliquer_capitaine

    ids_path      = DATA_DIR / "player_ids.json"
    lineups_path  = DATA_DIR / "lineups.json"
    corr_path     = DATA_DIR / "corrections.json"
    roster_path   = DATA_DIR / "roster.json"
    data_path     = BASE_DIR / "data.json"

    with open(ids_path, encoding="utf-8") as f:
        player_ids = json.load(f)
    with open(lineups_path, encoding="utf-8") as f:
        all_lineups = json.load(f)
    with open(corr_path, encoding="utf-8") as f:
        all_corrections = json.load(f)
    with open(roster_path, encoding="utf-8") as f:
        roster = json.load(f)
    with open(data_path, encoding="utf-8") as f:
        data = json.load(f)

    lineups = all_lineups.get(str(journee), {})
    corrections = all_corrections.get(str(journee), {})

    if not lineups:
        print(f"⚠️  Aucune compo définie pour J{journee}. Configure lineups.json via l'interface admin.")
        sys.exit(1)

    print(f"📡 Récupération des matchs J{journee}...")
    fixtures = get_fixtures(journee, headers)
    if not fixtures:
        print(f"⚠️  Aucun match trouvé pour J{journee} (saison {SEASON}, ligue {LEAGUE_ID})")
        sys.exit(1)
    print(f"  {len(fixtures)} match(s) trouvé(s)")

    # Construire un index : api_player_id → stats + équipe + score match
    player_stats_index = {}   # api_id → {stats, poste_api, goals_conceded}
    player_name_index  = {}   # api_id → nom canonique API

    for fixture in fixtures:
        fx_id      = fixture["fixture"]["id"]
        team_home  = fixture["teams"]["home"]["id"]
        team_away  = fixture["teams"]["away"]["id"]
        score_home = fixture["score"]["fulltime"]["home"] or 0
        score_away = fixture["score"]["fulltime"]["away"] or 0

        print(f"  📊 Fixture {fx_id} : {fixture['teams']['home']['name']} {score_home}-{score_away} {fixture['teams']['away']['name']}")

        fx_players = get_fixture_players(fx_id, headers)
        for team_data in fx_players:
            team_id = team_data["team"]["id"]
            goals_conceded = score_away if team_id == team_home else score_home

            for entry in team_data.get("players", []):
                api_id = entry["player"]["id"]
                api_name = entry["player"]["name"]
                stats = entry.get("statistics", [{}])[0]
                player_stats_index[api_id] = {
                    "stats": stats,
                    "goals_conceded": goals_conceded,
                }
                player_name_index[api_id] = api_name

    # Calculer les points pour chaque manager
    detail_journee = {}
    scores_journee = {}

    for manager, lineup in lineups.items():
        titulaires = set(lineup.get("titulaires", []))
        capitaine  = lineup.get("capitaine")
        coeff      = lineup.get("coeff", 1)  # classement du manager au début de la journée

        equipe_result = {"G": [], "D": [], "M": [], "A": []}
        total_manager = 0

        for poste, joueurs in roster.get(manager, {}).items():
            for nom in joueurs:
                is_titu = nom in titulaires
                api_id  = player_ids.get(nom)
                corr    = corrections.get(manager, {}).get(nom, {})

                if api_id and api_id in player_stats_index:
                    entry           = player_stats_index[api_id]
                    stats_raw       = entry["stats"]
                    goals_conceded  = entry["goals_conceded"]
                    from scoring import calcul_joueur
                    result = parse_player_stats(stats_raw, poste, goals_conceded)
                    # Appliquer corrections
                    if corr:
                        result = calcul_joueur(
                            poste=poste,
                            minutes=int(result["tj"]) if result["tj"] not in ("M", "0") else (90 if result["tj"] == "M" else 0),
                            is_full_match=(result["tj"] == "M"),
                            goals_scored=result["bm"]["val"],
                            assists=result["pd"]["val"],
                            goals_conceded=goals_conceded,
                            penalties_scored=result["pm"]["val"],
                            penalties_missed=result["pma"]["val"] if poste != "G" else 0,
                            penalties_saved_or_opp_missed=result["pma"]["val"] if poste == "G" else 0,
                            own_goals=result["bcsc"]["val"],
                            yellow_cards=result["cj"]["val"],
                            red_card=(result["cr"]["val"] < 0),
                            corrections=corr,
                        )
                else:
                    # Joueur sans ID ou pas joué ce match-là
                    from scoring import calcul_joueur
                    result = calcul_joueur(poste, 0, False, 0, 0, 0, 0, 0, 0, 0, 0, False, corr)

                # Appliquer capitaine
                if is_titu and nom == capitaine:
                    pts_cap = appliquer_capitaine(result["pts"], coeff)
                    result["cap"] = str(coeff)
                    result["pts"] = pts_cap
                elif not is_titu:
                    result["pts"] = 0  # remplaçant ne marque pas de points

                if is_titu:
                    total_manager += result["pts"]

                statut = "" if is_titu else "r"
                equipe_result[poste].append({
                    "nom":     nom,
                    "statut":  statut,
                    "cap":     result.get("cap", ""),
                    "tj":      result["tj"],
                    "tj_pts":  result["tj_pts"],
                    "bm":      result["bm"],
                    "be":      result["be"],
                    "bcsc":    result["bcsc"],
                    "cs":      result["cs"],
                    "pm":      result["pm"],
                    "pma":     result["pma"],
                    "pd":      result["pd"],
                    "cj":      result["cj"],
                    "cr":      result["cr"],
                    "pts":     result["pts"],
                })

        detail_journee[manager] = equipe_result
        scores_journee[manager] = total_manager
        print(f"  ✅ {manager}: {total_manager} pts")

    # Mettre à jour data.json
    data["historique"][str(journee)]       = scores_journee
    data["detail_journees"][str(journee)]  = detail_journee
    data["scores_journee"]                 = scores_journee
    data["derniere_journee"]               = max(data["derniere_journee"], journee)

    # Recalculer évolution
    noms = [j["nom"] for j in data["classement"]]
    cumul = {n: 0 for n in noms}
    evolution = {n: [] for n in noms}
    for jj in range(1, data["derniere_journee"] + 1):
        if str(jj) in data["historique"]:
            for n in noms:
                cumul[n] += data["historique"][str(jj)].get(n, 0)
            for n in noms:
                evolution[n].append({"j": jj, "pts": cumul[n]})
    data["evolution"] = evolution

    # Recalculer classement
    data["classement"] = sorted(
        [{"rang": 0, "nom": n, "pts": cumul[n]} for n in noms],
        key=lambda x: -x["pts"],
    )
    for i, j in enumerate(data["classement"]):
        j["rang"] = i + 1

    with open(data_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"\n✅ data.json mis à jour pour J{journee}")
    return scores_journee


# ─── Entrée ────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    sys.path.insert(0, str(SCRIPTS_DIR))
    api_key = load_api_key()
    headers = {"x-apisports-key": api_key}

    if len(sys.argv) < 2:
        print("Usage : python scripts/fetch_stats.py <journee>")
        print("        python scripts/fetch_stats.py --build-ids")
        sys.exit(1)

    if sys.argv[1] == "--build-ids":
        build_player_ids(headers)
    else:
        try:
            journee = int(sys.argv[1])
        except ValueError:
            print(f"Journée invalide : {sys.argv[1]}")
            sys.exit(1)
        fetch_journee(journee, headers)

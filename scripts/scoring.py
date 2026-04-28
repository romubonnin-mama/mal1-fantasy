"""
Moteur de calcul des points Ma L1.
Source: règlement 2025-2026.

Barème TJ (hors arrêts de jeu) — identique pour tous les postes :
  0 min           -> 0 pt
  1–29 min        -> 1 pt
  30–59 min       -> 2 pts
  60–89 min       -> 3 pts
  match entier    -> 4 pts
  carton rouge    -> 0 pt TJ (bonuses/maluses conservés, CJ annulé)
"""

# ─── Barème ────────────────────────────────────────────────────────────────────

# Points par but marqué selon le poste
BM_PTS = {"G": 5, "D": 3, "M": 2, "A": 2}

# Points passe décisive (tous postes)
PD_PTS = 2

# Bonus pénalty marqué (tous postes, en plus du but)
PM_PTS = 2

# Malus pénalty manqué par le joueur (D/M/A) ou bonus arrêté/adversaire manqué (G)
PMA_PTS = {"G": 2, "D": -2, "M": -2, "A": -2}

# Malus but contre son camp (tous postes)
BCSC_PTS = -2

# Malus carton jaune (tous postes)
CJ_PTS = -1

# Bonus clean sheet
CS_PTS = {"G": 2, "D": 1, "M": 0, "A": 0}

# Malus buts encaissés (≥3 goals conceded while on pitch) — flat malus
BE_PTS = {"G": -2, "D": -1, "M": 0, "A": 0}


# ─── TJ ────────────────────────────────────────────────────────────────────────

def tj_points(minutes: int, is_full_match: bool, red_card: bool) -> int:
    """
    minutes      : minutes played (excluding stoppage time, capped at 90)
    is_full_match: True si le joueur a joué les 90 min réglementaires entières
    red_card     : True si carton rouge reçu (annule les points TJ)
    """
    if red_card:
        return 0
    if is_full_match:
        return 4
    m = min(minutes, 90)
    if m <= 0:
        return 0
    elif m < 30:
        return 1
    elif m < 60:
        return 2
    else:
        return 3


# ─── CS et BE ──────────────────────────────────────────────────────────────────

def cs_points(poste: str, minutes_on_pitch: int, goals_conceded: int) -> int:
    """
    CS gardin  : n'a pas pris de but + a joué au moins une mi-temps entière (≥45 min).
    CS défenseur : n'a pas pris de but + a joué plus de 45 min.
    Le gardien qui sort en ayant conservé sa cage à 0 GARDE le CS même si le remplaçant prend un but.
    """
    if goals_conceded > 0:
        return 0
    if poste == "G" and minutes_on_pitch >= 45:
        return CS_PTS["G"]
    if poste == "D" and minutes_on_pitch > 45:
        return CS_PTS["D"]
    return 0


def be_points(poste: str, goals_conceded: int) -> int:
    """Malus buts encaissés (≥3 buts, malus unique peu importe le nombre)."""
    if goals_conceded >= 3:
        return BE_PTS[poste]
    return 0


# ─── Calcul principal ──────────────────────────────────────────────────────────

def calcul_joueur(
    poste: str,
    minutes: int,
    is_full_match: bool,
    goals_scored: int,
    assists: int,
    goals_conceded: int,
    penalties_scored: int,
    penalties_missed: int,
    penalties_saved_or_opp_missed: int,
    own_goals: int,
    yellow_cards: int,
    red_card: bool,
    corrections: dict = None,
) -> dict:
    """
    Calcule les points d'un joueur et retourne un dict au format data.json.

    corrections : dict optionnel avec les mêmes clés (bm, be, …) pour ajuster
                  les valeurs brutes (ex: {'pd': {'val': -1}} pour annuler une passe).
    """
    if corrections is None:
        corrections = {}

    def adj(field, val):
        """Applique une correction brute au champ."""
        if field in corrections and "val" in corrections[field]:
            return val + corrections[field]["val"]
        return val

    goals_scored    = adj("bm",   goals_scored)
    assists         = adj("pd",   assists)
    goals_conceded  = adj("be",   goals_conceded)
    penalties_scored = adj("pm",  penalties_scored)
    penalties_missed = adj("pma", penalties_missed)
    penalties_saved_or_opp_missed = adj("pma", penalties_saved_or_opp_missed) if poste == "G" else penalties_saved_or_opp_missed
    own_goals       = adj("bcsc", own_goals)
    yellow_cards    = adj("cj",   yellow_cards)

    # Carton rouge annule le carton jaune
    if red_card:
        yellow_cards = 0

    tj_p   = tj_points(minutes, is_full_match, red_card)
    bm_p   = goals_scored * BM_PTS[poste]
    pd_p   = assists * PD_PTS
    pm_p   = penalties_scored * PM_PTS
    bcsc_p = own_goals * BCSC_PTS
    cj_p   = yellow_cards * CJ_PTS
    cr_p   = 0  # red card gives 0 TJ but no additional malus here

    # PMA : pénalty arrêté/manqué adversaire (G) ou pénalty manqué par le joueur (D/M/A)
    if poste == "G":
        pma_val = penalties_saved_or_opp_missed
        pma_p   = pma_val * PMA_PTS["G"]
    else:
        pma_val = penalties_missed
        pma_p   = pma_val * PMA_PTS[poste]

    cs_p = cs_points(poste, minutes, goals_conceded)
    be_p = be_points(poste, goals_conceded)

    # TJ display
    if is_full_match:
        tj_str = "M"
    elif red_card and minutes == 0:
        tj_str = "0"
    else:
        tj_str = str(minutes)

    total = tj_p + bm_p + pd_p + pm_p + pma_p + cs_p + be_p + bcsc_p + cj_p + cr_p

    return {
        "tj":      tj_str,
        "tj_pts":  tj_p,
        "bm":      {"val": goals_scored,                       "pts": bm_p},
        "be":      {"val": goals_conceded,                     "pts": be_p},
        "bcsc":    {"val": own_goals,                          "pts": bcsc_p},
        "cs":      {"val": 1 if cs_p else 0,                   "pts": cs_p},
        "pm":      {"val": penalties_scored,                   "pts": pm_p},
        "pma":     {"val": pma_val,                            "pts": pma_p},
        "pd":      {"val": assists,                            "pts": pd_p},
        "cj":      {"val": yellow_cards,                       "pts": cj_p},
        "cr":      {"val": -1 if red_card else 0,              "pts": 0},
        "pts":     total,
    }


def appliquer_capitaine(pts: int, rang: int) -> int:
    """
    Points totaux capitaine = pts + pts × coeff.
    coeff = classement au début de la journée, plafonné à 7.
    Exemple : 1er → pts × 2, 7e → pts × 8, 8e/9e → pts × 8.
    """
    coeff = min(rang, 7)
    return pts + pts * coeff


# ─── Test rapide ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    # GK avec match entier + CS + 1 pénalty arrêté
    r = calcul_joueur("G", 90, True, 0, 0, 0, 0, 0, 1, 0, 0, False)
    print(f"G match entier + CS + PA : {r['pts']} pts  (attendu: 4+2+2=8)")

    # Défenseur 70 min + 1 but + CJ
    r = calcul_joueur("D", 70, False, 1, 0, 1, 0, 0, 0, 0, 1, False)
    print(f"D 70min + 1 but + CJ     : {r['pts']} pts  (attendu: 3+3-1=5)")

    # Attaquant match entier + 2 buts + 1 passe
    r = calcul_joueur("A", 90, True, 2, 1, 0, 0, 0, 0, 0, 0, False)
    print(f"A match entier + 2B + 1A : {r['pts']} pts  (attendu: 4+4+2=10)")

    # Milieu carton rouge après 30 min + 1 but
    r = calcul_joueur("M", 30, False, 1, 0, 0, 0, 0, 0, 0, 0, True)
    print(f"M CR à 30 min + 1 but    : {r['pts']} pts  (attendu: 0+2=2)")

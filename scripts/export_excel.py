"""
Synchronise les stats admin (manual_stats.json + lineups.json) vers le fichier Excel.
Appelé automatiquement après chaque compute_journee.
"""

import copy
import json
import sys
from pathlib import Path

import openpyxl

BASE_DIR = Path(__file__).parent.parent
DATA_DIR = BASE_DIR / "data"
EXCEL_PATH = Path(r"C:\Users\boro7\OneDrive\Documents\Draft Club\Saison 2025-2026.xlsm")

# Position de chaque manager dans la grille (groupe, colonne)
# Groupe 0 = lignes 2-40, groupe 1 = lignes 42-80, groupe 2 = lignes 82-120
# Colonne 0=gauche, 1=milieu, 2=droite
MANAGER_GRID = {
    (0, 0): "ROMU",    (0, 1): "ADRIEN",  (0, 2): "ANTHONY",
    (1, 0): "JEROME",  (1, 1): "FLORIAN", (1, 2): "BASTIEN",
    (2, 0): "VINCENT", (2, 1): "FAB",     (2, 2): "MICKA",
}
MANAGER_TO_GRID = {v: k for k, v in MANAGER_GRID.items()}

# Ligne d'en-tête de chaque groupe
GROUP_HEADER_ROWS = [2, 42, 82]

# Offsets de ligne (depuis la ligne d'en-tête) pour chaque joueur
POS_ROW_OFFSETS = {
    "G": [2],
    "D": [5, 7, 9, 11, 13],
    "M": [18, 20, 22, 24, 26, 28],
    "A": [31, 33, 35],
}


def is_new_format(sheet_num: int) -> bool:
    return sheet_num >= 14


def get_name_col(col_pos: int, new_fmt: bool) -> int:
    """Colonne 1-based du label de position / nom manager."""
    block_width = 19 if new_fmt else 18
    return 2 + col_pos * block_width


def get_offsets(new_fmt: bool) -> dict:
    """Offsets depuis name_col vers chaque cellule d'entrée."""
    if new_fmt:
        return {
            "name": 1, "status": 3, "cap": 4, "tj": 6,
            "bm": 7, "be": 8, "bcsc": 9, "cs": 10,
            "pm": 11, "pma": 12, "pd": 13, "cj": 14, "cr": 15,
        }
    else:
        return {
            "name": 1, "status": 2, "cap": 4, "tj": 5,
            "bm": 6, "be": 7, "bcsc": 8, "cs": 9,
            "pm": 10, "pma": 11, "pd": 12, "cj": 13, "cr": 14,
        }


def build_player_row_map(ws, group_idx: int, name_col: int, off: dict) -> dict:
    """Scanne la feuille pour trouver nom → numéro de ligne."""
    header_row = GROUP_HEADER_ROWS[group_idx]
    col = name_col + off["name"]
    rows = {}
    for pos, offsets in POS_ROW_OFFSETS.items():
        for delta in offsets:
            row = header_row + delta
            val = ws.cell(row=row, column=col).value
            if val:
                rows[str(val).strip().upper()] = row
    return rows


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


def compute_cs(poste: str, goals_conceded: int, minutes: int, full_match: bool) -> int:
    """Recalcule le clean sheet comme dans scoring.py."""
    if goals_conceded > 0:
        return 0
    if poste == "G" and (full_match or minutes >= 45):
        return 1
    if poste == "D" and (full_match or minutes > 45):
        return 1
    return 0


def _create_sheet(wb, journee: int, new_fmt: bool, off: dict):
    """Crée la feuille journée en copiant la feuille vierge '38' comme modèle."""
    if "38" not in wb.sheetnames:
        return None
    src = wb["38"]
    tgt = wb.copy_worksheet(src)
    tgt.title = str(journee)

    # copy_worksheet ne transfère pas toujours la mise en forme conditionnelle
    tgt.conditional_formatting = copy.deepcopy(src.conditional_formatting)

    # Zoom 72%
    tgt.sheet_view.zoomScale = 72

    # Code name VBA unique — évite qu'Excel confonde cette feuille avec "38" à la suppression
    tgt.sheet_properties.codeName = f"J{journee}"

    # Déplacer la feuille à la bonne position (ordre numérique entre les journées)
    sheets = wb.sheetnames
    current_idx = sheets.index(str(journee))
    try:
        target_idx = next(i for i, s in enumerate(sheets) if s.isdigit() and int(s) > journee)
        wb.move_sheet(str(journee), offset=target_idx - current_idx)
    except StopIteration:
        idx_38 = sheets.index("38")
        wb.move_sheet(str(journee), offset=idx_38 - current_idx)
    return tgt


def export_journee(journee: int, verbose: bool = True) -> None:
    with open(DATA_DIR / "roster.json", encoding="utf-8") as f:
        roster = json.load(f)
    with open(DATA_DIR / "lineups.json", encoding="utf-8") as f:
        lineups = json.load(f)
    with open(DATA_DIR / "manual_stats.json", encoding="utf-8") as f:
        manual_stats = json.load(f)

    j_lineups = lineups.get(str(journee), {})
    j_stats = manual_stats.get(str(journee), {})

    if not j_lineups and not j_stats:
        if verbose:
            print(f"Aucune donnée admin pour la journée {journee}, export ignoré.")
        return

    sheet_name = str(journee)
    new_fmt = is_new_format(journee)
    off = get_offsets(new_fmt)

    wb = openpyxl.load_workbook(EXCEL_PATH, keep_vba=True)

    if sheet_name not in wb.sheetnames:
        ws = _create_sheet(wb, journee, new_fmt, off)
        if ws is None:
            if verbose:
                print(f"Impossible de créer la feuille '{sheet_name}'.")
            wb.close()
            return
        if verbose:
            print(f"Feuille '{sheet_name}' créée par copie.")

    ws = wb[sheet_name]

    for manager, player_positions in roster.items():
        grid_pos = MANAGER_TO_GRID.get(manager)
        if not grid_pos:
            continue

        group_idx, col_pos = grid_pos
        name_col = get_name_col(col_pos, new_fmt)
        player_rows = build_player_row_map(ws, group_idx, name_col, off)

        lineup     = j_lineups.get(manager, {})
        titulaires = set(lineup.get("titulaires", []))
        capitaine  = lineup.get("capitaine", "")
        coeff      = int(lineup.get("coeff", 1))
        m_stats    = j_stats.get(manager, {})

        for poste in ["G", "D", "M", "A"]:
            for player in player_positions.get(poste, []):
                row = player_rows.get(player.strip().upper())
                if not row:
                    if verbose:
                        print(f"  [warn] {manager}/{player} introuvable dans feuille '{sheet_name}'")
                    continue

                is_titu = player in titulaires
                stats   = m_stats.get(player, {}) if is_titu else {}

                def w(field, value):
                    if field in off:
                        ws.cell(row=row, column=name_col + off[field]).value = value

                if not is_titu:
                    # Remplaçant: 'r' en colonne statut, tout le reste effacé
                    w("status", "r")
                    w("cap", None)
                    w("tj", None)
                    for field in ["bm","be","bcsc","cs","pm","pma","pd","cj","cr"]:
                        w(field, None)

                elif stats.get("absent"):
                    # Titulaire absent / blessé / non entré : 'A' en statut, tout le reste effacé
                    w("status", "A")
                    w("cap", None)
                    w("tj", None)
                    for field in ["bm","be","bcsc","cs","pm","pma","pd","cj","cr"]:
                        w(field, None)

                elif not stats:
                    # Titulaire sans stats saisies : ne rien écrire
                    pass

                else:
                    # Titulaire avec stats
                    full_match = bool(stats.get("full_match", False))
                    minutes    = _minutes(stats)
                    red_card   = bool(stats.get("red_card", False))

                    if red_card:
                        tj_val = minutes if minutes > 0 else None
                        status = None  # titulaire qui a joué, carton rouge ≠ absent
                    elif full_match:
                        tj_val, status = "M", None
                    elif minutes > 0:
                        tj_val, status = minutes, None
                    else:
                        tj_val, status = None, "A"

                    w("status", status)
                    w("tj",     tj_val)
                    w("cap",    coeff if player == capitaine else None)

                    goals_c = int(stats.get("goals_conceded", 0))
                    cs_val  = 1 if (stats.get("cs") or compute_cs(poste, goals_c, minutes, full_match)) else None
                    w("bm",   int(stats.get("goals", 0))        or None)
                    w("be",   goals_c                            or None)
                    w("bcsc", int(stats.get("own_goals", 0))     or None)
                    w("cs",   cs_val)
                    w("pm",   int(stats.get("pen_scored", 0))    or None)
                    w("pma",  int(stats.get("pen_mm_saved", 0))  or None)
                    w("pd",   int(stats.get("assists", 0))       or None)
                    w("cj",   int(stats.get("yellow_cards", 0))  or None)
                    w("cr",   1 if red_card                      else None)

    wb.save(EXCEL_PATH)
    if verbose:
        print(f"Excel mis a jour : journee {journee} -> {EXCEL_PATH.name}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python export_excel.py <journee>")
        sys.exit(1)
    export_journee(int(sys.argv[1]))

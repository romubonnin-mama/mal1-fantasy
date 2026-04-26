import json
import subprocess
import sys

try:
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

EXCEL_PATH = r"C:\Users\boro7\OneDrive\Documents\Draft Club\Saison 2025-2026.xlsm"

# Colonne de départ et ligne de départ pour chaque joueur
JOUEURS_CONFIG = {
    "ROMU":    {"col": 3,  "ligne": 4},
    "JEROME":  {"col": 3,  "ligne": 44},
    "VINCENT": {"col": 3,  "ligne": 84},
    "ADRIEN":  {"col": 22, "ligne": 4},
    "FLORIAN": {"col": 22, "ligne": 44},
    "FAB":     {"col": 22, "ligne": 84},
    "ANTHONY": {"col": 41, "ligne": 4},
    "BASTIEN": {"col": 41, "ligne": 44},
    "MICKA":   {"col": 41, "ligne": 84},
}

NOMS_COLS = {
    "ROMU":    (19, 3),
    "JEROME":  (19, 43),
    "VINCENT": (19, 83),
    "ADRIEN":  (38, 3),
    "FLORIAN": (38, 43),
    "FAB":     (38, 83),
    "ANTHONY": (57, 3),
    "BASTIEN": (57, 43),
    "MICKA":   (57, 83),
}

NOMS_COLS_ANCIEN = {
    "ROMU":    (18, 3),
    "JEROME":  (18, 43),
    "VINCENT": (18, 83),
    "ADRIEN":  (36, 3),
    "FLORIAN": (36, 43),
    "FAB":     (36, 83),
    "ANTHONY": (54, 3),
    "BASTIEN": (54, 43),
    "MICKA":   (54, 83),
}

def lire_classement(ws):
    joueurs = []
    for row in range(4, 13):
        rang = ws.cell(row=row, column=2).value
        nom = ws.cell(row=row, column=3).value
        pts = ws.cell(row=row, column=4).value
        if nom and isinstance(pts, (int, float)):
            joueurs.append({"rang": int(rang) if rang else row-3, "nom": str(nom), "pts": int(pts)})
    return sorted(joueurs, key=lambda x: x["rang"])

def lire_scores_journee(ws, ancien=False):
    cols = NOMS_COLS_ANCIEN if ancien else NOMS_COLS
    scores = {}
    for nom, (col, row) in cols.items():
        val = ws.cell(row=row, column=col).value
        if isinstance(val, (int, float)):
            scores[nom] = int(val)
        else:
            scores[nom] = 0
    return scores

def lire_joueur(ws, col, row, ancien=False):
    decalage = 0 if ancien else 1
    nom = ws.cell(row=row, column=col).value
    if not nom or nom in ("", None):
        return None
    statut = ws.cell(row=row, column=col+2).value
    cap = ws.cell(row=row, column=col+3).value
    tj_entree = ws.cell(row=row, column=col+4).value
    tj_sortie = ws.cell(row=row, column=col+5).value
    if str(tj_sortie).upper() == 'M':
        tj = "M"
    elif tj_entree and tj_sortie:
        tj = f"{tj_entree}-{tj_sortie}"
    elif tj_sortie:
        tj = str(tj_sortie)
    else:
        tj = "0"
    bm  = ws.cell(row=row+1, column=col+4+decalage).value or 0
    be  = ws.cell(row=row+1, column=col+5+decalage).value or 0
    bcsc= ws.cell(row=row+1, column=col+6+decalage).value or 0
    cs  = ws.cell(row=row+1, column=col+7+decalage).value or 0
    pm  = ws.cell(row=row+1, column=col+8+decalage).value or 0
    pma = ws.cell(row=row+1, column=col+9+decalage).value or 0
    pd  = ws.cell(row=row+1, column=col+10+decalage).value or 0
    cj  = ws.cell(row=row+1, column=col+11+decalage).value or 0
    cr  = ws.cell(row=row+1, column=col+12+decalage).value or 0
    pts = ws.cell(row=row, column=col+14+decalage).value or 0

    return {
        "nom": str(nom),
        "statut": str(statut) if statut else "",
        "cap": str(cap) if cap else "",
        "tj": str(tj) if tj else "0",
        "bm": int(bm) if isinstance(bm, (int,float)) else 0,
        "be": int(be) if isinstance(be, (int,float)) else 0,
        "bcsc": int(bcsc) if isinstance(bcsc, (int,float)) else 0,
        "cs": int(cs) if isinstance(cs, (int,float)) else 0,
        "pm": int(pm) if isinstance(pm, (int,float)) else 0,
        "pma": int(pma) if isinstance(pma, (int,float)) else 0,
        "pd": int(pd) if isinstance(pd, (int,float)) else 0,
        "cj": int(cj) if isinstance(cj, (int,float)) else 0,
        "cr": int(cr) if isinstance(cr, (int,float)) else 0,
        "pts": int(pts) if isinstance(pts, (int,float)) else 0,
    }

def lire_equipe(ws, nom_joueur, ancien=False):
    cfg = JOUEURS_CONFIG[nom_joueur]
    col = cfg["col"]
    ligne = cfg["ligne"]

    equipe = {"G": [], "D": [], "M": [], "A": []}

    # Gardien (ligne de base)
    g = lire_joueur(ws, col, ligne, ancien)
    if g:
        equipe["G"].append(g)

    # Défenseurs (ligne+3 à ligne+14, pas de 2)
    for r in range(ligne+3, ligne+15, 2):
        d = lire_joueur(ws, col, r, ancien)
        if d:
            equipe["D"].append(d)

    # Milieux (ligne+16 à ligne+27, pas de 2)
    for r in range(ligne+16, ligne+28, 2):
        m = lire_joueur(ws, col, r, ancien)
        if m:
            equipe["M"].append(m)

    # Attaquants (ligne+29 à ligne+36, pas de 2)
    for r in range(ligne+29, ligne+37, 2):
        a = lire_joueur(ws, col, r, ancien)
        if a:
            equipe["A"].append(a)

    return equipe

def derniere_journee(wb):
    max_j = 0
    for sheet in wb.sheetnames:
        if sheet.isdigit() and int(sheet) != 38:
            j = int(sheet)
            ws = wb[sheet]
            col = 18 if j < 14 else 19
            val = ws.cell(row=3, column=col).value
            if isinstance(val, (int, float)) and val > 0:
                max_j = max(max_j, j)
    return max_j if max_j > 0 else 1

def main():
    print("Lecture du fichier Excel...")
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    except FileNotFoundError:
        print("ERREUR : Fichier introuvable")
        input("Appuie sur Entree pour fermer...")
        return

    ws_scores = wb["SCORES"]
    classement = lire_classement(ws_scores)
    derniere_j = derniere_journee(wb)
    print(f"Derniere journee detectee : J{derniere_j}")

    historique = {}
    detail_journees = {}

    for j in range(1, 35):
        if str(j) in wb.sheetnames and j != 38:
            ws_j = wb[str(j)]
            ancien = (j < 14)
            historique[str(j)] = lire_scores_journee(ws_j, ancien)
            detail_journees[str(j)] = {}
            for nom in JOUEURS_CONFIG:
                detail_journees[str(j)][nom] = lire_equipe(ws_j, nom, ancien)
            print(f"  J{j} extraite")

    # Calcul evolution cumulee
    joueurs_noms = [j["nom"] for j in classement]
    evolution = {nom: [] for nom in joueurs_noms}
    cumul = {nom: 0 for nom in joueurs_noms}
    for j in range(1, derniere_j + 1):
        if str(j) in historique:
            for nom in joueurs_noms:
                cumul[nom] += historique[str(j)].get(nom, 0)
            for nom in joueurs_noms:
                evolution[nom].append({"j": j, "pts": cumul[nom]})

    score_max_val = ws_scores.cell(row=14, column=9).value
    score_max_nom = ws_scores.cell(row=14, column=10).value
    score_min_val = ws_scores.cell(row=15, column=9).value
    score_min_nom = ws_scores.cell(row=15, column=10).value

    data = {
        "classement": classement,
        "derniere_journee": derniere_j,
        "scores_journee": historique[str(derniere_j)],
        "historique": historique,
        "detail_journees": detail_journees,
        "evolution": evolution,
        "score_max": {"valeur": score_max_val, "joueur": score_max_nom},
        "score_min": {"valeur": score_min_val, "joueur": score_min_nom},
    }

    with open("data.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print("Fichier data.json cree !")
    print("Envoi sur GitHub...")
    subprocess.run(["git", "add", "."], check=True)
    result = subprocess.run(["git", "commit", "-m", f"Mise a jour J{derniere_j}"])
    if result.returncode == 0:
        subprocess.run(["git", "push"], check=True)
    else:
        print("Aucun changement a envoyer.")
    print("Site mis a jour avec succes !")
    input("Appuie sur Entree pour fermer...")

if __name__ == "__main__":
    main()
import json
import subprocess
import sys

try:
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

EXCEL_PATH = r"C:\Users\boro7\OneDrive\Documents\Draft Club\Saison 2025-2026.xlsm"

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

def lire_classement(ws):
    joueurs = []
    for row in range(4, 13):
        rang = ws.cell(row=row, column=2).value
        nom = ws.cell(row=row, column=3).value
        pts = ws.cell(row=row, column=4).value
        if nom and isinstance(pts, (int, float)):
            joueurs.append({"rang": int(rang) if rang else row-1, "nom": str(nom), "pts": int(pts)})
    return sorted(joueurs, key=lambda x: x["rang"])

def derniere_journee(wb):
    max_j = 0
    for sheet in wb.sheetnames:
        if sheet.isdigit():
            j = int(sheet)
            ws = wb[sheet]
            val = ws.cell(row=3, column=19).value
            if isinstance(val, (int, float)) and val > 0:
                max_j = max(max_j, j)
    return max_j if max_j > 0 else 1

def lire_scores_journee(ws):
    scores = {}
    for nom, (col, row) in NOMS_COLS.items():
        val = ws.cell(row=row, column=col).value
        if isinstance(val, (int, float)):
            scores[nom] = int(val)
        else:
            scores[nom] = 0
    return scores

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

    ws_j = wb[str(derniere_j)]
    scores_journee = lire_scores_journee(ws_j)

    score_max_val = ws_scores.cell(row=14, column=9).value
    score_max_nom = ws_scores.cell(row=14, column=10).value
    score_min_val = ws_scores.cell(row=15, column=9).value
    score_min_nom = ws_scores.cell(row=15, column=10).value

    data = {
        "classement": classement,
        "derniere_journee": derniere_j,
        "scores_journee": scores_journee,
        "score_max": {"valeur": score_max_val, "joueur": score_max_nom},
        "score_min": {"valeur": score_min_val, "joueur": score_min_nom},
    }

    with open("data.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print("Fichier data.json cree !")
    print("Envoi sur GitHub...")
    subprocess.run(["git", "add", "."], check=True)
    subprocess.run(["git", "commit", "-m", f"Mise a jour J{derniere_j}"], check=True)
    subprocess.run(["git", "push"], check=True)
    print("Site mis a jour avec succes !")
    input("Appuie sur Entree pour fermer...")

if __name__ == "__main__":
    main()
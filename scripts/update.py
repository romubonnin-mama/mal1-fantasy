import json
import subprocess
import sys

# Installe openpyxl si nécessaire
try:
    import openpyxl
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

# Chemin vers ton fichier Excel
EXCEL_PATH = r"C:\Users\boro7\OneDrive\Documents\Draft Club\Saison 2025-2026.xlsm"

def lire_classement(wb):
    ws = wb["SCORES"]
    joueurs = []
    for row in range(2, 11):
        rang = ws.cell(row=row, column=2).value
        nom = ws.cell(row=row, column=3).value
        pts = ws.cell(row=row, column=4).value
        if nom and pts:
            joueurs.append({"rang": rang, "nom": nom, "pts": pts})
    return joueurs

def lire_journee(wb, numero):
    try:
        ws = wb[str(numero)]
    except KeyError:
        return []
    
    scores = {}
    noms = ["ROMU", "ADRIEN", "ANTHONY", "JEROME", "FLORIAN", "BASTIEN", "VINCENT", "FAB", "MICKA"]
    cols = [1, 19, 37, 55, 73, 91, 109, 127, 145]
    
    for nom, col in zip(noms, cols):
        total = ws.cell(row=2, column=col+17).value
        if total is None:
            total = 0
        scores[nom] = total
    
    return scores

def main():
    print("Lecture du fichier Excel...")
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    except FileNotFoundError:
        print(f"ERREUR : Fichier introuvable : {EXCEL_PATH}")
        input("Appuie sur Entrée pour fermer...")
        return

    classement = lire_classement(wb)
    
    # Dernière journée disponible
    derniere_j = 1
    for sheet in wb.sheetnames:
        if sheet.isdigit():
            derniere_j = max(derniere_j, int(sheet))
    
    scores_journee = lire_journee(wb, derniere_j)
    
    data = {
        "classement": classement,
        "derniere_journee": derniere_j,
        "scores_journee": scores_journee,
        "score_max": {"valeur": 98, "joueur": "ROMU", "journee": 22},
        "score_min": {"valeur": 24, "joueur": "ADRIEN", "journee": 17}
    }
    
    with open("data.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    print("Fichier data.json créé !")
    print("Envoi sur GitHub...")
    
    subprocess.run(["git", "add", "."], check=True)
    subprocess.run(["git", "commit", "-m", f"Mise a jour J{derniere_j}"], check=True)
    subprocess.run(["git", "push"], check=True)
    
    print("Site mis à jour avec succès !")
    input("Appuie sur Entrée pour fermer...")

if __name__ == "__main__":
    main()
"""
Transfère les images/drawings d'un ancien fichier Excel vers le fichier actuel.
Usage : python restore_images.py
"""

import io
import re
import zipfile
from pathlib import Path

CURRENT = Path(r"C:\Users\boro7\OneDrive\Documents\Draft Club\Saison 2025-2026.xlsm")
OLD     = Path(r"C:\Users\boro7\OneDrive\Documents\Draft Club\Saison 2025-2026_old.xlsm")


def is_drawing(name):
    return (name.startswith('xl/media/') or
            name.startswith('xl/drawings/') or
            name.startswith('xl/charts/'))


def has_drawing_ref(text):
    return bool(re.search(r'drawings?/|/media/', text, re.IGNORECASE))


def transfer_images(current_path: Path, old_path: Path) -> None:
    print(f"Source images : {old_path.name}")
    print(f"Cible         : {current_path.name}")

    with zipfile.ZipFile(old_path, 'r') as old_z:
        drawing_files = [n for n in old_z.namelist() if is_drawing(n)]
        print(f"Fichiers drawing/media trouvés dans l'ancien : {len(drawing_files)}")
        for n in drawing_files:
            print(f"  {n}")

    if not drawing_files:
        print("Aucune image à transférer.")
        return

    result = io.BytesIO()
    with zipfile.ZipFile(current_path, 'r') as cur, \
         zipfile.ZipFile(old_path, 'r') as old_z, \
         zipfile.ZipFile(result, 'w', zipfile.ZIP_DEFLATED) as out:

        cur_names = set(cur.namelist())
        written   = set()

        for name in cur.namelist():
            if is_drawing(name) and name in old_z.namelist():
                out.writestr(name, old_z.read(name))
            elif name.endswith('.rels') and name in old_z.namelist():
                old_rels = old_z.read(name).decode('utf-8', errors='replace')
                if has_drawing_ref(old_rels):
                    out.writestr(name, old_z.read(name))
                else:
                    out.writestr(name, cur.read(name))
            elif name == '[Content_Types].xml':
                new_ct  = cur.read(name).decode('utf-8')
                old_ct  = old_z.read(name).decode('utf-8')
                for entry in re.findall(r'<(?:Override|Default)[^/]*/>', old_ct):
                    if ('drawing' in entry.lower() or 'chart' in entry.lower()
                            or re.search(r'png|jpe?g|gif|emf|wmf', entry, re.I)):
                        key = re.search(r'(?:PartName|Extension)="([^"]+)"', entry)
                        if key and key.group(1) not in new_ct:
                            new_ct = new_ct.replace('</Types>', f'  {entry}\n</Types>')
                out.writestr(name, new_ct.encode('utf-8'))
            else:
                out.writestr(name, cur.read(name))
            written.add(name)

        # Ajouter les fichiers drawing absents du fichier actuel
        for name in old_z.namelist():
            if name in written:
                continue
            if is_drawing(name):
                out.writestr(name, old_z.read(name))
                print(f"  Ajouté (manquant) : {name}")
            elif name.endswith('.rels'):
                content = old_z.read(name).decode('utf-8', errors='replace')
                if has_drawing_ref(content):
                    out.writestr(name, old_z.read(name))

    result.seek(0)
    with open(current_path, 'wb') as f:
        f.write(result.read())

    print("Transfert terminé — formules conservées, images restaurées.")


if __name__ == "__main__":
    if not OLD.exists():
        print(f"ERREUR : fichier source introuvable : {OLD}")
        print("Télécharge l'ancienne version depuis OneDrive et sauvegarde-la sous :")
        print(f"  {OLD}")
    elif not CURRENT.exists():
        print(f"ERREUR : fichier cible introuvable : {CURRENT}")
    else:
        transfer_images(CURRENT, OLD)

"""
Synchronise les stats admin (manual_stats.json + lineups.json) vers le fichier Excel.
Appelé automatiquement après chaque compute_journee.
"""

import copy
import io
import json
import posixpath
import re
import sys
import time
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

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
    """Scanne la feuille pour trouver nom -> numéro de ligne."""
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

    tgt.conditional_formatting = copy.deepcopy(src.conditional_formatting)

    # Zoom 72%
    tgt.sheet_view.zoomScale = 72

    # Code name VBA unique
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


def _save_preserving_images(wb, path: Path, target_sheet: str = None) -> None:
    """
    Sauvegarde le workbook en préservant les images/drawings.

    Stratégie par type de feuille :
    - Nouvelle feuille (ex: J32 inexistante) :
        XML original de '38' + sheetData openpyxl + drawing copié de '38'
    - Feuille cible existante (ex: J32 déjà présente) :
        XML original + sheetData openpyxl (namespaces corrigés)
    - Toutes les autres feuilles ('38', 'club', autres J) :
        XML original INCHANGÉ — on ne touche à rien
    """
    NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    NS_S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

    path = Path(path)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    if not path.exists():
        with open(path, 'wb') as f:
            f.write(buf.read())
        return

    def is_media(name):
        return (name.startswith('xl/media/') or
                name.startswith('xl/drawings/') or
                name.startswith('xl/charts/'))

    def sheet_has_drawing(xml_text):
        return bool(re.search(r'<(?:drawing|legacyDrawing)\b', xml_text))

    def find_sheet_xml(zip_obj, sheet_name):
        """Retourne 'xl/worksheets/sheetN.xml' pour un nom de feuille."""
        try:
            root = ET.fromstring(zip_obj.read('xl/workbook.xml'))
            rid = None
            for el in root.iter(f'{{{NS_S}}}sheet'):
                if el.get('name') == sheet_name:
                    rid = el.get(f'{{{NS_R}}}id')
                    break
            if not rid:
                print(f"[img] '{sheet_name}' introuvable dans workbook.xml")
                return None
            rels_root = ET.fromstring(zip_obj.read('xl/_rels/workbook.xml.rels'))
            for rel in rels_root:
                if rel.get('Id') == rid:
                    target = rel.get('Target', '')
                    if target.startswith('/'):
                        return target.lstrip('/')
                    elif target.startswith('xl/'):
                        return target
                    else:
                        return posixpath.normpath('xl/' + target)
            print(f"[img] rId '{rid}' introuvable dans workbook.xml.rels")
            return None
        except Exception as e:
            print(f"[img] find_sheet_xml('{sheet_name}') error: {e}")
            return None

    def find_drawing_info(zip_obj, sheet_xml_path):
        """Retourne (rels_path, drawing_rid, drawing_bytes) ou (None, 'rId1', None)."""
        try:
            sheet_xml = zip_obj.read(sheet_xml_path).decode('utf-8')
            if not sheet_has_drawing(sheet_xml):
                print(f"[img] pas de <drawing> dans {sheet_xml_path}")
                return None, 'rId1', None
            rels_path = sheet_xml_path.replace(
                'xl/worksheets/', 'xl/worksheets/_rels/') + '.rels'
            if rels_path not in zip_obj.namelist():
                print(f"[img] rels manquant: {rels_path}")
                return None, 'rId1', None
            rels_root = ET.fromstring(zip_obj.read(rels_path))
            for rel in rels_root:
                if 'drawing' in rel.get('Type', '').lower():
                    rid    = rel.get('Id', 'rId1')
                    target = rel.get('Target', '')
                    drw    = posixpath.normpath(posixpath.join('xl/worksheets', target))
                    if drw in zip_obj.namelist():
                        return rels_path, rid, zip_obj.read(drw)
                    print(f"[img] drawing manquant: {drw}")
            print(f"[img] aucun drawing rel dans {rels_path}")
            return None, 'rId1', None
        except Exception as e:
            print(f"[img] find_drawing_info error: {e}")
            return None, 'rId1', None

    def add_missing_ns(base_xml, source_xml):
        """Ajoute dans base_xml les xmlns:PREFIX de source_xml qui manquent."""
        src_ns = dict(re.findall(r'xmlns:(\w+)="([^"]+)"', source_xml[:3000]))
        base_head = base_xml[:3000]
        additions = ' '.join(
            f'xmlns:{pfx}="{uri}"'
            for pfx, uri in src_ns.items()
            if f'xmlns:{pfx}' not in base_head
        )
        if additions:
            base_xml = re.sub(
                r'(<worksheet\b[^>]*?)(/>|>)',
                lambda m: m.group(1) + ' ' + additions + m.group(2),
                base_xml, count=1)
        return base_xml

    def inject_sheetdata(base_xml, new_xml):
        """Remplace <sheetData> de base_xml par celui de new_xml."""
        m = re.search(r'(<sheetData(?:\s[^>]*)?>.*?</sheetData>)', new_xml, re.DOTALL)
        if not m:
            return base_xml
        return re.sub(
            r'<sheetData(?:\s[^>]*)?>.*?</sheetData>',
            m.group(1), base_xml, count=1, flags=re.DOTALL)

    result = io.BytesIO()
    with zipfile.ZipFile(path, 'r') as orig, \
         zipfile.ZipFile(buf, 'r') as new_z, \
         zipfile.ZipFile(result, 'w', zipfile.ZIP_DEFLATED) as out:

        orig_names = set(orig.namelist())
        written    = set()

        # ── Nouvelles feuilles : dans new_z mais absentes de orig ───────────────
        new_sheet_xmls = [n for n in new_z.namelist()
                          if re.match(r'xl/worksheets/sheet\d+\.xml$', n)
                          and n not in orig_names]
        print(f"[img] nouvelles feuilles XML: {new_sheet_xmls}")

        # ── Feuille cible ─────────────────────────────────────────────────────────
        target_xml_orig = find_sheet_xml(orig,  target_sheet) if target_sheet else None
        target_xml_new  = find_sheet_xml(new_z, target_sheet) if target_sheet else None
        print(f"[img] target '{target_sheet}': orig={target_xml_orig}, new={target_xml_new}")

        # ── Infos drawing de '38' pour les nouvelles feuilles ────────────────────
        s38_xml           = None
        s38_xml_bytes     = None
        s38_drw_rid       = 'rId1'
        s38_drw_bytes     = None
        new_drawing_map   = {}   # new_sheet_xml -> nouveau drawing path
        new_sheet_rels    = {}   # rels_path -> bytes

        if new_sheet_xmls:
            s38_xml = find_sheet_xml(orig, "38")
            print(f"[img] XML de '38': {s38_xml}")
            if s38_xml and s38_xml in orig_names:
                s38_xml_bytes = orig.read(s38_xml)
                _, rid, drw_bytes = find_drawing_info(orig, s38_xml)
                print(f"[img] drawing de '38': rid={rid}, trouvé={drw_bytes is not None}")
                if drw_bytes:
                    s38_drw_rid   = rid
                    s38_drw_bytes = drw_bytes
                    all_drw = [n for n in orig_names
                               if re.match(r'xl/drawings/drawing\d+\.xml$', n)]
                    max_n = max(
                        (int(re.search(r'(\d+)', n).group(1)) for n in all_drw), default=0)
                    print(f"[img] max drawing num existant: {max_n}")
                    for i, sxml in enumerate(new_sheet_xmls):
                        drw_num = max_n + 1 + i
                        new_drw = f'xl/drawings/drawing{drw_num}.xml'
                        new_drawing_map[sxml] = new_drw
                        rels_path = sxml.replace(
                            'xl/worksheets/', 'xl/worksheets/_rels/') + '.rels'
                        new_sheet_rels[rels_path] = (
                            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                            '<Relationships xmlns="http://schemas.openxmlformats.org'
                            '/package/2006/relationships">\n'
                            f'  <Relationship Id="{s38_drw_rid}"'
                            f' Type="http://schemas.openxmlformats.org/officeDocument'
                            f'/2006/relationships/drawing"'
                            f' Target="../drawings/drawing{drw_num}.xml"/>\n'
                            '</Relationships>'
                        ).encode('utf-8')
                    print(f"[img] drawing map: {new_drawing_map}")

        # ── Boucle principale ─────────────────────────────────────────────────────
        print(f"[img] debut boucle: {len(new_z.namelist())} fichiers")
        for name in new_z.namelist():
            data = None

            if is_media(name) and name in orig_names:
                # Media/drawing existant : toujours garder l'original
                data = orig.read(name)

            elif name.endswith('.rels'):
                if name in new_sheet_rels:
                    # Rels d'une nouvelle feuille : notre version avec drawing
                    data = new_sheet_rels[name]
                elif name == 'xl/_rels/workbook.xml.rels':
                    # Workbook rels : version openpyxl (référence la nouvelle feuille)
                    data = new_z.read(name)
                elif name in orig_names:
                    # Tous les autres rels existants : conserver l'original
                    data = orig.read(name)
                else:
                    data = new_z.read(name)

            elif name == '[Content_Types].xml':
                new_ct  = new_z.read(name).decode('utf-8')
                orig_ct = orig.read(name).decode('utf-8')
                # Copier les entrées drawing/media de l'original
                for entry in re.findall(r'<(?:Override|Default)[^>]*/>', orig_ct):
                    if ('drawing' in entry.lower() or 'chart' in entry.lower()
                            or re.search(r'png|jpe?g|gif|emf|wmf', entry, re.I)):
                        key = re.search(r'(?:PartName|Extension)="([^"]+)"', entry)
                        if key and key.group(1) not in new_ct:
                            new_ct = new_ct.replace('</Types>', f'  {entry}\n</Types>')
                # Enregistrer les nouveaux fichiers drawing
                drw_ct = 'application/vnd.openxmlformats-officedocument.drawing+xml'
                for new_drw in new_drawing_map.values():
                    part = '/' + new_drw
                    if part not in new_ct:
                        new_ct = new_ct.replace(
                            '</Types>',
                            f'  <Override PartName="{part}" ContentType="{drw_ct}"/>\n</Types>')
                data = new_ct.encode('utf-8')

            elif re.match(r'xl/worksheets/sheet\d+\.xml$', name):

                if name in new_drawing_map:
                    # ── Nouvelle feuille : XML original de '38' + sheetData openpyxl ──
                    new_xml = new_z.read(name).decode('utf-8')
                    if s38_xml_bytes:
                        base = s38_xml_bytes.decode('utf-8')
                        base = add_missing_ns(base, new_xml)
                        base = inject_sheetdata(base, new_xml)
                        data = base.encode('utf-8')
                        print(f"[img] nouvelle feuille: XML '38' + sheetData -> {name}")
                    else:
                        # Fallback si '38' introuvable : openpyxl + injection drawing
                        if not sheet_has_drawing(new_xml):
                            new_xml = new_xml.replace(
                                '</worksheet>',
                                f'  <drawing r:id="{s38_drw_rid}"/>\n</worksheet>')
                        data = new_xml.encode('utf-8')
                        print(f"[img] nouvelle feuille (fallback sans '38'): {name}")

                elif (target_xml_new and name == target_xml_new
                      and target_xml_orig and target_xml_orig in orig_names):
                    # ── Feuille cible existante : orig + sheetData openpyxl ──────────
                    orig_xml = orig.read(target_xml_orig).decode('utf-8')
                    new_xml  = new_z.read(name).decode('utf-8')
                    orig_xml = add_missing_ns(orig_xml, new_xml)
                    data     = inject_sheetdata(orig_xml, new_xml).encode('utf-8')
                    print(f"[img] feuille cible '{target_sheet}': sheetData patché -> {name}")

                elif name in orig_names:
                    # ── Toute autre feuille existante : XML original INCHANGÉ ─────────
                    data = orig.read(name)

                else:
                    data = new_z.read(name)

            else:
                # xl/workbook.xml doit venir d'openpyxl (référence la nouvelle feuille)
                # Tout le reste (styles, sharedStrings, etc.) : préférer l'original
                if name == 'xl/workbook.xml':
                    data = new_z.read(name)
                elif name in orig_names:
                    data = orig.read(name)
                else:
                    data = new_z.read(name)

            out.writestr(name, data)
            written.add(name)

        # ── Post-boucle : rels des nouvelles feuilles non écrits ─────────────────
        for rels_name, rels_bytes in new_sheet_rels.items():
            if rels_name not in written:
                out.writestr(rels_name, rels_bytes)
                written.add(rels_name)
                print(f"[img] rels post-boucle: {rels_name}")

        # ── Fichiers drawing des nouvelles feuilles ───────────────────────────────
        if s38_drw_bytes:
            for new_drw in new_drawing_map.values():
                if new_drw not in written:
                    out.writestr(new_drw, s38_drw_bytes)
                    written.add(new_drw)
                    print(f"[img] drawing copié: {new_drw}")

        # ── Tous les fichiers de l'original non encore écrits (rels, media, etc.) ──
        for name in orig.namelist():
            if name not in written:
                out.writestr(name, orig.read(name))
                written.add(name)

    result.seek(0)
    data_to_write = result.read()
    for attempt in range(6):
        try:
            with open(path, 'wb') as f:
                f.write(data_to_write)
            break
        except PermissionError:
            if attempt < 5:
                print(f"[img] Fichier verrouillé (OneDrive ?), retry dans 3s (tentative {attempt+1}/6)...")
                time.sleep(3)
            else:
                raise PermissionError(
                    f"Impossible d'écrire '{path.name}' après 6 tentatives. "
                    "Vérifie qu'Excel et OneDrive ne bloquent pas le fichier."
                )


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
                    w("status", "r")
                    w("cap", None)
                    w("tj", None)
                    for field in ["bm","be","bcsc","cs","pm","pma","pd","cj","cr"]:
                        w(field, None)

                elif stats.get("absent"):
                    w("status", "A")
                    w("cap", None)
                    w("tj", None)
                    for field in ["bm","be","bcsc","cs","pm","pma","pd","cj","cr"]:
                        w(field, None)

                elif not stats:
                    pass

                else:
                    full_match = bool(stats.get("full_match", False))
                    minutes    = _minutes(stats)
                    red_card   = bool(stats.get("red_card", False))

                    if red_card:
                        tj_val = minutes if minutes > 0 else None
                        status = None
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

    _save_preserving_images(wb, EXCEL_PATH, target_sheet=sheet_name)
    if verbose:
        print(f"Excel mis a jour : journee {journee} -> {EXCEL_PATH.name}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python export_excel.py <journee>")
        sys.exit(1)
    export_journee(int(sys.argv[1]))

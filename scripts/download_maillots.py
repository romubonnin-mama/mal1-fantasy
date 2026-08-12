"""
Télécharge un visuel de maillot par club vers le dossier maillots/.

Source : dépôt GitHub "dwdyer/football-kit-icons" (licence CC0-1.0, domaine public),
des maillots SVG par motif/couleur — même principe que les applis de composition
gratuites (couleurs/motif réels du club, pas une photo officielle sous licence).
https://github.com/dwdyer/football-kit-icons

Chaque club de la saison est associé au motif qui se rapproche le plus de son vrai
maillot domicile. Si un id n'a pas de mapping ou que le téléchargement échoue, un
maillot uni de secours (gris) est généré localement en SVG pour ne jamais laisser
un joueur sans maillot affiché.

Usage : python scripts/download_maillots.py
"""

import sys
from pathlib import Path

try:
    import requests
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "requests", "-q"])
    import requests

BASE_DIR     = Path(__file__).parent.parent
MAILLOTS_DIR = BASE_DIR / "maillots"
MAILLOTS_DIR.mkdir(exist_ok=True)

RAW_BASE = "https://raw.githubusercontent.com/dwdyer/football-kit-icons/master/shirts"

# id club (api-sports, cf. CLUB_IDS dans enchere.html/admin.html) -> fichier du dépôt
# choisi pour se rapprocher du vrai maillot domicile du club (couleurs + motif)
CLUB_MAILLOTS = {
    85:   "band_blue_white_red",     # PSG
    81:   "sleeves_white_skyblue",   # OM
    91:   "diagonal_white_red",      # Monaco
    80:   "plain_white",             # OL
    79:   "sleeves_red_white",       # LOSC (Lille)
    84:   "plain_red",               # Nice
    116:  "stripes_yellow_red",      # Lens ("Sang et Or")
    94:   "stripes_red_black",       # Rennes
    95:   "plain_blue",              # Strasbourg
    108:  "sleeves_white_blue",      # Auxerre
    111:  "sleeves_navy_white",      # Le Havre
    96:   "plain_purple",            # Toulouse
    77:   "plain_black",             # Angers
    106:  "band_white_red",          # Brest
    97:   "plain_orange",            # Lorient
    114:  "stripes_red_blue",        # Paris FC
    110:  "sleeves_blue_navy",       # Troyes
    1298: "quarters_white_blue",     # Le Mans
}

FALLBACK_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512">
<path d="M170 40 L10 120 L60 220 L110 190 L110 480 L400 480 L400 190 L450 220 L500 120 L340 40
  L320 40 C320 75 292 100 256 100 C220 100 192 75 192 40 Z"
  fill="#9aa0a6" stroke="#5f6368" stroke-width="6"/>
</svg>"""


def main():
    print(f"{len(CLUB_MAILLOTS)} clubs à traiter vers {MAILLOTS_DIR}\n")

    for club_id, fname in CLUB_MAILLOTS.items():
        dest = MAILLOTS_DIR / f"{club_id}.svg"
        if dest.exists():
            print(f"  {club_id}.svg  déjà présent, skip")
            continue

        url = f"{RAW_BASE}/{fname}.svg"
        try:
            r = requests.get(url, timeout=10)
            if r.status_code == 200 and "<svg" in r.text.lower():
                dest.write_text(r.text, encoding="utf-8")
                print(f"  {club_id}.svg  OK  ({fname})")
            else:
                dest.write_text(FALLBACK_SVG, encoding="utf-8")
                print(f"  {club_id}.svg  ERREUR HTTP {r.status_code} -> maillot de secours")
        except Exception as e:
            dest.write_text(FALLBACK_SVG, encoding="utf-8")
            print(f"  {club_id}.svg  ERREUR : {e} -> maillot de secours")

    # maillot de secours générique pour tout club sans id connu
    default_dest = MAILLOTS_DIR / "default.svg"
    if not default_dest.exists():
        default_dest.write_text(FALLBACK_SVG, encoding="utf-8")
        print(f"\n  default.svg  créé (maillot de secours générique)")

    print("\nTerminé.")


if __name__ == "__main__":
    main()

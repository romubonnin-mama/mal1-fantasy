"""Compare Excel journée scores vs data.json for selected managers."""
import json
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))
from maj import (
    EXCEL_PATH, JOUEURS_CONFIG, JOUEURS_CONFIG_ANCIEN,
    lire_equipe, calc_pts, calc_tj_pts
)

import openpyxl

MANAGERS = ["BASTIEN", "FAB", "ADRIEN", "VINCENT", "ANTHONY"]

DATA_JSON = os.path.join(os.path.dirname(__file__), "..", "data.json")
EXCEL_PATH_LOCAL = EXCEL_PATH


def score_equipe(equipe):
    return sum(p["pts"] for pos_players in equipe.values() for p in pos_players)


def main():
    with open(DATA_JSON, encoding="utf-8") as f:
        data = json.load(f)

    print("Lecture Excel...")
    wb = openpyxl.load_workbook(EXCEL_PATH_LOCAL, data_only=True)

    for mgr in MANAGERS:
        print(f"\n=== {mgr} ===")
        total_xl = 0
        total_site = 0
        diffs = []

        for j in range(1, 35):
            js = str(j)
            if js not in wb.sheetnames:
                continue
            ws = wb[js]
            ancien = (j < 14)
            equipe_xl = lire_equipe(ws, mgr, ancien)
            score_xl = score_equipe(equipe_xl)

            score_site = data.get("historique", {}).get(js, {}).get(mgr, 0)

            total_xl += score_xl
            total_site += score_site

            diff = score_xl - score_site
            if diff != 0:
                diffs.append((j, score_xl, score_site, diff, equipe_xl))

        for j, xl, site, diff, equipe_xl in diffs:
            print(f"  J{j}: xl={xl:3d} site={site:3d} diff={diff:+d}")
            # Show player-level breakdown for differing journées
            for poste, players in equipe_xl.items():
                for p in players:
                    site_detail = (
                        data.get("detail_journees", {})
                            .get(str(j), {})
                            .get(mgr, {})
                            .get(poste, [])
                    )
                    site_p = next((x for x in site_detail if x["nom"] == p["nom"]), None)
                    site_pts = site_p["pts"] if site_p else "?"
                    if p["pts"] != site_pts:
                        print(
                            f"    {p['nom']} ({poste}) xl={p['pts']:3} site={site_pts}"
                            f"  cap={p['cap'] or '-'} tj={p['tj']}"
                            f"  bm={p['bm']['val']} cs={p['cs']['val']}"
                            f"  be={p['be']['val']} pma={p['pma']['val']}"
                            f"  pd={p['pd']['val']} cj={p['cj']['val']}"
                        )

        print(f"  TOTAL xl={total_xl} site={total_site} diff={total_xl-total_site:+d}")


if __name__ == "__main__":
    main()

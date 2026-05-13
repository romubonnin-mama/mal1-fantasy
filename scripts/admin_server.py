"""
Serveur local pour l'interface admin Ma L1.
Lance avec : python scripts/admin_server.py
Puis ouvre : http://localhost:8765
"""

import json
import os
import subprocess
import sys
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.parse import urlparse

BASE_DIR   = Path(__file__).parent.parent
DATA_DIR   = BASE_DIR / "data"
SCRIPTS_DIR = BASE_DIR / "scripts"
ADMIN_HTML = BASE_DIR / "admin.html"
PORT       = 8765

FILES = {
    "/api/roster":        DATA_DIR / "roster.json",
    "/api/player-ids":    DATA_DIR / "player_ids.json",
    "/api/lineups":       DATA_DIR / "lineups.json",
    "/api/corrections":   DATA_DIR / "corrections.json",
    "/api/manual-stats":  DATA_DIR / "manual_stats.json",
    "/api/data":          BASE_DIR  / "data.json",
    "/api/clubs":         BASE_DIR  / "clubs.json",
}

JOURNEE_FILES = {
    "/api/lineups/":      "lineups.json",
    "/api/corrections/":  "corrections.json",
    "/api/manual-stats/": "manual_stats.json",
}


def read_json(path: Path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def _reconstruct_player_stats(player: dict) -> dict:
    """Convertit un joueur de detail_journees en entrée manual_stats."""
    tj   = str(player.get("tj", "0"))
    bm   = (player.get("bm")   or {}).get("val", 0) or 0
    pd   = (player.get("pd")   or {}).get("val", 0) or 0
    cs   = (player.get("cs")   or {}).get("val", 0) or 0
    be   = (player.get("be")   or {}).get("val", 0) or 0
    bcsc = (player.get("bcsc") or {}).get("val", 0) or 0
    pm   = (player.get("pm")   or {}).get("val", 0) or 0
    pma  = (player.get("pma")  or {}).get("val", 0) or 0
    cj   = (player.get("cj")   or {}).get("val", 0) or 0
    cr   = (player.get("cr")   or {}).get("val", 0) or 0

    s = {}
    if tj == "M":
        s["full_match"] = True
    elif "-" in tj:
        parts = tj.split("-")
        try:
            s["entre_a"] = int(parts[0])
            s["fin_a"]   = int(parts[1])
        except ValueError:
            pass
    elif tj != "0":
        try:
            v = int(tj)
            if v > 0:
                s["sort_a"] = v
        except ValueError:
            pass

    if bm   > 0:  s["goals"]       = bm
    if pd   > 0:  s["assists"]      = pd
    if cs   > 0:  s["cs"]           = True
    if be   >= 3: s["be_malus"]     = True
    if bcsc > 0:  s["own_goals"]    = bcsc
    if pm   > 0:  s["pen_scored"]   = pm
    if pma  > 0:  s["pen_mm_saved"] = pma
    if cj   > 0:  s["yellow_cards"] = cj
    if cr   < 0:  s["red_card"]     = True

    return s


def _merge_manual_from_detail(journee: str, current: dict) -> dict:
    """
    Si data.json contient detail_journees[journee], reconstruit les stats
    manquantes et les fusionne avec current (les stats manuelles existantes
    ont priorité sur les données reconstituées).
    """
    data_path = BASE_DIR / "data.json"
    if not data_path.exists():
        return current

    try:
        site_data = read_json(data_path)
    except Exception:
        return current

    detail = site_data.get("detail_journees", {}).get(journee)
    if not detail:
        return current

    reconstructed = {}
    for manager, postes in detail.items():
        m_stats = {}
        for poste, players in postes.items():
            for p in players:
                nom    = p.get("nom", "")
                statut = p.get("statut", "")
                if statut == "r":
                    continue
                if statut == "A":
                    m_stats[nom] = {"absent": True}
                    continue
                m_stats[nom] = _reconstruct_player_stats(p)
        if m_stats:
            reconstructed[manager] = m_stats

    # Fusion : base = reconstruit, overlay = manual existant (priorité)
    merged = {}
    all_managers = set(list(reconstructed.keys()) + list(current.keys()))
    for manager in all_managers:
        base    = reconstructed.get(manager, {})
        overlay = current.get(manager, {})
        merged[manager] = {**base, **overlay}

    return merged


def write_json(path: Path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


class AdminHandler(BaseHTTPRequestHandler):

    def log_message(self, fmt, *args):
        print(f"  {self.address_string()} {fmt % args}")

    def send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def send_error_json(self, msg, status=400):
        self.send_json({"error": msg}, status)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        path = urlparse(self.path).path.rstrip("/")

        # Serve admin.html
        if path in ("", "/"):
            with open(ADMIN_HTML, "rb") as f:
                body = f.read()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        # Full-file endpoints
        if path in FILES:
            self.send_json(read_json(FILES[path]))
            return

        # Per-journée endpoints: /api/lineups/32  /api/manual-stats/32  etc.
        for prefix, fname in JOURNEE_FILES.items():
            if path.startswith(prefix):
                journee = path[len(prefix):]
                data = read_json(DATA_DIR / fname)
                result = data.get(journee, {})

                # Pour manual-stats : si la journée est vide ou incomplète,
                # pré-remplir depuis detail_journees de data.json
                if fname == "manual_stats.json":
                    result = _merge_manual_from_detail(journee, result)

                self.send_json(result)
                return

        self.send_error_json("Not found", 404)

    def do_POST(self):
        path = urlparse(self.path).path.rstrip("/")
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length) if length else b"{}"
        try:
            payload = json.loads(body)
        except json.JSONDecodeError:
            self.send_error_json("JSON invalide")
            return

        # Per-journée save
        for prefix, fname in JOURNEE_FILES.items():
            if path.startswith(prefix):
                journee = path[len(prefix):]
                data = read_json(DATA_DIR / fname)
                data[journee] = payload
                write_json(DATA_DIR / fname, data)
                self.send_json({"ok": True})
                return

        # Full-file save (roster, player-ids)
        for endpoint, fpath in FILES.items():
            if path == endpoint and endpoint in ("/api/roster", "/api/player-ids"):
                write_json(fpath, payload)
                self.send_json({"ok": True})
                return

        # Compute: calcule les points et met à jour data.json
        if path.startswith("/api/compute/"):
            journee = int(path[len("/api/compute/"):])
            try:
                sys.path.insert(0, str(SCRIPTS_DIR))
                import importlib
                import scoring
                import compute_journee
                importlib.reload(scoring)
                importlib.reload(compute_journee)
                result = compute_journee.compute(journee)

                # Sync Excel après le calcul
                try:
                    import export_excel
                    importlib.reload(export_excel)
                    export_excel.export_journee(journee, verbose=False)
                    result["excel"] = "ok"
                except Exception as exc_xl:
                    import traceback
                    print(f"[excel] ERREUR: {exc_xl}")
                    traceback.print_exc()
                    result["excel_warn"] = str(exc_xl)

                self.send_json(result)
            except Exception as e:
                self.send_error_json(str(e), 500)
            return

        # Git push
        if path == "/api/push":
            journee = payload.get("journee", "?")
            try:
                subprocess.run(["git", "add", "."],     cwd=BASE_DIR, check=True)
                r = subprocess.run(
                    ["git", "commit", "-m", f"Mise a jour J{journee}"],
                    cwd=BASE_DIR, capture_output=True
                )
                if r.returncode == 0 or b"nothing to commit" in r.stdout:
                    subprocess.run(["git", "push"], cwd=BASE_DIR, check=True)
                    self.send_json({"ok": True, "msg": f"J{journee} publié sur GitHub"})
                else:
                    self.send_json({"ok": False, "msg": r.stderr.decode()})
            except Exception as e:
                self.send_error_json(str(e), 500)
            return

        self.send_error_json("Not found", 404)


if __name__ == "__main__":
    print(f"Interface admin : http://localhost:{PORT}")
    print("   Ctrl+C pour arrêter.\n")
    server = HTTPServer(("localhost", PORT), AdminHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServeur arrêté.")

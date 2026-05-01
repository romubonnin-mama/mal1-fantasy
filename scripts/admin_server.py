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
                self.send_json(data.get(journee, {}))
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
                import compute_journee
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

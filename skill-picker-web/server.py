#!/usr/bin/env python3
"""
Cross-platform local web server for the Skill Picker.

Works on Windows, macOS, and Linux with nothing but the Python standard
library (no pip installs). Serves a small HTML/JS front end that lists
every skill from ../skills-manifest.json, lets the user pick a skill and
its inputs (using native OS file dialogs via tkinter), runs it, and shows
the captured output plus links to any result files it produced.

Run:
    python server.py            (or: py server.py   on Windows)

Then open the browser window it launches automatically, or go to
http://127.0.0.1:8765/ manually.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse, parse_qs

ROOT = Path(__file__).resolve().parent          # .../skills/skill-picker-web
SKILLS_ROOT = ROOT.parent                        # .../skills
MANIFEST_PATH = SKILLS_ROOT / "skills-manifest.json"
STATIC_DIR = ROOT / "static"
PORT = 8765

# tkinter dialogs must each run with their own short-lived Tk root. A lock
# keeps two browser tabs from popping dialogs on top of each other.
_dialog_lock = threading.Lock()


def load_manifest() -> dict:
    with open(MANIFEST_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def open_native_dialog(kind: str, multiple: bool = False) -> dict:
    """Open a native OS file/folder dialog on the machine running the server."""
    import tkinter as tk
    from tkinter import filedialog

    with _dialog_lock:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        try:
            if kind == "folder":
                path = filedialog.askdirectory(parent=root)
                result = {"path": path} if path else {"path": ""}
            elif kind == "save":
                path = filedialog.asksaveasfilename(parent=root)
                result = {"path": path} if path else {"path": ""}
            else:  # open
                filetypes = [
                    ("Data files", "*.csv *.xlsx *.xls *.docx *.pptx *.pdf"),
                    ("All files", "*.*"),
                ]
                if multiple:
                    paths = filedialog.askopenfilenames(parent=root, filetypes=filetypes)
                    result = {"paths": list(paths)}
                else:
                    path = filedialog.askopenfilename(parent=root, filetypes=filetypes)
                    result = {"path": path} if path else {"path": ""}
        finally:
            root.destroy()
        return result


def open_with_os_default(path: str) -> None:
    if sys.platform.startswith("win"):
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def build_command(skill: dict, params: dict) -> tuple[list[str], Path]:
    """Mirror the terminal launcher's argument-building logic."""
    skill_dir = SKILLS_ROOT / skill["folder"]
    entry_path = skill_dir / skill["entry"]
    if not entry_path.exists():
        raise FileNotFoundError(f"Entry script not found: {entry_path}")

    args: list[str] = []
    for p in skill.get("params", []):
        flag = p["flag"]
        value = params.get(flag)
        ptype = p.get("type")

        if ptype == "switch":
            if value:
                args.append(flag)
        elif ptype == "repeat-csv":
            if value:
                items = [v.strip() for v in str(value).split(",") if v.strip()]
                for item in items:
                    args.extend([flag, item])
        elif ptype == "path-list":
            if value:
                # value is a ';'-joined list of paths collected from a multi-select dialog
                items = [v for v in str(value).split(";") if v]
                for item in items:
                    args.extend([flag, item])
        else:
            if value not in (None, ""):
                args.extend([flag, str(value)])

    if skill["runner"] == "python":
        cmd = [sys.executable, str(entry_path)] + args
    else:
        pwsh = "pwsh" if _command_exists("pwsh") else "powershell"
        cmd = [pwsh, "-ExecutionPolicy", "Bypass", "-File", str(entry_path)] + args

    return cmd, skill_dir


def _command_exists(name: str) -> bool:
    from shutil import which
    return which(name) is not None


def run_skill(skill: dict, params: dict) -> dict:
    skill_dir = SKILLS_ROOT / skill["folder"]
    start_time = time.time()
    try:
        cmd, cwd = build_command(skill, params)
    except FileNotFoundError as e:
        return {"success": False, "output": str(e), "resultFiles": []}

    try:
        proc = subprocess.run(
            cmd,
            cwd=str(cwd),
            capture_output=True,
            text=True,
            timeout=900,
        )
        output = proc.stdout + ("\n" + proc.stderr if proc.stderr else "")
        success = proc.returncode == 0
    except subprocess.TimeoutExpired:
        output = "Timed out after 15 minutes."
        success = False
    except Exception as e:  # noqa: BLE001
        output = f"Failed to run: {e}"
        success = False

    result_files = []
    try:
        for f in skill_dir.rglob("*"):
            if f.is_file() and f.stat().st_mtime >= start_time - 1:
                result_files.append(str(f.relative_to(skill_dir)))
    except Exception:
        pass

    return {"success": success, "output": output.strip(), "resultFiles": sorted(result_files)}


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):  # quiet default logging
        pass

    def _send_json(self, payload: dict, status: int = 200):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, path: Path, content_type: str):
        data = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        parsed = urlparse(self.path)
        route = parsed.path

        if route == "/" or route == "/index.html":
            self._send_file(STATIC_DIR / "index.html", "text/html; charset=utf-8")
        elif route == "/app.js":
            self._send_file(STATIC_DIR / "app.js", "application/javascript; charset=utf-8")
        elif route == "/style.css":
            self._send_file(STATIC_DIR / "style.css", "text/css; charset=utf-8")
        elif route == "/api/skills":
            self._send_json(load_manifest())
        elif route == "/api/open-file":
            qs = parse_qs(parsed.query)
            skill_folder = qs.get("skill", [""])[0]
            rel = qs.get("path", [""])[0]
            target = (SKILLS_ROOT / skill_folder / rel).resolve()
            skill_dir = (SKILLS_ROOT / skill_folder).resolve()
            if skill_dir in target.parents or target == skill_dir:
                open_with_os_default(str(target))
                self._send_json({"ok": True})
            else:
                self._send_json({"ok": False, "error": "Path outside skill folder"}, 400)
        else:
            self._send_json({"error": "Not found"}, 404)

    def do_POST(self):
        parsed = urlparse(self.path)
        route = parsed.path
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length) if length else b""
        try:
            data = json.loads(raw) if raw else {}
        except json.JSONDecodeError:
            data = {}

        if route == "/api/browse":
            kind = data.get("kind", "open")
            multiple = bool(data.get("multiple", False))
            try:
                result = open_native_dialog(kind, multiple)
            except Exception as e:  # noqa: BLE001
                result = {"error": str(e)}
            self._send_json(result)
        elif route == "/api/run":
            manifest = load_manifest()
            skill_id = data.get("skillId")
            params = data.get("params", {})
            skill = next((s for s in manifest["skills"] if s["id"] == skill_id), None)
            if not skill:
                self._send_json({"success": False, "output": "Unknown skill id", "resultFiles": []}, 404)
                return
            result = run_skill(skill, params)
            self._send_json(result)
        else:
            self._send_json({"error": "Not found"}, 404)


def main():
    server = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    url = f"http://127.0.0.1:{PORT}/"
    print(f"Skill Picker running at {url}")
    threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()

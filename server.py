#!/usr/bin/env python3
"""Flask web server frontend for the AI Document Revision Agent."""

import json
import os
import queue
import subprocess
import sys
import threading
import time
import webbrowser
from pathlib import Path

try:
    from flask import Flask, Response, jsonify, request, send_from_directory
except ImportError:
    print("\n  ERROR: Flask is not installed.")
    print("  Run:  pip install flask\n")
    sys.exit(1)

try:
    from agent import run_pipeline
except ImportError as e:
    print(f"\n  ERROR: Could not import agent.py: {e}\n")
    sys.exit(1)

_UI = Path(__file__).parent / "ui"
app = Flask(__name__, static_folder=str(_UI), static_url_path="")
_jobs: dict = {}

# Hide technical [bracketed] debug lines from the live log unless explicitly on.
_SHOW_DEBUG = os.getenv("DOCREVISE_DEBUG") == "1"


# ── Serve UI ──────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory(str(_UI), "index.html")


# ── Native file/folder dialogs ────────────────────────────────────────────────

def _dialog(script: str) -> str:
    """Run a tkinter dialog in a subprocess to avoid main-thread conflicts."""
    kwargs: dict = {}
    if sys.platform == "win32":
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    try:
        r = subprocess.run(
            [sys.executable, "-c", script],
            capture_output=True, text=True, timeout=300, **kwargs,
        )
        return r.stdout.strip()
    except Exception:
        return ""


@app.route("/api/browse/file", methods=["POST"])
def api_browse_file():
    path = _dialog(
        "from tkinter import filedialog,Tk\n"
        "r=Tk();r.withdraw();r.wm_attributes('-topmost',1)\n"
        "p=filedialog.askopenfilename(filetypes=[('Documents','*.txt *.docx *.pdf'),('All','*.*')])\n"
        "r.destroy();print(p or '',end='')"
    )
    return jsonify({"path": path})


@app.route("/api/browse/folder", methods=["POST"])
def api_browse_folder():
    path = _dialog(
        "from tkinter import filedialog,Tk\n"
        "r=Tk();r.withdraw();r.wm_attributes('-topmost',1)\n"
        "p=filedialog.askdirectory()\n"
        "r.destroy();print(p or '',end='')"
    )
    return jsonify({"path": path})


# ── Run pipeline ──────────────────────────────────────────────────────────────

@app.route("/api/run", methods=["POST"])
def api_run():
    data    = request.get_json(force=True) or {}
    docs    = (data.get("docs")    or "").strip()
    prompt  = (data.get("prompt")  or "").strip()
    gallery = (data.get("gallery") or "").strip()

    errors = []
    if not docs:
        errors.append("Documents folder is required")
    elif not os.path.isdir(docs):
        errors.append("Documents folder not found")
    if not prompt:
        errors.append("Change order file is required")
    elif not os.path.isfile(prompt):
        errors.append("Change order file not found")
    if gallery and not os.path.isdir(gallery):
        errors.append("Image gallery folder not found")

    if errors:
        return jsonify({"error": "; ".join(errors)}), 400

    job_id = str(time.time_ns())
    q: queue.Queue = queue.Queue()
    _jobs[job_id] = {"queue": q}

    def _log(msg: str):
        if not _SHOW_DEBUG and msg.lstrip().startswith("["):
            return
        q.put({"type": "log", "msg": msg})

    def _done(ok: bool, msg: str):
        q.put({"type": "done", "ok": ok, "msg": msg})

    def _view(evt: dict):
        q.put({"type": "view", **evt})

    threading.Thread(
        target=run_pipeline,
        args=(prompt, docs, gallery, _log, _done),
        kwargs={"view": _view},
        daemon=True,
    ).start()

    return jsonify({"job_id": job_id})


@app.route("/api/stream/<job_id>")
def api_stream(job_id: str):
    if job_id not in _jobs:
        return "Job not found", 404

    def _generate():
        q = _jobs[job_id]["queue"]
        try:
            while True:
                try:
                    evt = q.get(timeout=25)
                    yield f"data: {json.dumps(evt)}\n\n"
                    if evt.get("type") == "done":
                        break
                except queue.Empty:
                    yield 'data: {"type":"ping"}\n\n'
        finally:
            # Runs on normal finish and on client disconnect (GeneratorExit), so
            # repeated runs don't accumulate job entries for the life of the process.
            _jobs.pop(job_id, None)

    return Response(
        _generate(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


# ── Startup ───────────────────────────────────────────────────────────────────

def _pick_port(start: int = 7479) -> int:
    import socket
    for p in range(start, start + 20):
        with socket.socket() as s:
            try:
                s.bind(("127.0.0.1", p))
                return p
            except OSError:
                continue
    return start


if __name__ == "__main__":
    port = _pick_port()
    url  = f"http://localhost:{port}"
    threading.Timer(1.5, lambda: webbrowser.open(url)).start()
    print(f"\n  DocRevise — AI Document Revision Agent")
    print(f"  Open:  {url}\n")
    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)

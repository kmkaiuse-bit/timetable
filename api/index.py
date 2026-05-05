import base64
import os
import sys

# Allow importing timetable_scheduler from the parent directory
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from flask import Flask, jsonify, request, send_from_directory
from timetable_scheduler import run_from_bytes

_PUBLIC = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "public")

app = Flask(__name__, static_folder=_PUBLIC)


@app.route("/", defaults={"path": ""})
@app.route("/<path:path>")
def serve_ui(path):
    # Try exact file first (e.g. CSS, JS assets)
    if path:
        full = os.path.join(_PUBLIC, path)
        if os.path.isfile(full):
            return send_from_directory(_PUBLIC, path)

    # Serve index.html for all other routes
    html = os.path.join(_PUBLIC, "index.html")
    if os.path.isfile(html):
        with open(html, encoding="utf-8") as f:
            return f.read(), 200, {"Content-Type": "text/html; charset=utf-8"}

    # Last resort: debug info
    return (
        f"<pre>public dir: {_PUBLIC}\n"
        f"exists: {os.path.isdir(_PUBLIC)}\n"
        f"files: {os.listdir(_PUBLIC) if os.path.isdir(_PUBLIC) else 'N/A'}</pre>"
    ), 404


@app.route("/api/schedule", methods=["POST"])
def schedule():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400
    if not (f.filename or "").lower().endswith(".xlsx"):
        return jsonify({"error": "Please upload an .xlsx file"}), 400

    try:
        output_bytes, stats = run_from_bytes(f.read())
        return jsonify({
            "stats": stats,
            "file":  base64.b64encode(output_bytes).decode("utf-8"),
        })
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500

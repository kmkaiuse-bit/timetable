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
    """Serve static UI files (for local dev; Vercel serves public/ directly)."""
    if path and os.path.exists(os.path.join(_PUBLIC, path)):
        return send_from_directory(_PUBLIC, path)
    return send_from_directory(_PUBLIC, "index.html")


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

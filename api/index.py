import base64
import os
import sys

# Allow importing timetable_scheduler from the parent directory
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from flask import Flask, jsonify, request, send_from_directory
from timetable_scheduler import run_from_bytes

# index.html lives in the same directory as this file (api/)
_HERE = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)


@app.route("/", defaults={"path": ""})
@app.route("/<path:path>")
def serve_ui(path):
    html = os.path.join(_HERE, "index.html")
    if os.path.isfile(html):
        with open(html, encoding="utf-8") as f:
            return f.read(), 200, {"Content-Type": "text/html; charset=utf-8"}
    return f"<pre>index.html not found at {html}</pre>", 404


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

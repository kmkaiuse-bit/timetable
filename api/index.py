import base64
import os
import sys

# timetable_scheduler.py lives in the same api/ directory
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flask import Flask, jsonify, request, send_from_directory, send_file, Response
from timetable_scheduler import run_from_bytes, run_v4_from_bytes

# index.html lives in the same directory as this file (api/)
_HERE = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)


@app.after_request
def add_security_headers(response):
    # Allow inline scripts/styles for this internal tool
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; "
        "script-src 'self' 'unsafe-inline'; "
        "style-src 'self' 'unsafe-inline';"
    )
    return response


@app.route("/", defaults={"path": ""})
@app.route("/<path:path>")
def serve_ui(path):
    # Serve v4.html for root and /v4; index.html only for explicit /v3
    p = path.strip("/")
    filename = "index.html" if p == "v3" else "v4.html"
    html = os.path.join(_HERE, filename)
    if os.path.isfile(html):
        with open(html, encoding="utf-8") as f:
            return f.read(), 200, {"Content-Type": "text/html; charset=utf-8"}
    return f"<pre>{filename} not found at {html}</pre>", 404


@app.route("/api/template")
def download_template():
    path = os.path.join(_HERE, "template.xlsx")
    return send_file(path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True, download_name="Planning for Timetable.xlsx")

@app.route("/api/template/v4")
def download_template_v4():
    path = os.path.join(_HERE, "template_v4.xlsx")
    return send_file(path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True, download_name="Planning for Timetable V4.xlsx")


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


@app.route("/api/download", methods=["POST"])
def download_edited():
    from io import BytesIO
    from timetable_scheduler import write_output_fast
    data = request.get_json(silent=True) or {}
    results = data.get("results", [])
    if not results:
        return jsonify({"error": "No results provided"}), 400
    try:
        output_bytes = write_output_fast(results)
        return send_file(
            BytesIO(output_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="Timetable_Output.xlsx",
        )
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/schedule/v4", methods=["POST"])
def schedule_v4():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400
    if not (f.filename or "").lower().endswith(".xlsx"):
        return jsonify({"error": "Please upload an .xlsx file"}), 400
    try:
        output_bytes, stats = run_v4_from_bytes(f.read())
        return jsonify({
            "stats": stats,
            "file":  base64.b64encode(output_bytes).decode("utf-8"),
        })
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500

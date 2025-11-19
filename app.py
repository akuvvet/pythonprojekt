from flask import Flask, render_template, request, jsonify, send_from_directory, redirect, url_for, session
import os
from mieten import fuehre_mietabgleich_durch
import traceback
from datetime import timedelta, datetime

UPLOAD_FOLDER = "uploads"
RESULTS_FOLDER = "results"
ADMIN_EMAIL = "akuvvet@gmail.com"
ADMIN_PASSWORD = "AKuvvet"

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["RESULTS_FOLDER"] = RESULTS_FOLDER
app.secret_key = os.environ.get("SECRET_KEY", "please-change-me-very-secret")
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(minutes=60)

# Upload-Verzeichnis erzeugen
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULTS_FOLDER, exist_ok=True)


# --- Auth Schutz global ---
@app.before_request
def require_login():
    open_endpoints = {"login", "static"}
    if request.endpoint in open_endpoints or request.endpoint is None:
        return
    if not session.get("user_email"):
        return redirect(url_for("login"))
    # Session-Timeout (60 Minuten Inaktivität)
    last = session.get("last_activity")
    now_ts = int(datetime.utcnow().timestamp())
    if last and (now_ts - int(last)) > 60 * 60:
        session.clear()
        return redirect(url_for("login"))
    session["last_activity"] = now_ts


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form.get("email", "").strip()
        password = request.form.get("password", "")
        if email.lower() == ADMIN_EMAIL.lower() and password == ADMIN_PASSWORD:
            session["user_email"] = email
            session.permanent = True
            session["last_activity"] = int(datetime.utcnow().timestamp())
            return redirect(url_for("index"))
        return render_template("login.html", error="Zugangsdaten ungültig.")
    # GET
    if session.get("user_email"):
        return redirect(url_for("index"))
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
def index():
    return render_template("upload.html")


@app.route("/process", methods=["POST"])
def process():
    excel = request.files["excel"]
    konto_file = request.files["konto"]

    if not excel or not konto_file:
        return jsonify({"status": "error", "message": "Bitte Excel (Mieter) und Excel (Kontoauszug) hochladen."}), 400

    excel_path = os.path.join(UPLOAD_FOLDER, excel.filename)
    konto_path = os.path.join(UPLOAD_FOLDER, konto_file.filename)

    excel.save(excel_path)
    konto_file.save(konto_path)

    # Script aufrufen → Übergib Datei-Pfade
    try:
        result_path = fuehre_mietabgleich_durch(excel_path, konto_path)
        if not result_path or not os.path.exists(result_path):
            return jsonify({"status": "error", "message": "Ergebnisdatei wurde nicht erstellt."}), 500

        download_name = os.path.basename(result_path)
        return jsonify({
            "status": "ok",
            "message": "Mietabgleich abgeschlossen",
            "download": f"/results/{download_name}"
        })
    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e),
            "trace": traceback.format_exc()
        })


@app.route("/results/<path:filename>")
def download_result(filename):
    file_path = os.path.join(RESULTS_FOLDER, filename)
    if not os.path.exists(file_path):
        return jsonify({"status": "error", "message": "Datei nicht gefunden"}), 404
    return send_from_directory(
        RESULTS_FOLDER,
        filename,
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        download_name=filename
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)

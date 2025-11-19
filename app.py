from flask import Flask, render_template, request, jsonify, send_from_directory
import os
from mieten import fuehre_mietabgleich_durch
import traceback

UPLOAD_FOLDER = "uploads"
RESULTS_FOLDER = "results"

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["RESULTS_FOLDER"] = RESULTS_FOLDER

# Upload-Verzeichnis erzeugen
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULTS_FOLDER, exist_ok=True)


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

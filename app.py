from flask import Flask, render_template, request, redirect
import os
from mieten import fuehre_mietabgleich_durch

UPLOAD_FOLDER = "uploads"

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# Upload-Verzeichnis erzeugen
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route("/")
def index():
    return render_template("upload.html")


@app.route("/process", methods=["POST"])
def process():
    excel = request.files["excel"]
    csv_file = request.files["csv"]

    if not excel or not csv_file:
        return "Bitte Excel und CSV hochladen."

    excel_path = os.path.join(UPLOAD_FOLDER, excel.filename)
    csv_path = os.path.join(UPLOAD_FOLDER, csv_file.filename)

    excel.save(excel_path)
    csv_file.save(csv_path)

    # Script aufrufen → Übergib Datei-Pfade
    try:
        fuehre_mietabgleich_durch(excel_path, csv_path)
        return "✔️ Mietabgleich erfolgreich ausgeführt!"
    except Exception as e:
        return f"Fehler: {e}"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)

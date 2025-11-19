from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
import os
from mieten import fuehre_mietabgleich_durch  # Deine Mietabgleich-Funktion

# --- App initialisieren ---
app = FastAPI(title="Mieten-Abgleich")

# Ordner erstellen, falls nicht vorhanden
os.makedirs("uploads", exist_ok=True)
os.makedirs("results", exist_ok=True)

# Statische Dateien bereitstellen (optional für HTML/CSS)
app.mount("/static", StaticFiles(directory="static"), name="static")


# --- HTML-Startseite ---
@app.get("/", response_class=HTMLResponse)
async def index():
    return """
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <title>Mieten-Abgleich</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
</head>
<body class="bg-light">
<div class="container py-5">
    <h1 class="mb-4">Mieten-Abgleich</h1>
    <form id="uploadForm" enctype="multipart/form-data">
        <div class="mb-3">
            <label class="form-label">Excel-Datei (Mieter.xlsx)</label>
            <input type="file" name="excel" class="form-control" required>
        </div>
        <div class="mb-3">
            <label class="form-label">CSV Kontoauszug</label>
            <input type="file" name="csv" class="form-control" required>
        </div>
        <button type="submit" class="btn btn-primary">Starten</button>
    </form>
    <div id="result" class="mt-4"></div>
</div>

<script>
const form = document.getElementById('uploadForm');
form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const formData = new FormData(form);
    const resultDiv = document.getElementById('result');
    resultDiv.innerHTML = "Mietabgleich läuft...";

    const response = await fetch("/process", { method: "POST", body: formData });
    const data = await response.json();

    if (data.status === "ok") {
        resultDiv.innerHTML = `<div class="alert alert-success">Mietabgleich erfolgreich! <a href="${data.download}" class="alert-link" download>Hier herunterladen</a></div>`;
    } else {
        resultDiv.innerHTML = `<div class="alert alert-danger">Fehler: ${data.message}</div>`;
    }
});
</script>
</body>
</html>
"""


# --- Mietabgleich starten ---
@app.post("/process")
async def process_files(excel: UploadFile = File(...), csv: UploadFile = File(...)):
    try:
        # Dateien speichern
        excel_path = f"uploads/{excel.filename}"
        csv_path = f"uploads/{csv.filename}"
        with open(excel_path, "wb") as f:
            f.write(excel.file.read())
        with open(csv_path, "wb") as f:
            f.write(csv.file.read())

        # Mietabgleich ausführen
        output_file = f"results/mieten_abgleich.xlsx"
        fuehre_mietabgleich_durch(excel_path, csv_path, output_file)

        # Sauberes JSON zurückgeben
        logs = []
        logs.append("Mietabgleich erfolgreich")
        return JSONResponse({
            "status": "ok",
            "message": "Mietabgleich abgeschlossen",
            "download": "/results/mieten_abgleich.xlsx",
            "logs": logs
        })
    except Exception as e:
        return JSONResponse({"status": "error", "message": str(e)})


# --- Download-Endpunkt ---
@app.get("/results/{filename}")
def download_file(filename: str):
    file_path = f"results/{filename}"
    if os.path.exists(file_path):
        return FileResponse(
            file_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=filename
        )
    return JSONResponse({"status": "error", "message": "Datei nicht gefunden"})

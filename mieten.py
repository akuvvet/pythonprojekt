import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from openpyxl.comments import Comment # Import für Kommentarfunktion
import os
import re

# --- Konfiguration der Spalten und Monatszuordnung ---

# Basisspalten für Miete (Betrag) und Datum in Mieter.xlsx (Buchstaben der Spalten)
MONATS_ZUORDNUNG = {
    "Jan": ["D", "E"],
    "Feb": ["F", "G"],
    "Mrz": ["H", "I"],
    "Apr": ["J", "K"],
    "Mai": ["L", "M"],
    "Jun": ["N", "O"],
    "Jul": ["P", "Q"],
    "Aug": ["R", "S"],
    "Sep": ["T", "U"],
    "Okt": ["V", "W"],
    "Nov": ["X", "Y"],
    "Dez": ["Z", "AA"],
}

# Deutsche Monatsnamen für die Suche im Verwendungszweck (Regex)
MONATS_NAMENS_MAPPING = {
    "Januar": "Jan", "Jan": "Jan",
    "Februar": "Feb", "Feb": "Feb",
    "März": "Mrz", "Marz": "Mrz", "Mrz": "Mrz",
    "April": "Apr", "Apr": "Apr",
    "Mai": "Mai",
    "Juni": "Jun", "Jun": "Jun",
    "Juli": "Jul", "Jul": "Jul",
    "August": "Aug", "Aug": "Aug",
    "September": "Sep", "Sep": "Sep", "Sept": "Sep",
    "Oktober": "Okt", "Okt": "Okt",
    "November": "Nov", "Nov": "Nov",
    "Dezember": "Dez", "Dez": "Dez",
}

# Excel-Blattname und Spalten
BLATTNAME = "haeselerstr"
MIETER_SPALTE_LETTER = "A" # Eigentümer/Mieter

# CSV Spaltennamen
CSV_DATUM_SPALTE = "Buchungstag"
CSV_VERWENDUNGSZWECK_SPALTE = "Verwendungszweck"
CSV_ABSENDER_SPALTE = "Beguenstigter/Zahlungspflichtiger"
CSV_BETRAG_SPALTE = "Betrag"


def fuehre_mietabgleich_durch(excel_pfad, csv_pfad):



def fuehre_mietabgleich_durch():
    """Hauptfunktion zur Durchführung des Mietabgleichs."""
    print("--- 🏠 Starte Mietabgleich (Akkumulation/Kommentar-Logik aktiv) ---")
    
    # 1. Dateipfade abfragen
    excel_pfad = waehle_datei("Wählen Sie die Mieter.xlsx Datei", "excel")
    if not excel_pfad: return

    csv_pfad = waehle_datei("Wählen Sie die Kontoauszug CSV Datei", "csv")
    if not csv_pfad: return

    print(f"\nExcel-Datei: {os.path.basename(excel_pfad)}")
    print(f"CSV-Datei: {os.path.basename(csv_pfad)}\n")

    # 2. Excel-Datei mit openpyxl und pandas einlesen
    try:
        # Lade die Arbeitsmappe mit openpyxl für das Schreiben (behält Formatierung)
        workbook = load_workbook(excel_pfad)
        if BLATTNAME not in workbook.sheetnames:
            print(f"FEHLER: Das Blatt '{BLATTNAME}' wurde in der Excel-Datei nicht gefunden.")
            return
            
        worksheet = workbook[BLATTNAME]
        
        # Lese die Excel-Daten mit Pandas für die Mietersuche (Header in Zeile 1)
        df_mieter = pd.read_excel(excel_pfad, sheet_name=BLATTNAME, header=0, na_filter=False)
        mieter_col_name = df_mieter.columns[column_index_from_string(MIETER_SPALTE_LETTER) - 1]
        
        print(f"Mieter-Datei ({BLATTNAME}) erfolgreich eingelesen.")
    except Exception as e:
        print(f"FEHLER beim Lesen/Laden der Excel-Datei: {e}")
        return

    # 3. CSV-Kontoauszug einlesen (robuster Block)
    df_konto = None
    separators = [';', ',', '\t']
    encodings = ['iso-8859-1', 'utf-8', 'latin-1'] 
    
    for sep in separators:
        if df_konto is not None: break
        for enc in encodings:
            try:
                temp_df = pd.read_csv(csv_pfad, sep=sep, encoding=enc, na_filter=False)
                required_cols = [CSV_BETRAG_SPALTE, CSV_DATUM_SPALTE, CSV_ABSENDER_SPALTE, CSV_VERWENDUNGSZWECK_SPALTE]
                if all(col in temp_df.columns for col in required_cols):
                    df_konto = temp_df
                    print(f"Kontoauszug erfolgreich eingelesen (Separator: '{sep}', Kodierung: '{enc}').")
                    break 
            except Exception:
                continue
                
    if df_konto is None:
        print(f"FEHLER: Die CSV-Datei konnte nicht gelesen oder die benötigten Spalten nicht gefunden werden.")
        return

    # --- Datenbereinigung und -vorbereitung der CSV ---
    try:
        df_konto[CSV_BETRAG_SPALTE] = df_konto[CSV_BETRAG_SPALTE].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
        df_konto[CSV_BETRAG_SPALTE] = pd.to_numeric(df_konto[CSV_BETRAG_SPALTE], errors='coerce')
        df_konto[CSV_DATUM_SPALTE] = pd.to_datetime(df_konto[CSV_DATUM_SPALTE], format='%d.%m.%Y', errors='coerce')
    except Exception as e:
        print(f"FEHLER bei der Konvertierung von Betrag/Datum in der CSV: {e}")
        return

    # 4. Abgleich und Eintragung
    eintraege_gefunden = 0
    monat_suchmuster = "|".join(re.escape(m) for m in MONATS_NAMENS_MAPPING.keys())

    for index, mieter_reihe in df_mieter.iterrows():
        excel_row_num = index + 2 # Zeilennummer in Excel (Header=1 -> Daten ab Zeile 2)
        mieter_name = mieter_reihe[mieter_col_name]
        
        if not mieter_name:
            continue

        gefilterte_buchungen = df_konto[
            (df_konto[CSV_ABSENDER_SPALTE].astype(str).str.contains(str(mieter_name), case=False, regex=True)) & 
            (df_konto[CSV_VERWENDUNGSZWECK_SPALTE].astype(str).str.contains('Miete', case=False, regex=True))
        ]

        if gefilterte_buchungen.empty:
            continue
            
        print(f"\n-> Verarbeite Mieter: **{mieter_name}** (Zeile {excel_row_num})")
            
        for _, buchung in gefilterte_buchungen.iterrows():
            buchungs_datum = buchung[CSV_DATUM_SPALTE]
            betrag = buchung[CSV_BETRAG_SPALTE]
            verwendungszweck = buchung[CSV_VERWENDUNGSZWECK_SPALTE]
            
            if pd.isna(buchungs_datum) or pd.isna(betrag):
                continue
            
            # --- 4.1. Bestimme den Zielmonat nach Priorität ---
            
            ziel_monat_kuerzel = None
            
            # 1. Priorität: Expliziter Monatsname im Verwendungszweck?
            match = re.search(monat_suchmuster, verwendungszweck, re.IGNORECASE)
            
            if match:
                gefundener_monat = match.group(0)
                ziel_monat_kuerzel = MONATS_NAMENS_MAPPING.get(gefundener_monat.capitalize(), None)
                if ziel_monat_kuerzel:
                     print(f"   Monatsname '{gefundener_monat}' im VWZ gefunden. Priorität: **{ziel_monat_kuerzel}**.")
            
            # 2. Priorität: Datumsverschiebe-Logik, WENN kein expliziter Monat gefunden wurde
            if ziel_monat_kuerzel is None:
                if buchungs_datum.day > 25:
                    naechster_monat_datum = buchungs_datum + pd.DateOffset(months=1)
                    monat_nr_final = naechster_monat_datum.month
                    ziel_monat_kuerzel = list(MONATS_ZUORDNUNG.keys())[monat_nr_final - 1]
                    print(f"   Kein Monatsname gefunden. Buchungstag {buchungs_datum.day}. > 25. Buche in den **nächsten Monat ({ziel_monat_kuerzel})**.")
                else:
                    monat_nr_final = buchungs_datum.month
                    ziel_monat_kuerzel = list(MONATS_ZUORDNUNG.keys())[monat_nr_final - 1]
                    print(f"   Kein Monatsname gefunden. Buchungstag {buchungs_datum.day}. <= 25. Buche in den Monat der Buchung: **{ziel_monat_kuerzel}**.")

            # --- 4.2. Eintragung mit openpyxl (mit Prüf- und Akkumulationslogik) ---
            
            ziel_betrag_spalte_letter = MONATS_ZUORDNUNG[ziel_monat_kuerzel][0]
            ziel_datum_spalte_letter = MONATS_ZUORDNUNG[ziel_monat_kuerzel][1]
            
            betrag_cell_ref = f"{ziel_betrag_spalte_letter}{excel_row_num}"
            datum_cell_ref = f"{ziel_datum_spalte_letter}{excel_row_num}"

            # Lese bestehende Werte aus Excel
            existing_betrag = worksheet[betrag_cell_ref].value
            existing_datum_cell = worksheet[datum_cell_ref]
            existing_datum = existing_datum_cell.value
            
            # Neue Werte formatieren
            new_betrag_val = betrag # float
            new_datum_str = buchungs_datum.strftime('%d.%m.')
            new_betrag_str_anzeige = f"{new_betrag_val:.2f}".replace('.', ',')

            # Vorbereitung für den Datumsvergleich (bestehendes Datum in String-Format)
            # Konvertiere Datum/datetime-Objekte in String TT.MM.
            existing_datum_str = ""
            if existing_datum:
                if isinstance(existing_datum, pd.Timestamp): # Openpyxl kann Pandas/Datetime-Objekte zurückgeben
                    existing_datum_str = existing_datum.strftime('%d.%m.')
                else:
                    existing_datum_str = str(existing_datum).strip()
            
            # Float-Vergleich mit Toleranz (Toleranz: 0.01 EUR)
            is_same_amount = existing_betrag is not None and abs(existing_betrag - new_betrag_val) < 0.01
            
            if existing_betrag is None or existing_betrag == "":
                # Fall 1: Zelle ist leer -> Normales Eintragen
                worksheet[betrag_cell_ref] = new_betrag_val
                worksheet[datum_cell_ref] = new_datum_str
                worksheet[datum_cell_ref].comment = None 
                print(f"   Eintrag für Monat **{ziel_monat_kuerzel}** NEU eingetragen: {new_betrag_str_anzeige} EUR am {new_datum_str}")
                eintraege_gefunden += 1

            elif is_same_amount and existing_datum_str == new_datum_str:
                # Fall 2: Exaktes Duplikat -> Ignorieren
                print(f"   Eintrag für Monat **{ziel_monat_kuerzel}** ist ein Exakt-Duplikat (Betrag/Datum). **Ignoriert.**")

            else:
                # Fall 3: Zelle gefüllt, aber es ist keine exakte Kopie -> Akkumulation erforderlich
                
                # 1. Neuen Gesamtbetrag berechnen
                new_total_betrag = existing_betrag + new_betrag_val
                new_total_betrag_anzeige = f"{new_total_betrag:.2f}".replace('.', ',')

                # 2. Kommentar-Vorbereitung (inkl. aller Details)
                current_comment_text = existing_datum_cell.comment.text if existing_datum_cell.comment else None

                if current_comment_text:
                    # Wenn bereits ein Kommentar existiert (d.h. es wurde schon einmal akkumuliert)
                    comment_text = current_comment_text + f"\n+ {new_datum_str}: {new_betrag_str_anzeige} EUR"
                else:
                    # Wenn kein Kommentar existiert, erstelle einen neuen, der die bestehende Zahlung und die neue enthält
                    existing_betrag_anzeige = f"{existing_betrag:.2f}".replace('.', ',')
                    
                    comment_text = f"Urspr.: {existing_datum_str}: {existing_betrag_anzeige} EUR"
                    comment_text += f"\n+ Neu: {new_datum_str}: {new_betrag_str_anzeige} EUR"
                    
                # 3. Zelle aktualisieren (Gesamtbetrag und aktuellstes Datum)
                worksheet[betrag_cell_ref] = new_total_betrag # Speichere den akkumulierten Betrag als Zahl
                worksheet[datum_cell_ref] = new_datum_str 
                
                # 4. Kommentar zur Datum-Zelle hinzufügen
                comment = Comment(comment_text, "Automatisches Update")
                worksheet[datum_cell_ref].comment = comment

                print(f"   Eintrag für Monat **{ziel_monat_kuerzel}** AKKUMULIERT: Gesamt {new_total_betrag_anzeige} EUR. Details im Kommentar.")
                eintraege_gefunden += 1

    # 5. Speichern der geänderten Excel-Datei
    if eintraege_gefunden > 0:
        try:
            workbook.save(excel_pfad)
            print(f"\n✅ Erfolgreich **{eintraege_gefunden}** Einträge aktualisiert.")
            print(f"✅ Datei **{os.path.basename(excel_pfad)}** automatisch gespeichert. **Die Formatierung wurde beibehalten.**")
            
        except Exception as e:
            print(f"\nFEHLER beim Speichern der Excel-Datei: {e}")
            print("Stellen Sie sicher, dass die Datei **nicht** geöffnet ist.")
    else:
        print("\nℹ️ Es wurden keine neuen Mieteinträge gefunden oder nur Duplikate ignoriert. Die Excel-Datei wurde nicht geändert.")

if __name__ == "__main__":
    try:
        fuehre_mietabgleich_durch()
    except Exception as e:
        print(f"\nEin kritischer Fehler ist aufgetreten: {e}")
import os
import re
import glob
from datetime import datetime
from flask import Flask, request, jsonify
from dotenv import load_dotenv
import pandas as pd
import resend

# Ladda miljövariabler från .env
load_dotenv()

app = Flask(__name__)

# Konfigurera Resend
resend.api_key = os.getenv("RESEND_API_KEY")
CHEF_EMAIL = os.getenv("CHEF_EMAIL", "marcus.hager@edsvikensel.se")
# På Render: använd data-mappen i projektet, lokalt: OneDrive
if os.path.exists("/opt/render"):
    EXCEL_PATH = os.getenv("EXCEL_PATH", "/opt/render/project/src/data")
else:
    EXCEL_PATH = os.getenv("EXCEL_PATH", r"C:\Users\MarcusHäger\OneDrive - SELATEK\Statistik")


def get_latest_excel_file():
    """Hitta senaste Excel-filen i Statistik-mappen."""
    pattern = os.path.join(EXCEL_PATH, "*.xlsx")
    files = glob.glob(pattern)
    if not files:
        return None
    # Returnera senast modifierade filen
    return max(files, key=os.path.getmtime)


def parse_excel_for_person(file_path, person_name):
    """
    Läs Excel-filen och hitta alla inköp för en person.
    Returnerar dict med personinfo och lista med inköp.

    Excel-struktur (kolumnindex):
    - Kolumn 2: Datum (datetime) eller personrubrik
    - Kolumn 3: Ordernr
    - Kolumn 5: Belopp
    - Kolumn 7: Artikelben1
    - Kolumn 8: Artikelben2
    - Kolumn 11: Saldo (i rubrikraden)
    - Kolumn 12: Kundref
    """
    df = pd.read_excel(file_path, header=None)

    person_name_upper = person_name.upper().strip()
    # Ta bort vanliga prefix/suffix för bättre matchning
    search_names = [person_name_upper]
    # Lägg till varianter (t.ex. "TOMAS" om man söker "Tomas Fredriksson")
    name_parts = person_name_upper.split()
    if len(name_parts) >= 2:
        search_names.append(name_parts[0])  # Förnamn
        search_names.append(name_parts[-1])  # Efternamn

    result = {
        "namn": person_name,
        "saldo": None,
        "kontobelopp": None,
        "inkop": []
    }

    found_person_section = False
    in_person_section = False

    for idx, row in df.iterrows():
        # Konvertera raden till strängar för enkel sökning
        row_values = []
        for v in row:
            if pd.isna(v):
                row_values.append("")
            elif isinstance(v, pd.Timestamp):
                row_values.append(v.strftime("%Y-%m-%d %H:%M:%S"))
            else:
                row_values.append(str(v))

        row_text = " ".join(row_values).upper()

        # Kolla om detta är en personrubrik (t.ex. "231   TOMAS FREDRIKSSON/300")
        col2_val = row_values[2] if len(row_values) > 2 else ""
        is_person_header = re.match(r'^\d+\s+[A-ZÅÄÖ]', col2_val)

        if is_person_header:
            # Detta är en personrubrik
            if any(name in col2_val.upper() for name in search_names):
                found_person_section = True
                in_person_section = True
            elif found_person_section:
                # Vi har hittat en ny person, sluta
                break
            else:
                in_person_section = False
            continue

        if not in_person_section:
            continue

        # Leta efter kontobelopp och saldo i rubrikraden (row_text är UPPERCASE)
        if "KONTOBELOPP:" in row_text:
            for val in row_values:
                val_upper = val.upper()
                if "KONTOBELOPP:" in val_upper:
                    match = re.search(r'Kontobelopp:\s*([\d,\.]+)', val, re.IGNORECASE)
                    if match:
                        result["kontobelopp"] = match.group(1)
                if "SALDO:" in val_upper:
                    match = re.search(r'Saldo:\s*([\d,\.\-]+)', val, re.IGNORECASE)
                    if match:
                        result["saldo"] = match.group(1)
            continue

        # Hoppa över headerraden
        if "DATUM" in row_text and "ARTIKELNR" in row_text:
            continue

        # Kolla om detta är en inköpsrad (har datum i kolumn 2)
        # Använd original row-data, inte row_values (som är strängar)
        col2_raw = row[2] if len(row) > 2 else None
        if isinstance(col2_raw, (datetime, pd.Timestamp)):
            # Detta är en inköpsrad
            inkop = {
                "datum": col2_raw.strftime("%Y-%m-%d"),
                "belopp": str(row[5]) if len(row) > 5 and pd.notna(row[5]) else None,
                "artikel": str(row[7]) if len(row) > 7 and pd.notna(row[7]) else None,
                "beskrivning": str(row[8]) if len(row) > 8 and pd.notna(row[8]) else None
            }
            result["inkop"].append(inkop)

    return result


def create_email_html(person_data):
    """Skapa HTML-formaterat mejlinnehåll."""
    html = f"""
    <h2>Inköpshistorik för {person_data['namn']}</h2>

    <p><strong>Kontobelopp:</strong> {person_data['kontobelopp'] or 'Ej angivet'} kr</p>
    <p><strong>Saldo:</strong> {person_data['saldo'] or 'Ej angivet'} kr</p>

    <h3>Inköp:</h3>
    """

    if person_data['inkop']:
        html += "<table border='1' cellpadding='8' cellspacing='0'>"
        html += "<tr><th>Datum</th><th>Artikel</th><th>Beskrivning</th><th>Belopp</th></tr>"

        for inkop in person_data['inkop']:
            html += f"""
            <tr>
                <td>{inkop['datum']}</td>
                <td>{inkop['artikel'] or '-'}</td>
                <td>{inkop['beskrivning'] or '-'}</td>
                <td>{inkop['belopp'] or '-'} kr</td>
            </tr>
            """
        html += "</table>"
    else:
        html += "<p>Inga inköp hittades.</p>"

    html += "<br><p><em>Automatiskt genererat av Kläder-systemet</em></p>"

    return html


def send_email(to_email, person_name, html_content):
    """Skicka mejl via Resend."""
    params = {
        "from": "Klädsystem <onboarding@resend.dev>",
        "to": [to_email],
        "subject": f"Inköpshistorik för {person_name}",
        "html": html_content
    }

    email = resend.Emails.send(params)
    return email


@app.route("/", methods=["GET"])
def home():
    """Enkel startsida för att verifiera att appen körs."""
    return jsonify({"status": "ok", "message": "Kläder-API är igång!"})


@app.route("/webhook", methods=["POST"])
def webhook():
    """
    Webhook-endpoint som tar emot förfrågan från Power Automate.
    Förväntar JSON med 'namn' eller 'email_body' (hela mejlinnehållet).
    Om 'email_body' skickas extraheras namnet automatiskt från "Namn: XXX".
    """
    try:
        data = request.get_json(force=True, silent=True)
    except Exception as e:
        return jsonify({"error": f"Kunde inte parsa JSON: {str(e)}"}), 400

    if not data:
        # Försök läsa raw body
        raw_body = request.get_data(as_text=True)
        return jsonify({"error": "Ingen data mottagen", "raw_body_preview": raw_body[:200] if raw_body else "tom"}), 400

    # Hantera både direkt 'namn' och 'email_body'
    person_name = None

    vill_kopa = ""

    if "namn" in data and data["namn"]:
        person_name = data["namn"].strip()
    elif "email_body" in data:
        # Extrahera namn och vill_kopa från mejlkroppen
        email_body = str(data["email_body"])

        # Ta bort HTML-taggar för enklare parsing
        clean_body = re.sub(r'<[^>]+>', ' ', email_body)

        # Försök hitta "Namn: XXX" - ta bara bokstäver, mellanslag och svenska tecken
        # Stoppa vid "Vill" för att inte fånga "Vill köpa" som en del av namnet
        match = re.search(r'Namn:\s*([A-Za-zÅÄÖåäö\s\-]+?)(?:\s*Vill|\s*$|\n|\r)', clean_body, re.IGNORECASE)
        if match:
            person_name = match.group(1).strip()
            # Ta bort eventuella extra mellanslag
            person_name = ' '.join(person_name.split())

        # Försök hitta "Vill köpa: XXX" från mejlkroppen
        vill_kopa_match = re.search(r'Vill\s*köpa:\s*([^\n\r]+)', clean_body, re.IGNORECASE)
        if vill_kopa_match:
            vill_kopa = vill_kopa_match.group(1).strip()
            # Ta bort eventuella extra mellanslag
            vill_kopa = ' '.join(vill_kopa.split())

    if not person_name:
        return jsonify({
            "error": "Kunde inte hitta namn. Skicka 'namn' eller 'email_body' med 'Namn: XXX'",
            "received_keys": list(data.keys()) if data else []
        }), 400

    # Tillåt även vill_kopa direkt från JSON
    if not vill_kopa and data.get("vill_kopa"):
        vill_kopa = data.get("vill_kopa", "")

    # Hitta senaste Excel-filen
    excel_file = get_latest_excel_file()
    if not excel_file:
        return jsonify({"error": "Ingen Excel-fil hittades"}), 404

    # Hämta personens inköp
    person_data = parse_excel_for_person(excel_file, person_name)

    if not person_data["inkop"] and not person_data["saldo"]:
        return jsonify({"error": f"Hittade ingen data för {person_name}"}), 404

    # Skapa mejlinnehåll
    html_content = create_email_html(person_data)

    if vill_kopa:
        html_content += f"<p><strong>Vill köpa:</strong> {vill_kopa}</p>"

    # Skicka mejl
    try:
        email_result = send_email(CHEF_EMAIL, person_name, html_content)
        return jsonify({
            "status": "success",
            "message": f"Mejl skickat till {CHEF_EMAIL}",
            "person": person_name,
            "antal_inkop": len(person_data["inkop"]),
            "email_id": email_result.get("id")
        })
    except Exception as e:
        return jsonify({"error": f"Kunde inte skicka mejl: {str(e)}"}), 500


@app.route("/test/<namn>", methods=["GET"])
def test_person(namn):
    """Test-endpoint för att söka efter en person utan att skicka mejl."""
    excel_file = get_latest_excel_file()
    if not excel_file:
        return jsonify({"error": "Ingen Excel-fil hittades"}), 404

    person_data = parse_excel_for_person(excel_file, namn)
    return jsonify(person_data)


if __name__ == "__main__":
    app.run(debug=True, port=5001)

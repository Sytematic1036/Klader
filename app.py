import os
import re
import glob
from datetime import datetime, timedelta
from functools import wraps
from flask import Flask, request, jsonify, render_template, redirect, url_for, session
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from dotenv import load_dotenv
import pandas as pd
import resend

# Ladda miljövariabler från .env
load_dotenv()

app = Flask(__name__)

# Konfiguration
app.secret_key = os.getenv("SECRET_KEY", "dev-secret-key-change-in-production")
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=3)

# Databaskonfiguration
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///klader.db")
# Fix för Render PostgreSQL URL - använd psycopg3-drivrutinen
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+psycopg://", 1)
elif DATABASE_URL.startswith("postgresql://") and "+psycopg" not in DATABASE_URL:
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+psycopg://", 1)
app.config['SQLALCHEMY_DATABASE_URI'] = DATABASE_URL
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# Konfigurera Resend
resend.api_key = os.getenv("RESEND_API_KEY")
CHEF_EMAIL = os.getenv("CHEF_EMAIL", "marcus.hager@edsvikensel.se")
INVITE_CODE = os.getenv("INVITE_CODE", "klader2024")
ALLOWED_EMAILS = [e.strip().lower() for e in os.getenv("ALLOWED_EMAILS", "").split(",") if e.strip()]

# På Render: använd data-mappen i projektet, lokalt: OneDrive
if os.path.exists("/opt/render"):
    EXCEL_PATH = os.getenv("EXCEL_PATH", "/opt/render/project/src/data")
else:
    EXCEL_PATH = os.getenv("EXCEL_PATH", r"C:\Users\MarcusHäger\OneDrive - SELATEK\Statistik")


# ==================== DATABAS-MODELLER ====================

class User(db.Model):
    __tablename__ = 'users'
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(256), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)


class LoginEvent(db.Model):
    __tablename__ = 'login_events'
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    login_time = db.Column(db.DateTime, default=datetime.utcnow)
    ip_address = db.Column(db.String(50))


# Skapa tabeller
with app.app_context():
    db.create_all()


# ==================== AUTH DECORATORS ====================

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session:
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function


@app.before_request
def check_session_timeout():
    """Kontrollera session timeout före varje request."""
    # Skippa för statiska filer och login/register
    if request.endpoint in ['login', 'register', 'webhook', 'test_person', 'static', 'health']:
        return

    if 'user' in session:
        last_activity = session.get('last_activity')
        if last_activity:
            last_activity = datetime.fromisoformat(last_activity)
            if datetime.utcnow() - last_activity > timedelta(hours=3):
                session.clear()
                return redirect(url_for('login'))
        session['last_activity'] = datetime.utcnow().isoformat()


# ==================== AUTH ROUTES ====================

@app.route("/", methods=["GET"])
def index():
    """Omdirigera till login eller dashboard."""
    if 'user' in session:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))


@app.route("/login", methods=["GET", "POST"])
def login():
    """Login-sida."""
    error = None
    success = request.args.get('registered')

    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")

        user = User.query.filter_by(email=email).first()

        if user and user.check_password(password):
            session['user'] = email
            session['last_activity'] = datetime.utcnow().isoformat()
            session.permanent = True

            # Logga login-event
            login_event = LoginEvent(
                user_id=user.id,
                ip_address=request.remote_addr
            )
            db.session.add(login_event)
            db.session.commit()

            return redirect(url_for('dashboard'))
        else:
            error = "Fel e-post eller lösenord"

    return render_template("login.html", error=error, success="Registrering lyckades! Logga in." if success else None)


@app.route("/register", methods=["GET", "POST"])
def register():
    """Registrerings-sida."""
    error = None

    if request.method == "POST":
        invite_code = request.form.get("invite_code", "").strip()
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        password_confirm = request.form.get("password_confirm", "")

        # Validera inbjudningskod
        if invite_code != INVITE_CODE:
            error = "Felaktig inbjudningskod"
        # Validera e-post (om ALLOWED_EMAILS är satt)
        elif ALLOWED_EMAILS and email not in ALLOWED_EMAILS:
            error = "Denna e-postadress är inte tillåten"
        # Validera lösenord
        elif password != password_confirm:
            error = "Lösenorden matchar inte"
        elif len(password) < 6:
            error = "Lösenordet måste vara minst 6 tecken"
        # Kontrollera om användaren redan finns
        elif User.query.filter_by(email=email).first():
            error = "E-postadressen är redan registrerad"
        else:
            # Skapa användare
            user = User(email=email)
            user.set_password(password)
            db.session.add(user)
            db.session.commit()

            return redirect(url_for('login', registered=1))

    return render_template("register.html", error=error)


@app.route("/logout")
def logout():
    """Logga ut."""
    session.clear()
    return redirect(url_for('login'))


@app.route("/dashboard")
@login_required
def dashboard():
    """Dashboard för inloggade användare."""
    return render_template("dashboard.html", user=session.get('user'), chef_email=CHEF_EMAIL)


# ==================== API ROUTES ====================

@app.route("/health", methods=["GET"])
def health():
    """Health check endpoint."""
    return jsonify({"status": "ok", "message": "Kläder-API är igång!"})


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
    """
    df = pd.read_excel(file_path, header=None)

    person_name_upper = person_name.upper().strip()
    search_names = [person_name_upper]
    name_parts = person_name_upper.split()
    if len(name_parts) >= 2:
        search_names.append(name_parts[0])
        search_names.append(name_parts[-1])

    result = {
        "namn": person_name,
        "saldo": None,
        "kontobelopp": None,
        "inkop": []
    }

    found_person_section = False
    in_person_section = False

    for idx, row in df.iterrows():
        row_values = []
        for v in row:
            if pd.isna(v):
                row_values.append("")
            elif isinstance(v, pd.Timestamp):
                row_values.append(v.strftime("%Y-%m-%d %H:%M:%S"))
            else:
                row_values.append(str(v))

        row_text = " ".join(row_values).upper()

        col2_val = row_values[2] if len(row_values) > 2 else ""
        is_person_header = re.match(r'^\d+\s+[A-ZÅÄÖ]', col2_val)

        if is_person_header:
            if any(name in col2_val.upper() for name in search_names):
                found_person_section = True
                in_person_section = True
            elif found_person_section:
                break
            else:
                in_person_section = False
            continue

        if not in_person_section:
            continue

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

        if "DATUM" in row_text and "ARTIKELNR" in row_text:
            continue

        col2_raw = row[2] if len(row) > 2 else None
        if isinstance(col2_raw, (datetime, pd.Timestamp)):
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


@app.route("/webhook", methods=["POST"])
def webhook():
    """
    Webhook-endpoint som tar emot förfrågan från Power Automate.
    Förväntar JSON med 'namn' eller 'email_body' (hela mejlinnehållet).
    """
    try:
        data = request.get_json(force=True, silent=True)
    except Exception as e:
        return jsonify({"error": f"Kunde inte parsa JSON: {str(e)}"}), 400

    if not data:
        raw_body = request.get_data(as_text=True)
        return jsonify({"error": "Ingen data mottagen", "raw_body_preview": raw_body[:200] if raw_body else "tom"}), 400

    person_name = None
    vill_kopa = ""

    if "namn" in data and data["namn"]:
        person_name = data["namn"].strip()
    elif "email_body" in data:
        email_body = str(data["email_body"])
        clean_body = re.sub(r'<[^>]+>', ' ', email_body)

        match = re.search(r'Namn:\s*([A-Za-zÅÄÖåäö\s\-]+?)(?:\s*Vill|\s*$|\n|\r)', clean_body, re.IGNORECASE)
        if match:
            person_name = match.group(1).strip()
            person_name = ' '.join(person_name.split())

        vill_kopa_match = re.search(r'Vill\s*köpa:\s*([^\n\r]+)', clean_body, re.IGNORECASE)
        if vill_kopa_match:
            vill_kopa = vill_kopa_match.group(1).strip()
            vill_kopa = ' '.join(vill_kopa.split())

    if not person_name:
        return jsonify({
            "error": "Kunde inte hitta namn. Skicka 'namn' eller 'email_body' med 'Namn: XXX'",
            "received_keys": list(data.keys()) if data else []
        }), 400

    if not vill_kopa and data.get("vill_kopa"):
        vill_kopa = data.get("vill_kopa", "")

    excel_file = get_latest_excel_file()
    if not excel_file:
        return jsonify({"error": "Ingen Excel-fil hittades"}), 404

    person_data = parse_excel_for_person(excel_file, person_name)

    if not person_data["inkop"] and not person_data["saldo"]:
        return jsonify({"error": f"Hittade ingen data för {person_name}"}), 404

    html_content = create_email_html(person_data)

    if vill_kopa:
        html_content += f"<p><strong>Vill köpa:</strong> {vill_kopa}</p>"

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

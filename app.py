import os
import re
import io
import secrets
from datetime import datetime, timedelta
from functools import wraps
from flask import Flask, request, jsonify, render_template, render_template_string, redirect, url_for, session, flash
from werkzeug.utils import secure_filename
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
INVITE_CODE = os.getenv("INVITE_CODE", "klader2024")

# Chefer och deras e-postadresser
CHEFER = {
    "marcus": "marcus.hager@edsvikensel.se",
    "marcus häger": "marcus.hager@edsvikensel.se",
    "andreas": "andreas.danielsson@edsvikensel.se",
    "andreas danielsson": "andreas.danielsson@edsvikensel.se",
    "pernilla": "pernilla.ostberg@msjobergsel.se",
    "pernilla östberg": "pernilla.ostberg@msjobergsel.se",
}
DEFAULT_CHEF_EMAIL = "marcus.hager@edsvikensel.se"
ALLOWED_EMAILS = [e.strip().lower() for e in os.getenv("ALLOWED_EMAILS", "").split(",") if e.strip()]

# Rekvisitionsnummer-konfiguration
CODE_LENGTH = 8
# Alfabet utan lättförväxlade tecken (0/O, 1/I osv)
CODE_ALPHABET = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"



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


class ExcelFile(db.Model):
    """Lagrar Excel-filen i databasen för persistent lagring."""
    __tablename__ = 'excel_files'
    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    data = db.Column(db.LargeBinary, nullable=False)
    size = db.Column(db.Integer, nullable=False)
    uploaded_at = db.Column(db.DateTime, default=datetime.utcnow)
    uploaded_by = db.Column(db.String(120))


class Requisition(db.Model):
    """Lagrar rekvisitioner med status för godkännande."""
    __tablename__ = 'requisitions'
    id = db.Column(db.Integer, primary_key=True)
    code = db.Column(db.String(8), unique=True, nullable=False, index=True)
    employee_name = db.Column(db.String(255), nullable=False)
    chef_name = db.Column(db.String(255))
    chef_email = db.Column(db.String(255))
    vill_kopa = db.Column(db.Text)
    status = db.Column(db.String(20), default='Väntar')  # Väntar, Godkänt, Ej godkänt
    request_date = db.Column(db.Date, default=datetime.utcnow)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    decided_at = db.Column(db.DateTime)  # När chefen fattade beslut


# Skapa tabeller
with app.app_context():
    db.create_all()


# ==================== REKVISITIONSNUMMER ====================

def generate_requisition_code():
    """Generera ett unikt 8-teckens rekvisitionsnummer."""
    while True:
        code = ''.join(secrets.choice(CODE_ALPHABET) for _ in range(CODE_LENGTH))
        # Kontrollera att koden inte redan finns
        existing = Requisition.query.filter_by(code=code).first()
        if not existing:
            return code


def create_requisition(employee_name, chef_name=None, chef_email=None, vill_kopa=None):
    """Skapa en ny rekvisition och returnera rekvisitionsnumret."""
    code = generate_requisition_code()

    requisition = Requisition(
        code=code,
        employee_name=employee_name,
        chef_name=chef_name,
        chef_email=chef_email,
        vill_kopa=vill_kopa,
        status='Väntar',
        request_date=datetime.utcnow().date()
    )

    db.session.add(requisition)
    db.session.commit()

    return code


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
    return render_template("dashboard.html", user=session.get('user'), chef_email=DEFAULT_CHEF_EMAIL)


# ==================== FILHANTERING ====================

def get_current_file_info():
    """Hämta info om nuvarande Excel-fil från databasen."""
    excel_file = get_excel_file_from_db()
    if not excel_file:
        return None

    return {
        "name": excel_file.filename,
        "size": round(excel_file.size / 1024, 1),
        "modified": excel_file.uploaded_at.strftime("%Y-%m-%d %H:%M"),
        "uploaded_by": excel_file.uploaded_by
    }


@app.route("/files", methods=["GET"])
@login_required
def files_page():
    """Sida för filhantering."""
    current_file = get_current_file_info()
    message = request.args.get("message")
    message_type = request.args.get("type", "success")

    return render_template("upload.html",
                         user=session.get('user'),
                         current_file=current_file,
                         message=message,
                         message_type=message_type)


@app.route("/upload", methods=["POST"])
@login_required
def upload_file():
    """Ladda upp ny Excel-fil och spara i databasen."""
    if 'file' not in request.files:
        return redirect(url_for('files_page', message="Ingen fil vald", type="error"))

    file = request.files['file']

    if file.filename == '':
        return redirect(url_for('files_page', message="Ingen fil vald", type="error"))

    if not file.filename.endswith('.xlsx'):
        return redirect(url_for('files_page', message="Endast .xlsx-filer tillåtna", type="error"))

    # Läs fildata
    filename = secure_filename(file.filename)
    file_data = file.read()
    file_size = len(file_data)

    # Ta bort befintliga filer från databasen
    ExcelFile.query.delete()

    # Spara ny fil i databasen
    excel_file = ExcelFile(
        filename=filename,
        data=file_data,
        size=file_size,
        uploaded_by=session.get('user')
    )
    db.session.add(excel_file)
    db.session.commit()

    return redirect(url_for('files_page', message=f"Filen '{filename}' har laddats upp", type="success"))


@app.route("/delete-file", methods=["POST"])
@login_required
def delete_file():
    """Radera nuvarande Excel-fil från databasen."""
    excel_file = get_excel_file_from_db()

    if not excel_file:
        return redirect(url_for('files_page', message="Ingen fil att radera", type="error"))

    try:
        filename = excel_file.filename
        db.session.delete(excel_file)
        db.session.commit()
        return redirect(url_for('files_page', message=f"Filen '{filename}' har raderats", type="success"))
    except Exception as e:
        return redirect(url_for('files_page', message=f"Kunde inte radera filen: {str(e)}", type="error"))


# ==================== API ROUTES ====================

@app.route("/health", methods=["GET"])
def health():
    """Health check endpoint."""
    return jsonify({"status": "ok", "message": "Kläder-API är igång!"})


def get_excel_file_from_db():
    """Hämta Excel-filen från databasen."""
    return ExcelFile.query.order_by(ExcelFile.uploaded_at.desc()).first()


def get_excel_data_as_file():
    """Hämta Excel-data som en fil-lik objekt för pandas."""
    excel_file = get_excel_file_from_db()
    if not excel_file:
        return None
    return io.BytesIO(excel_file.data)


def parse_excel_for_person(excel_data, person_name):
    """
    Läs Excel-filen och hitta alla inköp för en person.
    excel_data kan vara en filsökväg eller ett BytesIO-objekt.

    Nytt format:
    - Kolumn E (index 4) = Kundref (personens namn)
    - Kolumn F (index 5) = ArtNr
    - Kolumn G (index 6) = Artikelben1 (beskrivning)
    - Kolumn H (index 7) = Artikelben2 (detaljer)
    - Kolumn I (index 8) = Kvantitet
    - Kolumn J (index 9) = Belopp
    - Kolumn L (index 11) = Fakturadat.

    Returnerar dict med personinfo och lista med inköp.
    """
    df = pd.read_excel(excel_data, header=0)  # Första raden är rubrik

    person_name_upper = person_name.upper().strip()
    # Kräv matchning på hela namnet (både för- och efternamn)
    search_name = person_name_upper

    result = {
        "namn": person_name,
        "total_belopp": 0,
        "inkop": []
    }

    total_belopp = 0

    for idx, row in df.iterrows():
        # Hämta Kundref (kolumn E)
        kundref = str(row.iloc[4]) if pd.notna(row.iloc[4]) else ""

        # Rensa Kundref från nummer och suffix
        # Ta bort prefix: nummer eller ZZ+nummer i början (t.ex. "267 Johan" eller "ZZ123456 Johan")
        kundref_clean = re.sub(r'^[A-Z]{0,2}\d+\s*', '', kundref).strip()
        # Ta bort suffix: /nnn, korta nummer (3+ siffror), eller ZZ+telefonnummer
        kundref_clean = re.sub(r'[/\s]+\d{3,}.*$', '', kundref_clean).strip()
        kundref_clean = re.sub(r'\s+[A-Z]{0,2}\d{6,}.*$', '', kundref_clean).strip()
        kundref_upper = kundref_clean.upper()

        # Kolla om denna rad matchar personen vi söker (hela namnet måste finnas)
        if search_name not in kundref_upper:
            continue

        # Hämta data från raden
        artikelnr = str(row.iloc[5]) if pd.notna(row.iloc[5]) else None
        artikelben1 = str(row.iloc[6]) if pd.notna(row.iloc[6]) else ""
        artikelben2 = str(row.iloc[7]) if pd.notna(row.iloc[7]) else ""
        kvantitet = int(row.iloc[8]) if pd.notna(row.iloc[8]) else 1
        belopp = float(row.iloc[9]) if pd.notna(row.iloc[9]) else 0
        fakturadatum = row.iloc[11] if pd.notna(row.iloc[11]) else None

        # Formatera datum
        if isinstance(fakturadatum, (datetime, pd.Timestamp)):
            datum_str = fakturadatum.strftime("%Y-%m-%d")
        else:
            datum_str = str(fakturadatum) if fakturadatum else None

        # Kombinera artikelbeskrivning
        beskrivning = f"{artikelben1} {artikelben2}".strip()

        inkop = {
            "datum": datum_str,
            "artikelnr": artikelnr,
            "beskrivning": beskrivning,
            "kvantitet": kvantitet,
            "belopp": round(belopp, 2)
        }
        result["inkop"].append(inkop)
        total_belopp += belopp

    result["total_belopp"] = round(total_belopp, 2)

    return result


def create_email_html(person_data):
    """Skapa HTML-formaterat mejlinnehåll."""
    html = f"""
    <h2>Inköpshistorik för {person_data['namn']}</h2>

    <p><strong>Totalt belopp:</strong> {person_data['total_belopp']} kr</p>
    <p><strong>Antal inköp:</strong> {len(person_data['inkop'])}</p>

    <h3>Inköp:</h3>
    """

    if person_data['inkop']:
        html += "<table border='1' cellpadding='8' cellspacing='0'>"
        html += "<tr><th>Datum</th><th>Artikelnr</th><th>Beskrivning</th><th>Antal</th><th>Belopp</th></tr>"

        for inkop in person_data['inkop']:
            html += f"""
            <tr>
                <td>{inkop['datum'] or '-'}</td>
                <td>{inkop['artikelnr'] or '-'}</td>
                <td>{inkop['beskrivning'] or '-'}</td>
                <td>{inkop['kvantitet']}</td>
                <td>{inkop['belopp']} kr</td>
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
    chef_email = DEFAULT_CHEF_EMAIL
    chef_name_parsed = ""

    if "namn" in data and data["namn"]:
        person_name = data["namn"].strip()
    elif "email_body" in data:
        email_body = str(data["email_body"])
        clean_body = re.sub(r'<[^>]+>', ' ', email_body)

        # Parsa Namn - matchar allt efter "Namn:" tills nästa fält eller radslut
        match = re.search(r'Namn:\s*([A-Za-zÅÄÖåäö\s\-]+)', clean_body, re.IGNORECASE)
        if match:
            person_name = match.group(1).strip()
            # Ta bort eventuella följande nyckelord
            person_name = re.split(r'\s+(?:Chef|Vill)', person_name, flags=re.IGNORECASE)[0].strip()
            person_name = ' '.join(person_name.split())

        # Parsa Vill köpa
        vill_kopa_match = re.search(r'Vill\s*köpa:\s*([A-Za-zÅÄÖåäö0-9\s\-]+)', clean_body, re.IGNORECASE)
        if vill_kopa_match:
            vill_kopa = vill_kopa_match.group(1).strip()
            vill_kopa = re.split(r'\s+(?:Chef|Namn)', vill_kopa, flags=re.IGNORECASE)[0].strip()
            vill_kopa = ' '.join(vill_kopa.split())

        # Parsa Chef
        chef_match = re.search(r'Chef:\s*([A-Za-zÅÄÖåäö\s\-]+)', clean_body, re.IGNORECASE)
        if chef_match:
            chef_name_parsed = chef_match.group(1).strip()
            # Ta bort eventuella följande nyckelord (Namn, Vill)
            chef_name_parsed = re.split(r'\s+(?:Namn|Vill)', chef_name_parsed, flags=re.IGNORECASE)[0].strip()
            chef_name_parsed = ' '.join(chef_name_parsed.split())
            chef_name_lower = chef_name_parsed.lower()
            # Hitta matchande chef
            if chef_name_lower in CHEFER:
                chef_email = CHEFER[chef_name_lower]

    if not person_name:
        return jsonify({
            "error": "Kunde inte hitta namn. Skicka 'namn' eller 'email_body' med 'Namn: XXX'",
            "received_keys": list(data.keys()) if data else []
        }), 400

    if not vill_kopa and data.get("vill_kopa"):
        vill_kopa = data.get("vill_kopa", "")

    excel_data = get_excel_data_as_file()
    if not excel_data:
        return jsonify({"error": "Ingen Excel-fil hittades"}), 404

    person_data = parse_excel_for_person(excel_data, person_name)

    if not person_data["inkop"]:
        return jsonify({"error": f"Hittade ingen data för {person_name}"}), 404

    # Skapa rekvisition och få rekvisitionsnummer
    requisition_code = create_requisition(
        employee_name=person_name,
        chef_name=chef_name_parsed,
        chef_email=chef_email,
        vill_kopa=vill_kopa
    )

    # Skapa URL:er för godkännande/avslag (knappar i mejlet)
    base_url = os.getenv("BASE_URL", "https://klader.onrender.com")
    approve_url = f"{base_url}/requisition/{requisition_code}/approve"
    reject_url = f"{base_url}/requisition/{requisition_code}/reject"

    html_content = create_email_html(person_data)

    # Text som chefen kan vidarebefordra till den anställde
    forward_text_html = f"""
    <div style="background: #e8f5e9; padding: 20px; border-radius: 10px; margin-bottom: 20px; border: 2px solid #28a745;">
        <p style="margin: 0; color: #333; font-size: 15px; line-height: 1.6;">
            Hej!<br><br>
            Ditt inköp är godkänt. Visa detta rekvisitionsnummer för personalen på Ahlsell och be dem lägga in det som referens på ordern.<br><br>
            <strong style="font-family: 'Courier New', Courier, monospace; font-size: 20px; letter-spacing: 2px; color: #667eea;">Rekvisitionsnummer: {requisition_code}</strong><br>
            <strong>Godkänt inköp: {vill_kopa if vill_kopa else 'Ej angivet'}</strong><br><br>
            Mvh<br>
            {chef_name_parsed if chef_name_parsed else 'Din chef'}
        </p>
        <p style="margin-top: 15px; margin-bottom: 0; font-size: 12px; color: #666; font-style: italic;">
            (Tryck på någon av knapparna "Godkänn" eller "Avslå" nedan. Vidarebefordra detta mejl till den anställde. Ändra "godkänt" till "ej godkänt" om du nekar.)
        </p>
    </div>
    """

    # Lägg till rekvisitionsnummer och godkännandeknappar
    approval_buttons_html = f"""
    <div style="background: #f8f9fa; padding: 20px; border-radius: 10px; margin-bottom: 20px; border: 2px solid #667eea;">
        <p style="margin-top: 0; color: #333; font-size: 16px;">Rekvisitionsnummer:</p>
        <p style="color: #667eea; font-family: 'Courier New', Courier, monospace; font-size: 28px; font-weight: bold; letter-spacing: 3px; margin: 10px 0 15px 0;">{requisition_code}</p>
        <p><strong>Vill köpa:</strong> {vill_kopa if vill_kopa else 'Ej angivet'}</p>
        <p style="margin-bottom: 15px;">Klicka på en knapp nedan för att registrera ditt beslut (öppnas i ny flik):</p>
        <div style="display: inline-block;">
            <a href="{approve_url}" target="_blank" style="display: inline-block; background: #28a745; color: white; padding: 12px 30px; text-decoration: none; border-radius: 5px; font-weight: bold; margin-right: 10px;">Godkänn</a>
            <a href="{reject_url}" target="_blank" style="display: inline-block; background: #dc3545; color: white; padding: 12px 30px; text-decoration: none; border-radius: 5px; font-weight: bold;">Avslå</a>
        </div>
    </div>
    """

    html_content = forward_text_html + approval_buttons_html + html_content

    # Returnera data till Power Automate (som skickar mejlet)
    return jsonify({
        "status": "success",
        "person": person_name,
        "total_belopp": person_data["total_belopp"],
        "antal_inkop": len(person_data["inkop"]),
        "vill_kopa": vill_kopa,
        "chef_email": chef_email,
        "requisition_code": requisition_code,
        "email_subject": f"Inköpshistorik för {person_name} - Rekv: {requisition_code}",
        "email_body_html": html_content
    })


@app.route("/test/<namn>", methods=["GET"])
def test_person(namn):
    """Test-endpoint för att söka efter en person utan att skicka mejl."""
    excel_data = get_excel_data_as_file()
    if not excel_data:
        return jsonify({"error": "Ingen Excel-fil hittades"}), 404

    person_data = parse_excel_for_person(excel_data, namn)
    return jsonify(person_data)


# ==================== REKVISITIONS-ENDPOINTS ====================

@app.route("/requisition/<code>/approve", methods=["GET"])
def approve_requisition(code):
    """Godkänn en rekvisition via länk i mejlet."""
    requisition = Requisition.query.filter_by(code=code.upper()).first()

    if not requisition:
        return render_template_string(DECISION_RESULT_HTML,
            title="Rekvisition hittades inte",
            message=f"Rekvisitionsnummer {code} finns inte i systemet.",
            status_class="error")

    # Kolla om redan hanterad (status innehåller 'Godkänt' eller 'godkänt')
    current_status = (requisition.status or '').lower()
    if 'godkänt' in current_status:
        return render_template_string(DECISION_RESULT_HTML,
            title="Redan hanterad",
            message=f"Denna rekvisition har redan status: {requisition.status}",
            status_class="warning")

    requisition.status = 'Godkänt'
    requisition.decided_at = datetime.utcnow()
    db.session.commit()

    return render_template_string(DECISION_RESULT_HTML,
        title="Beslutet är nu registrerat",
        message="",
        status_class="success",
        requisition=requisition)


@app.route("/requisition/<code>/reject", methods=["GET"])
def reject_requisition(code):
    """Avslå en rekvisition via länk i mejlet."""
    requisition = Requisition.query.filter_by(code=code.upper()).first()

    if not requisition:
        return render_template_string(DECISION_RESULT_HTML,
            title="Rekvisition hittades inte",
            message=f"Rekvisitionsnummer {code} finns inte i systemet.",
            status_class="error")

    # Kolla om redan hanterad (status innehåller 'Godkänt' eller 'godkänt')
    current_status = (requisition.status or '').lower()
    if 'godkänt' in current_status:
        return render_template_string(DECISION_RESULT_HTML,
            title="Redan hanterad",
            message=f"Denna rekvisition har redan status: {requisition.status}",
            status_class="warning")

    requisition.status = 'Ej godkänt'
    requisition.decided_at = datetime.utcnow()
    db.session.commit()

    return render_template_string(DECISION_RESULT_HTML,
        title="Beslutet är nu registrerat",
        message="",
        status_class="rejected",
        requisition=requisition)


# HTML-mall för beslutssida (visas när chef klickar på knapp)
DECISION_RESULT_HTML = """
<!DOCTYPE html>
<html lang="sv">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ title }} - Klädsystem</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: #f5f5f5;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }
        .container {
            background: white;
            border-radius: 15px;
            padding: 40px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.1);
            text-align: center;
            max-width: 600px;
            width: 100%;
        }
        .icon { font-size: 60px; margin-bottom: 20px; }
        .success .icon { color: #28a745; }
        .rejected .icon { color: #dc3545; }
        .warning .icon { color: #ffc107; }
        .error .icon { color: #6c757d; }
        h1 { color: #333; margin-bottom: 15px; }
        p { color: #666; line-height: 1.6; }
        .details { background: #f8f9fa; padding: 15px; border-radius: 8px; margin-top: 20px; text-align: left; }
        .details strong { color: #333; }
        .copy-section {
            margin-top: 25px;
            text-align: left;
        }
        .copy-section h3 {
            color: #333;
            margin-bottom: 10px;
            font-size: 16px;
        }
        .copy-box {
            background: #f8f9fa;
            border: 2px solid #dee2e6;
            border-radius: 8px;
            padding: 15px;
            font-size: 14px;
            line-height: 1.6;
            color: #333;
            position: relative;
        }
        .copy-btn {
            display: block;
            width: 100%;
            margin-top: 10px;
            padding: 12px 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            border-radius: 5px;
            font-size: 14px;
            font-weight: bold;
            cursor: pointer;
            transition: transform 0.2s;
        }
        .copy-btn:hover {
            transform: translateY(-2px);
        }
        .copy-btn.copied {
            background: #28a745;
        }
        .close-hint {
            margin-top: 25px;
            padding: 12px 20px;
            background: #e9ecef;
            border-radius: 8px;
            color: #495057;
            font-size: 14px;
        }
        .code-display {
            font-family: 'Courier New', Courier, monospace;
            font-size: 18px;
            font-weight: bold;
            letter-spacing: 2px;
            color: #667eea;
        }
    </style>
</head>
<body>
    <div class="container {{ status_class }}">
        <div class="icon">
            {% if status_class == 'success' %}✓{% elif status_class == 'rejected' %}✗{% elif status_class == 'warning' %}⚠{% else %}?{% endif %}
        </div>
        <h1>{{ title }}</h1>
        <p>{{ message }}</p>
        {% if requisition %}
        <div class="details">
            <p><strong>Rekvisitionsnummer:</strong> <span class="code-display">{{ requisition.code }}</span></p>
            <p><strong>Anställd:</strong> {{ requisition.employee_name }}</p>
            <p><strong>Vill köpa:</strong> {{ requisition.vill_kopa or 'Ej angivet' }}</p>
            <p><strong>Datum:</strong> {{ requisition.request_date }}</p>
        </div>

        <div class="copy-section">
            <h3>Kopiera och vidarebefordra till {{ requisition.employee_name }}:</h3>
            <div class="copy-box" id="copyText">Hej!

Ditt inköp är godkänt. Visa detta rekvisitionsnummer för personalen på Ahlsell och be dem lägga in det som referens på ordern.

Rekvisitionsnummer: {{ requisition.code }}

Mvh
{{ requisition.chef_name or 'Din chef' }}</div>
            <button class="copy-btn" onclick="copyToClipboard()">Kopiera text</button>
            <p style="font-size: 12px; color: #666; margin-top: 8px; text-align: center;">Ändra till "ej godkänt" om du nekar inköpet.</p>
        </div>
        {% endif %}
        <div class="close-hint">
            Du kan nu stänga denna flik och gå tillbaka till mejlet.
        </div>
    </div>

    <script>
        function copyToClipboard() {
            var text = document.getElementById('copyText').innerText;
            navigator.clipboard.writeText(text).then(function() {
                var btn = document.querySelector('.copy-btn');
                btn.textContent = 'Kopierat!';
                btn.classList.add('copied');
                setTimeout(function() {
                    btn.textContent = 'Kopiera text';
                    btn.classList.remove('copied');
                }, 2000);
            });
        }
    </script>
</body>
</html>
"""


@app.route("/requisitions")
@login_required
def requisitions_page():
    """Sida för att visa alla rekvisitioner."""
    requisitions = Requisition.query.order_by(Requisition.created_at.desc()).all()
    return render_template("requisitions.html",
                         user=session.get('user'),
                         requisitions=requisitions)


@app.route("/requisition/<int:req_id>/delete", methods=["POST"])
@login_required
def delete_requisition(req_id):
    """Radera en rekvisition."""
    requisition = Requisition.query.get(req_id)
    if requisition:
        db.session.delete(requisition)
        db.session.commit()
        return redirect(url_for('requisitions_page', message="Rekvisition raderad", type="success"))
    return redirect(url_for('requisitions_page', message="Rekvisition hittades inte", type="error"))


@app.route("/requisitions/delete-all", methods=["POST"])
@login_required
def delete_all_requisitions():
    """Radera alla rekvisitioner (för testning)."""
    Requisition.query.delete()
    db.session.commit()
    return redirect(url_for('requisitions_page', message="Alla rekvisitioner raderade", type="success"))


if __name__ == "__main__":
    app.run(debug=True, port=5001)

import os
import re
import smtplib
import ssl
import time
import uuid
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formatdate
from pathlib import Path

from flask import Flask, request, jsonify, render_template, send_file
from werkzeug.utils import secure_filename
from fpdf import FPDF
import json
import unicodedata
from io import BytesIO

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10MB
UPLOAD_FOLDER = Path("uploads")
UPLOAD_FOLDER.mkdir(exist_ok=True)
GENERATED_FOLDER = Path("generated")
GENERATED_FOLDER.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {'.pdf', '.doc', '.docx'}

# --- Helpers ---

def allowed_file(filename):
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS

GENERIC_DOMAINS = {"gmail","yahoo","hotmail","outlook","live","icloud","free","laposte","orange","sfr","aol","wanadoo","gmx","proton","protonmail"}

# --- Excel Suivi ---
EXCEL_HEADERS = [
    "Date envoi de la candidature ou du contact",
    "Entreprise",
    "Candidature spontanée ou réponse à une offre",
    "Stage ciblé",
    "Nom de l'interlocuteur",
    "Coordonnées de l'interlocuteur (tel, @)",
    "Date d'un éventuel entretien",
    "Résultats de la démarche (en attente de réponse, doit rappeler, rdv fixé, étude de la candidature, etc.)"
]

def strip_accents(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def normalize_header(s):
    if not s:
        return ""
    s = strip_accents(str(s).lower())
    s = re.sub(r'[^a-z0-9]', '', s)
    return s

# Map normalized header -> canonical key
HEADER_KEYS = {
    normalize_header("Date envoi de la candidature ou du contact"): "date_envoi",
    normalize_header("Date envoi"): "date_envoi",
    normalize_header("Entreprise"): "entreprise",
    normalize_header("Candidature spontanée ou réponse à une offre"): "type_candidature",
    normalize_header("Candidature spontanee ou reponse a une offre"): "type_candidature",
    normalize_header("Type candidature"): "type_candidature",
    normalize_header("Stage ciblé"): "stage",
    normalize_header("Stage cible"): "stage",
    normalize_header("Nom de l'interlocuteur"): "interlocuteur",
    normalize_header("Nom interlocuteur"): "interlocuteur",
    normalize_header("Coordonnées de l'interlocuteur (tel, @)"): "coordonnees",
    normalize_header("Coordonnees de l'interlocuteur"): "coordonnees",
    normalize_header("Coordonnees"): "coordonnees",
    normalize_header("Date d'un éventuel entretien"): "date_entretien",
    normalize_header("Date entretien"): "date_entretien",
    normalize_header("Résultats de la démarche (en attente de réponse, doit rappeler, rdv fixé, étude de la candidature, etc.)"): "resultats",
    normalize_header("Resultats de la demarche"): "resultats",
    normalize_header("Resultats"): "resultats",
}

def create_suivi_excel(entries, poste_global="", type_default="Candidature spontanée"):
    """Génère un Excel suivi avec headers ligne 4, données dès ligne 5 à partir de entries (email+entreprise)"""
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Suivi Candidatures"
    # Lignes 1-3 : Titre + consigne
    ws.merge_cells('A1:H1')
    c1 = ws['A1']
    c1.value = "Suivi des candidatures — Généré par Candidatures Auto"
    c1.font = Font(name='Calibri', size=13, bold=True, color="FFFFFF")
    c1.fill = PatternFill(start_color="4F46E5", end_color="4F46E5", fill_type="solid")
    c1.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 22
    ws['A2'] = "Consigne : Remplissez une ligne par entreprise à partir de la ligne 5. Les en-têtes sont en ligne 4. Ne supprimez pas la ligne 4."
    ws['A2'].font = Font(italic=True, size=9, color="64748B")
    ws['A2'].alignment = Alignment(horizontal='left')
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 6
    # Ligne 4 : headers
    header_fill = PatternFill(start_color="6366F1", end_color="6366F1", fill_type="solid")
    header_font = Font(name='Calibri', size=9, bold=True, color="FFFFFF")
    thin = Side(style="thin", color="E2E8F0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col_idx, h in enumerate(EXCEL_HEADERS, start=1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
    ws.row_dimensions[4].height = 36
    # Largeurs colonnes
    widths = [16, 22, 24, 22, 20, 26, 16, 30]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:H4"
    # Données dès ligne 5
    today = datetime.now().strftime("%d/%m/%Y")
    for row_idx, entry in enumerate(entries, start=5):
        email = entry.get("email","")
        entreprise = entry.get("entreprise") or extract_company_name(email) if email else ""
        # Si entreprise vide (generic) -> laisser vide ou email domain?
        if entreprise == "votre entreprise":
            entreprise = ""
        stage = entry.get("stage") or poste_global or ""
        # Colonnes - Date envoi = date du jour (jj/mm/aaaa), interlocuteur/entretien = "aucun" par défaut si vide
        ws.cell(row=row_idx, column=1, value=today).alignment = Alignment(horizontal='center')
        ws.cell(row=row_idx, column=2, value=entreprise)
        ws.cell(row=row_idx, column=3, value=entry.get("type_candidature") or type_default)
        ws.cell(row=row_idx, column=4, value=stage)
        ws.cell(row=row_idx, column=5, value=entry.get("interlocuteur") or "aucun")
        ws.cell(row=row_idx, column=6, value=email)  # Coordonnées = email (tel peut être ajouté manuellement)
        ws.cell(row=row_idx, column=7, value=entry.get("date_entretien") or "aucun")
        ws.cell(row=row_idx, column=8, value=entry.get("resultats") or "en attente")
        # Style lignes données
        for col in range(1, 9):
            c = ws.cell(row=row_idx, column=col)
            c.font = Font(name='Calibri', size=9)
            c.alignment = Alignment(vertical='center', wrap_text=True, horizontal='center' if col in (1,7) else 'left')
            c.border = border
            # Alternate row color
            if row_idx % 2 == 0:
                c.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
        ws.row_dimensions[row_idx].height = 16
    # Si aucune entrée, laisser vide (l'utilisateur remplira dès ligne 5). Pas d'exemples auto pour ne pas polluer l'import.
    # Print settings
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

def parse_suivi_excel(file_stream):
    """Parse un Excel suivi : headers ligne 4, données dès ligne 5. Retourne (entries, invalid_rows)"""
    import openpyxl
    wb = openpyxl.load_workbook(file_stream, data_only=True, read_only=False)
    ws = wb.active
    # Détecter headers ligne 4
    headers = []
    col_map = {}  # col_idx -> key
    max_col = min(ws.max_column, len(EXCEL_HEADERS)+2)
    for col in range(1, max_col+1):
        val = ws.cell(row=4, column=col).value
        if val:
            norm = normalize_header(val)
            key = HEADER_KEYS.get(norm)
            # fallback fuzzy: si contient "entreprise" etc.
            if not key:
                if "entreprise" in norm:
                    key = "entreprise"
                elif "coordonn" in norm:
                    key = "coordonnees"
                elif "stage" in norm:
                    key = "stage"
                elif "interlocuteur" in norm and "coordonn" not in norm:
                    key = "interlocuteur"
                elif "resulta" in norm:
                    key = "resultats"
                elif "entretien" in norm:
                    key = "date_entretien"
                elif "envoi" in norm:
                    key = "date_envoi"
                elif "candidature" in norm and "offre" in norm:
                    key = "type_candidature"
            if key:
                col_map[col] = key
            headers.append((col, val, key))
    if not col_map:
        # Fallback : supposer ordre fixe EXCEL_HEADERS
        for i in range(1, len(EXCEL_HEADERS)+1):
            # Map par position
            keys_order = ["date_envoi","entreprise","type_candidature","stage","interlocuteur","coordonnees","date_entretien","resultats"]
            col_map[i] = keys_order[i-1]
    entries = []
    invalid_rows = []
    email_regex = re.compile(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+')
    for row in range(5, ws.max_row+1):
        # Vérifier si ligne vide
        is_empty = True
        row_data = {}
        for col, key in col_map.items():
            v = ws.cell(row=row, column=col).value
            if v is not None and str(v).strip() != "":
                is_empty = False
            row_data[key] = str(v).strip() if v is not None else ""
        if is_empty:
            continue
        # Extraire email depuis coordonnees
        coord = row_data.get("coordonnees","")
        m = email_regex.search(coord or "")
        email = m.group(0).lower() if m else ""
        # Si pas d'email dans coordonnees, chercher dans toute la ligne
        if not email:
            for v in row_data.values():
                mm = email_regex.search(v or "")
                if mm:
                    email = mm.group(0).lower()
                    break
        if not email:
            invalid_rows.append({"row": row, "reason": "Aucun email trouvé en colonne Coordonnées", "data": row_data})
            continue
        entreprise = row_data.get("entreprise","").strip()
        if not entreprise:
            entreprise = extract_company_name(email)
            if entreprise == "votre entreprise":
                entreprise = ""
        entries.append({
            "email": email,
            "entreprise": entreprise,
            "stage": row_data.get("stage",""),
            "type_candidature": row_data.get("type_candidature",""),
            "interlocuteur": row_data.get("interlocuteur",""),
            "coordonnees": coord,
            "date_envoi": row_data.get("date_envoi",""),
            "date_entretien": row_data.get("date_entretien",""),
            "resultats": row_data.get("resultats",""),
        })
    # Dédup par email
    seen = {}
    uniq = []
    for e in entries:
        if e["email"] not in seen:
            seen[e["email"]] = True
            uniq.append(e)
    return uniq, invalid_rows, headers

def parse_entries(raw_text):
    """
    Parse une liste d'entrées où chaque entrée peut contenir un nom d'entreprise + email.
    Formats supportés (un par ligne recommandé) :
      - rh@capgemini.fr
      - Capgemini <rh@capgemini.fr>
      - Capgemini | rh@capgemini.fr
      - rh@capgemini.fr | Capgemini
      - Capgemini ; rh@capgemini.fr
      - rh@capgemini.fr ; Capgemini
      - Dassault Systemes | jobs@dassault.com
    Retourne (entries, invalid) où entries = list of {email, entreprise}
    """
    if not raw_text or not raw_text.strip():
        return [], []
    email_regex = re.compile(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+')
    full_email_regex = re.compile(r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$')
    entries = []
    invalid = []

    # Split by newlines first, then also handle comma-separated entries on same line if they contain multiple emails
    # We treat newline as primary delimiter. For each line, if it contains multiple emails separated by commas not in <> form, we split.
    lines = re.split(r'[\r\n]+', raw_text.strip())
    raw_candidates = []
    for line in lines:
        line = line.strip()
        if not line:
            continue
        # If line contains '<' we treat it as single entry (to not split "A, B <email>")
        if '<' in line and '>' in line:
            raw_candidates.append(line)
        elif line.count('@') > 1:
            # multiple emails on same line, split by comma/semicolon
            parts = re.split(r'[,\;]+', line)
            for p in parts:
                p = p.strip()
                if p:
                    raw_candidates.append(p)
        elif ',' in line and line.count('@') == 1 and '|' not in line and ';' not in line:
            # Single email but comma inside company name? ambiguous. Treat as single entry if company name before comma? 
            # We check if splitting by comma yields one email part and one name part
            # For safety, if line like "Capgemini, rh@capgemini.fr" we interpret correctly
            # Use detection: split by comma, if one part is email, other is name -> keep as single entry
            if re.search(email_regex, line):
                # Contains email, keep whole line as single entry if it looks like "Name, email" (name before comma)
                # To support both, we keep whole line; parsing below will extract correctly
                raw_candidates.append(line)
            else:
                raw_candidates.extend([p.strip() for p in line.split(',') if p.strip()])
        else:
            raw_candidates.append(line)

    for cand in raw_candidates:
        cand = cand.strip()
        if not cand:
            continue
        # Extract email
        m = email_regex.search(cand)
        if not m:
            # No email found, invalid
            # Also try splitting by spaces and testing each token as full email
            invalid.append(cand)
            continue
        email = m.group(0).lower()
        # Validate full email
        if not full_email_regex.match(email):
            invalid.append(cand)
            continue
        # Enterprise = cand without email, clean delimiters
        entreprise_raw = cand.replace(m.group(0), '')
        # Remove common delimiters and brackets
        entreprise_raw = re.sub(r'[<>\|\;\:,"]', ' ', entreprise_raw)
        entreprise_raw = re.sub(r'\s+', ' ', entreprise_raw).strip()
        # If entreprise_raw empty, keep as "" (fallback later)
        entries.append({"email": email, "entreprise": entreprise_raw})

    # Deduplicate by email preserve order, keep first entreprise
    seen = {}
    uniq = []
    for e in entries:
        if e["email"] not in seen:
            seen[e["email"]] = e["entreprise"]
            uniq.append(e)
        else:
            # if existing had no entreprise and new has one, update
            if not seen[e["email"]] and e["entreprise"]:
                for u in uniq:
                    if u["email"] == e["email"]:
                        u["entreprise"] = e["entreprise"]
                        break
                seen[e["email"]] = e["entreprise"]
    return uniq, invalid

def parse_emails(raw_text):
    # Backward compat: returns list of emails only
    entries, invalid = parse_entries(raw_text)
    emails = [e["email"] for e in entries]
    return emails, invalid

def generate_letter_content(data):
    """
    Génère le contenu texte de la lettre à partir des infos utilisateur.
    Si pas d'IA, on utilise un template intelligent.
    """
    prenom = data.get('prenom', '').strip() or '[Prénom]'
    nom = data.get('nom', '').strip() or '[Nom]'
    poste = data.get('poste', '').strip() or 'le poste proposé'
    entreprise_placeholder = data.get('entreprisePlaceholder', 'votre entreprise')
    ville = data.get('ville', '').strip()
    telephone = data.get('telephone', '').strip()
    email = data.get('emailCandidat', '').strip()
    experience = data.get('experience', '').strip()
    motivation = data.get('motivation', '').strip()
    ton = data.get('ton', 'professionnel')  # professionnel, dynamique, sobre
    dispo = data.get('disponibilite', '').strip()

    # Extraction du nom d'entreprise à partir de l'email pour personnalisation
    # sera fait côté envoi si option personnalisée

    date_str = datetime.now().strftime("%d/%m/%Y")

    # Templates selon ton
    intro_tons = {
        'professionnel': "C'est avec un vif intérêt que je vous adresse ma candidature",
        'dynamique': "Passionné(e) et motivé(e), je souhaite vous proposer ma candidature",
        'sobre': "Je vous adresse ma candidature",
        'creatif': "Attiré(e) par l'univers de votre entreprise, je vous propose ma candidature"
    }
    intro = intro_tons.get(ton, intro_tons['professionnel'])

    lettre = f"""{prenom} {nom}
{ville + ' - ' if ville else ''}{telephone + ' - ' if telephone else ''}{email}

À l'attention du service Recrutement
{{ENTREPRISE_NOM}}

Le {date_str}

Objet : Candidature au poste de {poste}

Madame, Monsieur,

{intro} au poste de {poste} au sein de {{ENTREPRISE_NOM}}.

"""

    if experience:
        lettre += f"{experience}\n\n"
    else:
        lettre += f"""Titulaire d'un parcours en adéquation avec les exigences du poste de {poste}, je suis convaincu(e) de pouvoir apporter une réelle valeur ajoutée à vos équipes. Rigoureux(se), autonome et doté(e) d'un excellent relationnel, je m'intègre rapidement et m'investis pleinement dans les missions qui me sont confiées.

"""

    if motivation:
        lettre += f"{motivation}\n\n"
    else:
        lettre += f"""Votre entreprise, reconnue pour son expertise et ses valeurs, représente pour moi une opportunité idéale de mettre mes compétences au service d'un projet ambitieux. Rejoindre {{ENTREPRISE_NOM}} serait pour moi l'occasion de contribuer activement à votre développement tout en poursuivant mon évolution professionnelle.

"""

    lettre += f"""Disponible {dispo if dispo else 'immédiatement'}, je serais honoré(e) d'échanger avec vous lors d'un entretien afin de vous exposer plus en détail mes motivations.

En vous remerciant de l'attention portée à ma candidature, je vous prie d'agréer, Madame, Monsieur, l'expression de mes salutations distinguées.

{prenom} {nom}
"""

    return lettre

def extract_company_name(email):
    """ Essaie de deviner le nom d'entreprise depuis l'email: contact@capgemini.fr -> Capgemini """
    try:
        domain = email.split('@')[1].split('.')[0]
        # Si domaine générique (gmail etc.), on ne peut pas deviner
        if domain.lower() in GENERIC_DOMAINS:
            return "votre entreprise"
        # nettoyage
        domain = re.sub(r'[^a-zA-Z0-9]', ' ', domain)
        # Capitalize chaque mot
        return ' '.join(w.capitalize() for w in domain.split())
    except:
        return "votre entreprise"

def create_pdf_letter(content_text, output_path, candidat_infos):
    """ Génère un PDF lettre avec fpdf2 """
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    # En-tête
    pdf.set_font("Helvetica", "B", 11)
    prenom = candidat_infos.get('prenom','')
    nom = candidat_infos.get('nom','')
    pdf.cell(0, 7, f"{prenom} {nom}", new_x="LMARGIN", new_y="NEXT", align="L")
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(80,80,80)
    contact_line = " - ".join(filter(None, [candidat_infos.get('ville',''), candidat_infos.get('telephone',''), candidat_infos.get('emailCandidat','')]))
    if contact_line:
        pdf.cell(0, 5, contact_line, new_x="LMARGIN", new_y="NEXT")
    pdf.set_text_color(0,0,0)
    pdf.ln(8)

    # Corps - on remplace les lignes
    pdf.set_font("Helvetica", "", 10)
    for line in content_text.split('\n'):
        if not line.strip():
            pdf.ln(4)
            continue
        # Détection objet / date
        if line.startswith("Objet :") or line.startswith("Le ") or line.startswith("À l'attention"):
            pdf.set_font("Helvetica", "B", 10) if "Objet" in line else pdf.set_font("Helvetica", "", 9)
            if "À l'attention" in line:
                pdf.set_text_color(40,40,40)
            pdf.multi_cell(0, 5, line, new_x="LMARGIN", new_y="NEXT")
            pdf.set_font("Helvetica", "", 10)
            pdf.set_text_color(0,0,0)
        else:
            pdf.multi_cell(0, 5.5, line, new_x="LMARGIN", new_y="NEXT")

    pdf.output(str(output_path))


@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/generate-letter", methods=["POST"])
def api_generate_letter():
    data = request.get_json() or {}
    # validation minimale
    if not data.get('prenom') or not data.get('nom') or not data.get('poste'):
        return jsonify({"error": "Prénom, nom et poste visé sont requis"}), 400

    lettre_template = generate_letter_content(data)

    # Si personnalisation par entreprise demandée, on renvoie le template avec placeholder
    # sinon on remplace par générique
    personnalise = data.get('personnalise', False)

    # Pour la preview, on montre avec le placeholder remplacé par "votre entreprise"
    preview = lettre_template.replace("{ENTREPRISE_NOM}", data.get('entreprisePlaceholder') or "votre entreprise")

    # Génération PDF preview
    preview_path = GENERATED_FOLDER / f"lettre_preview_{uuid.uuid4().hex[:8]}.pdf"
    # on remplace placeholder pour PDF preview
    create_pdf_letter(preview, preview_path, data)

    return jsonify({
        "template": lettre_template,  # avec {ENTREPRISE_NOM} placeholder
        "preview": preview,
        "pdf_url": f"/generated/{preview_path.name}"
    })

@app.route("/generated/<filename>")
def serve_generated(filename):
    f = GENERATED_FOLDER / filename
    if f.exists():
        return send_file(f, mimetype="application/pdf")
    return jsonify({"error": "Fichier non trouvé"}), 404

@app.route("/api/send", methods=["POST"])
def api_send():
    # multipart/form-data
    # fields: smtp_host, smtp_port, smtp_user, smtp_pass, smtp_secure, from_name, subject, message, emails_text, lettre_template, candidat infos...
    smtp_host = request.form.get('smtp_host', '').strip()
    smtp_port = request.form.get('smtp_port', '').strip()
    smtp_user = request.form.get('smtp_user', '').strip()
    smtp_pass = request.form.get('smtp_pass', '').strip()
    smtp_secure = request.form.get('smtp_secure', 'tls')  # tls, ssl, none
    from_name = request.form.get('from_name', '').strip() or f"{request.form.get('prenom','')} {request.form.get('nom','')}"
    subject = request.form.get('subject', '').strip() or f"Candidature - {request.form.get('poste','')}"
    message_body = request.form.get('message_body', '').strip()
    emails_text = request.form.get('emails_text', '')
    lettre_template = request.form.get('lettre_template', '')
    poste = request.form.get('poste', '')
    delay = float(request.form.get('delay', '1.5') or 1.5)

    # infos candidat pour PDF
    candidat_infos = {
        'prenom': request.form.get('prenom',''),
        'nom': request.form.get('nom',''),
        'ville': request.form.get('ville',''),
        'telephone': request.form.get('telephone',''),
        'emailCandidat': request.form.get('emailCandidat',''),
    }

    if not smtp_host or not smtp_port or not smtp_user:
        return jsonify({"error": "Configuration SMTP incomplète (hôte, port, utilisateur requis)"}), 400
    # smtp_pass peut être vide pour serveurs locaux sans auth (ex: smtp dummy test)
    if not smtp_pass and smtp_secure != 'none':
        # on autorise quand même mais on prévient
        pass

    # Parse avec noms d'entreprise si fournis - support Excel (entries enrichies depuis tableau ligne 4+)
    excel_entries_raw = request.form.get('excel_entries', '')
    if excel_entries_raw:
        try:
            excel_entries = json.loads(excel_entries_raw)
            # excel_entries doit être list of dicts avec email, entreprise, stage...
            entries = []
            invalid = []
            email_regex = re.compile(r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$')
            for en in excel_entries:
                em = (en.get("email") or "").strip().lower()
                if not em or not email_regex.match(em):
                    invalid.append(str(en))
                    continue
                entries.append({
                    "email": em,
                    "entreprise": (en.get("entreprise") or "").strip(),
                    "stage": (en.get("stage") or "").strip(),
                    "type_candidature": (en.get("type_candidature") or "").strip(),
                })
            if not entries:
                return jsonify({"error": "Aucune adresse email valide dans le tableau Excel", "invalid": invalid}), 400
        except Exception as e:
            return jsonify({"error": f"Excel entries invalide: {str(e)}"}), 400
    else:
        entries, invalid = parse_entries(emails_text)
        if not entries:
            return jsonify({"error": "Aucune adresse email valide fournie", "invalid": invalid}), 400
    emails = [e["email"] for e in entries]

    # CV file
    if 'cv' not in request.files:
        return jsonify({"error": "CV manquant"}), 400
    cv_file = request.files['cv']
    if cv_file.filename == '':
        return jsonify({"error": "Aucun fichier CV sélectionné"}), 400
    if not allowed_file(cv_file.filename):
        return jsonify({"error": "Format CV non supporté. Utilisez PDF, DOC ou DOCX"}), 400

    # Sauvegarde CV temporaire
    cv_filename = secure_filename(cv_file.filename)
    cv_path = UPLOAD_FOLDER / f"{uuid.uuid4().hex}_{cv_filename}"
    cv_file.save(cv_path)

    # Vérif lettre template
    if not lettre_template:
        # générer à partir des infos
        data_for_letter = {
            'prenom': request.form.get('prenom',''),
            'nom': request.form.get('nom',''),
            'poste': poste,
            'ville': request.form.get('ville',''),
            'telephone': request.form.get('telephone',''),
            'emailCandidat': request.form.get('emailCandidat',''),
            'experience': request.form.get('experience',''),
            'motivation': request.form.get('motivation',''),
            'ton': request.form.get('ton','professionnel'),
            'disponibilite': request.form.get('disponibilite',''),
            'entreprisePlaceholder': 'votre entreprise'
        }
        lettre_template = generate_letter_content(data_for_letter)

    # Options personnalisation
    personnalise = request.form.get('personnalise') == 'true'
    use_custom_names = request.form.get('useCustomNames') == 'true'
    # Map email -> entreprise saisie manuellement
    custom_map = {e["email"]: e["entreprise"] for e in entries if e["entreprise"]}

    # Test connexion SMTP
    try:
        smtp_port_int = int(smtp_port)
    except:
        return jsonify({"error": "Port SMTP invalide"}), 400

    results = []
    # Connexion une fois
    server = None
    try:
        if smtp_secure == 'ssl':
            context = ssl.create_default_context()
            server = smtplib.SMTP_SSL(smtp_host, smtp_port_int, context=context, timeout=15)
        else:
            server = smtplib.SMTP(smtp_host, smtp_port_int, timeout=15)
            if smtp_secure == 'tls':
                context = ssl.create_default_context()
                server.starttls(context=context)
        # Login seulement si identifiants fournis et serveur supporte AUTH
        if smtp_user and smtp_pass:
            try:
                server.login(smtp_user, smtp_pass)
            except smtplib.SMTPException as login_err:
                # Si le serveur ne supporte pas AUTH (dummy local), on continue sans login
                if "AUTH" in str(login_err):
                    print(f"[INFO] SMTP AUTH non supporté, envoi sans authentification: {login_err}")
                else:
                    raise
    except Exception as e:
        cv_path.unlink(missing_ok=True)
        return jsonify({"error": f"Connexion SMTP échouée: {str(e)}"}), 400

    # Envoi boucle
    for idx, entry in enumerate(entries):
        dest = entry["email"]
        try:
            # Génération lettre personnalisée
            if not personnalise:
                entreprise_nom = "votre entreprise"
            elif use_custom_names and entry.get("entreprise"):
                entreprise_nom = entry["entreprise"]
            else:
                entreprise_nom = extract_company_name(dest)
            # Stage ciblé par entreprise (depuis Excel ligne 5+ col 4) sinon poste global
            poste_entry = (entry.get("stage") or "").strip() or poste
            lettre_personnalisee = lettre_template.replace("{ENTREPRISE_NOM}", entreprise_nom)
            # Si stage par entreprise différent du poste global, on personnalise aussi le poste dans la lettre
            if poste and poste_entry != poste:
                # Remplace le poste global par le stage spécifique dans la lettre (toutes occurrences)
                lettre_personnalisee = lettre_personnalisee.replace(poste, poste_entry)

            # Génération PDF lettre pour cet envoi
            safe_entreprise = re.sub(r'[^a-zA-Z0-9_-]', '_', entreprise_nom)[:30] or "entreprise"
            lettre_pdf_path = GENERATED_FOLDER / f"Lettre_Motivation_{safe_entreprise}_{uuid.uuid4().hex[:6]}.pdf"
            create_pdf_letter(lettre_personnalisee, lettre_pdf_path, candidat_infos)

            # Message email
            msg = MIMEMultipart()
            msg['From'] = f"{from_name} <{smtp_user}>"
            msg['To'] = dest
            msg['Date'] = formatdate(localtime=True)
            # Sujet : personnalise {ENTREPRISE} et aussi poste si différent par ligne
            subj = subject
            if "{ENTREPRISE}" in subj:
                subj = subj.replace("{ENTREPRISE}", entreprise_nom)
            if poste and poste_entry != poste and poste in subj:
                subj = subj.replace(poste, poste_entry)
            msg['Subject'] = subj

            # Corps
            body_text = message_body or f"""Madame, Monsieur,

Veuillez trouver ci-joint ma candidature au poste de {poste_entry}.

Je me permets de vous adresser mon CV ainsi que ma lettre de motivation.

Dans l'attente de votre retour, je vous prie d'agréer, Madame, Monsieur, l'expression de mes salutations distinguées.

Cordialement,
{candidat_infos['prenom']} {candidat_infos['nom']}
{candidat_infos['telephone']}
{candidat_infos['emailCandidat']}
"""
            # Remplacer placeholders entreprise + poste par ligne
            body_text = body_text.replace("{ENTREPRISE}", entreprise_nom)
            if poste and poste_entry != poste:
                body_text = body_text.replace(poste, poste_entry)

            msg.attach(MIMEText(body_text, 'plain', 'utf-8'))

            # Pièces jointes CV
            with open(cv_path, 'rb') as f:
                part_cv = MIMEApplication(f.read(), Name=cv_filename)
                part_cv['Content-Disposition'] = f'attachment; filename="{cv_filename}"'
                msg.attach(part_cv)

            # Pièce jointe lettre
            with open(lettre_pdf_path, 'rb') as f:
                part_lettre = MIMEApplication(f.read(), Name=lettre_pdf_path.name)
                part_lettre['Content-Disposition'] = f'attachment; filename="{lettre_pdf_path.name}"'
                msg.attach(part_lettre)

            server.sendmail(smtp_user, dest, msg.as_string())
            results.append({"email": dest, "status": "success", "entreprise": entreprise_nom})
            # cleanup lettre pdf après envoi? on garde pour logs mais on peut supprimer
            lettre_pdf_path.unlink(missing_ok=True)

            if idx < len(entries) - 1:
                time.sleep(delay)

        except Exception as e:
            results.append({"email": dest, "status": "error", "error": str(e)})

    try:
        server.quit()
    except:
        pass

    cv_path.unlink(missing_ok=True)

    success_count = sum(1 for r in results if r['status']=='success')
    failed = [r for r in results if r['status']=='error']

    return jsonify({
        "total": len(emails),
        "success": success_count,
        "failed": len(failed),
        "invalid_emails": invalid,
        "results": results
    })

@app.route("/api/parse-emails", methods=["POST"])
def api_parse_emails():
    data = request.get_json() or {}
    raw = data.get('text','')
    entries, invalid = parse_entries(raw)
    emails = [e["email"] for e in entries]
    return jsonify({"emails": emails, "entries": entries, "invalid": invalid, "count": len(emails)})

@app.route("/api/excel-template", methods=["GET"])
def api_excel_template():
    """Télécharge le modèle Excel vide (headers ligne 4)"""
    out = create_suivi_excel([], poste_global="")
    return send_file(out, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name="modele_suivi_candidatures.xlsx")

@app.route("/api/generate-excel", methods=["POST"])
def api_generate_excel():
    """Génère le tableau de suivi à partir de la liste d'emails fournie au site (remplit ligne 5+)."""
    data = request.get_json() or {}
    emails_text = data.get('emails_text', '')
    poste = data.get('poste', '')
    entries, invalid = parse_entries(emails_text)
    if not entries:
        return jsonify({"error": "Aucune adresse email valide fournie", "invalid": invalid}), 400
    # Enrichir avec poste global si stage vide
    for e in entries:
        if not e.get("stage") and poste:
            e["stage"] = poste
    out = create_suivi_excel(entries, poste_global=poste)
    return send_file(out, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name=f"suivi_candidatures_{datetime.now().strftime('%Y%m%d')}.xlsx")

@app.route("/api/parse-excel", methods=["POST"])
def api_parse_excel():
    """Parse un Excel de suivi : headers ligne 4, données dès ligne 5 → retourne entries pour remplir le site"""
    if 'file' not in request.files:
        return jsonify({"error": "Aucun fichier fourni"}), 400
    f = request.files['file']
    if f.filename == '':
        return jsonify({"error": "Fichier vide"}), 400
    if not f.filename.lower().endswith(('.xlsx','.xlsm','.xls')):
        return jsonify({"error": "Format non supporté, utilisez .xlsx"}), 400
    try:
        entries, invalid_rows, headers = parse_suivi_excel(f.stream)
    except Exception as e:
        return jsonify({"error": f"Erreur lecture Excel: {str(e)}"}), 400
    # Convertir en format compatible avec emails_text (pour remplir textarea)
    # On reconstruit emails_text avec "Entreprise <email>" si entreprise présente
    lines = []
    for e in entries:
        if e.get("entreprise"):
            lines.append(f"{e['entreprise']} <{e['email']}>")
        else:
            lines.append(e['email'])
    emails_text = "\n".join(lines)
    return jsonify({
        "entries": entries,
        "count": len(entries),
        "invalid_rows": invalid_rows,
        "headers": headers,
        "emails_text": emails_text,
        "emails": [e["email"] for e in entries]
    })

if __name__ == "__main__":
    print("→ Serveur candidature-auto sur http://localhost:5000")
    # debug=False pour stabilité en production, use_reloader désactivé pour éviter double process
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=5000, debug=debug, use_reloader=debug)

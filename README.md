# Candidatures Auto — Envoi automatique CV + Lettre de motivation

Site web complet (Flask + JS vanilla) qui permet de candidater en masse : tu renseignes une liste d'emails d'entreprises, ton CV, et le site génère une lettre de motivation personnalisée + envoie les mails avec les 2 pièces jointes.

## 📁 Structure
```
/home/lapinou/candidatures-auto/
├── app.py                 # Backend Flask (génération PDF + envoi SMTP)
├── templates/index.html   # Frontend unique page
├── static/css/style.css   # Style
├── static/js/app.js       # Logique frontend
├── requirements.txt
├── start.sh               # Lancement rapide
├── uploads/               # CV temporaires
└── generated/             # PDF lettres générées
```

## 🚀 Lancement rapide
```bash
cd /home/lapinou/candidatures-auto
./start.sh
# ou
pip install --break-system-packages -r requirements.txt
python3 app.py
```
Ouvre **http://localhost:5000**

Testé avec succès : `parse_emails`, `generate_letter`, `create_pdf`, `send` (mock SMTP) — tous OK.

## 🔧 Fonctionnalités
- **Profil candidat** : prénom, nom, poste visé, ville, tél, expérience, motivation, ton (professionnel/dynamique/sobre/créatif), disponibilité
- **CV** : drag & drop PDF/DOC/DOCX (10 Mo), prévisualisation
- **Lettre de motivation** : génération template intelligent avec placeholder `{ENTREPRISE_NOM}` → personnalisation auto par domaine email (ex: `rh@capgemini.fr` → `Capgemini`). Éditable avant envoi + prévisualisation texte + PDF (fpdf2).
- **Liste entreprises** : textarea (virgule/point-virgule/saut de ligne), déduplication, validation regex, import CSV/TXT, compteur temps réel
- **Email** : config SMTP complète (host/port/user/pass/secure), objet avec `{ENTREPRISE}`, corps personnalisable, délai anti-spam (1.5s), envoi avec CV + lettre en pièces jointes, suivi temps réel (progress bar + résultats par email)

## 📧 Configuration SMTP

### Gmail (recommandé)
1. Active la validation 2 étapes : https://myaccount.google.com/signinoptions/two-step-verification
2. Crée un **App Password** : https://myaccount.google.com/apppasswords → choisis "Mail", copie les 16 caractères
3. Dans le site :
   - Hôte : `smtp.gmail.com`
   - Port : `587`
   - Sécurité : `STARTTLS`
   - Utilisateur : `ton.email@gmail.com`
   - Mot de passe : `le app password` (sans espaces)

### Autres (Outlook, OVH, etc.)
- Outlook : `smtp.office365.com:587` STARTTLS
- OVH : `ssl0.ovh.net:465` SSL

> Astuce : garde un délai ≥1s et ≤100 envois/jour avec Gmail pour éviter le spam.

## 🧪 API
- `POST /api/generate-letter` JSON `{prenom, nom, poste, ...}` → `{preview, template, pdf_url}`
- `POST /api/parse-emails` JSON `{text}` → `{emails, invalid, count}`
- `POST /api/send` multipart (cv + smtp_* + emails_text + lettre_template + ...) → `{total, success, failed, results[]}`
- `GET /generated/<pdf>` → PDF

## 📝 Flux utilisateur
1. Renseigne profil + dépose CV
2. Clique **Générer / Prévisualiser** → édite la lettre si besoin (garde `{ENTREPRISE_NOM}`)
3. Colle liste emails ou importe CSV
4. Configure SMTP + objet/corps
5. **Envoyer** → suivi live, chaque lettre est régénérée en PDF personnalisé par entreprise

## ⚠️ Notes
- Données ne quittent pas ta machine sauf envoi SMTP
- Nécessite `FLASK_DEBUG=1` pour reloader auto : `FLASK_DEBUG=1 python3 app.py`
- Pour test local sans vrai SMTP : utilise `python3 /tmp/dummy_smtp.py` (aiosmtpd sur 1025) avec `smtp_secure=none`

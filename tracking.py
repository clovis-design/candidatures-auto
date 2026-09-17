"""Suivi local persistant. Aucun identifiant SMTP n'est enregistré."""
import json
import sqlite3
import re
import unicodedata
from contextlib import contextmanager
from datetime import datetime, timedelta
from urllib.parse import urlsplit

from spreadsheets import EMAIL

STATUSES = {
    'draft': 'À envoyer',
    'sent': 'En attente de réponse',
    'follow_up': 'À relancer',
    'interview': 'Entretien',
    'accepted': 'Acceptée',
    'rejected': 'Refusée',
    'error': 'Échec d’envoi',
    'sending': 'Envoi en cours',
}
SUBMITTED_STATUSES = {'sent', 'follow_up', 'interview', 'accepted', 'rejected'}
EXTERNAL_ORIGINS = {'manual', 'spreadsheet'}


class DuplicateApplication(ValueError):
    pass


def normalized(value):
    text = unicodedata.normalize('NFD', str(value or '').casefold())
    return re.sub(r'[^a-z0-9]', '', ''.join(c for c in text if unicodedata.category(c) != 'Mn'))


def tracking_date(value, label):
    value = str(value or '').strip()
    if normalized(value) in ('', 'aucun', 'aucune', 'non'):
        return ''
    for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%Y-%m-%d %H:%M:%S', '%Y-%m-%dT%H:%M:%S'):
        try:
            return datetime.strptime(value, fmt).date().isoformat()
        except ValueError:
            continue
    raise ValueError(f'{label} : date invalide, utilisez JJ/MM/AAAA ou AAAA-MM-JJ.')


def external_entry(data, origin):
    """Valide une candidature de suivi sans inventer de date d'envoi."""
    if not isinstance(data, dict):
        raise ValueError('Les informations de candidature doivent être un objet.')
    entry = {key: str(data.get(key) or '').strip() for key in (
        'entreprise', 'stage', 'email', 'interlocuteur', 'coordonnees',
        'type_candidature', 'notes', 'source', 'url')}
    if not entry['entreprise'] or not entry['stage']:
        raise ValueError('Entreprise et poste sont requis.')
    if entry['email'] and not EMAIL.fullmatch(entry['email']):
        raise ValueError('Adresse email invalide ; laissez ce champ vide si elle est inconnue.')
    entry['email'] = entry['email'].lower()
    if any(len(value) > (10000 if key == 'notes' else 2000) for key, value in entry.items()):
        raise ValueError('Champ trop long (2 000 caractères, ou 10 000 pour les notes).')
    if entry['url']:
        try:
            url = urlsplit(entry['url'])
            if url.scheme not in ('http', 'https') or not url.hostname or any(c.isspace() for c in entry['url']):
                raise ValueError()
        except ValueError:
            raise ValueError('Le lien de l’offre doit être une URL http:// ou https:// valide.')
    entry['date_envoi'] = tracking_date(data.get('date_envoi'), 'Date de candidature')
    entry['date_entretien'] = tracking_date(data.get('date_entretien'), 'Date d’entretien')
    entry['follow_up_date'] = tracking_date(data.get('follow_up_date'), 'Date de relance')
    aliases = {normalized(label): key for key, label in STATUSES.items() if key != 'sending'}
    aliases.update({normalized(key): key for key in STATUSES if key != 'sending'})
    aliases.update({
        'enattente': 'sent', 'sansreponse': 'sent', 'aucunereponse': 'sent',
        'envoye': 'sent', 'envoyee': 'sent', 'candidatureenvoyee': 'sent',
        'etudedelacandidature': 'sent', 'encours': 'sent',
        'relancer': 'follow_up', 'doitrappeler': 'follow_up',
        'rdvfixe': 'interview', 'entretienprevu': 'interview',
        'accepte': 'accepted', 'refuse': 'rejected', 'refus': 'rejected',
        'brouillon': 'draft', 'apreparer': 'draft', 'echec': 'error',
    })
    raw_status = str(data.get('status') or data.get('resultats') or '').strip()
    status = aliases.get(normalized(raw_status)) if raw_status else ('sent' if entry['date_envoi'] else 'draft')
    if not status:
        raise ValueError(f'Statut non reconnu : « {raw_status} ». Utilisez À envoyer, En attente, À relancer, Entretien, Acceptée, Refusée ou Échec d’envoi.')
    if status in ('draft', 'error') and entry['date_envoi']:
        raise ValueError('Une candidature à envoyer ou en échec ne peut pas avoir de date d’envoi.')
    entry['status'] = status
    entry['origin'] = origin
    entry['source'] = entry['source'] or ('Saisie manuelle' if origin == 'manual' else 'Tableur')
    return entry


class TrackingStore:
    def __init__(self, path):
        self.path = str(path)
        with self.connect() as db:
            db.execute('BEGIN IMMEDIATE')
            db.execute('''CREATE TABLE IF NOT EXISTS applications (
                id INTEGER PRIMARY KEY, email TEXT, stage TEXT NOT NULL,
                details TEXT NOT NULL, status TEXT NOT NULL, sent_at TEXT DEFAULT '',
                updated_at TEXT NOT NULL, follow_up_date TEXT DEFAULT '',
                notes TEXT DEFAULT '', error TEXT DEFAULT '',
                UNIQUE(email, stage))''')
            # Migration des historiques existants : NULL permet plusieurs sociétés sans email.
            columns = db.execute('PRAGMA table_info(applications)').fetchall()
            if next(column for column in columns if column['name'] == 'email')['notnull']:
                db.execute('''CREATE TABLE applications_migrated (
                    id INTEGER PRIMARY KEY, email TEXT, stage TEXT NOT NULL,
                    details TEXT NOT NULL, status TEXT NOT NULL, sent_at TEXT DEFAULT '',
                    updated_at TEXT NOT NULL, follow_up_date TEXT DEFAULT '',
                    notes TEXT DEFAULT '', error TEXT DEFAULT '', UNIQUE(email, stage))''')
                db.execute('INSERT INTO applications_migrated SELECT * FROM applications')
                db.execute('DROP TABLE applications')
                db.execute('ALTER TABLE applications_migrated RENAME TO applications')

    @contextmanager
    def connect(self):
        db = sqlite3.connect(self.path, timeout=15)
        db.row_factory = sqlite3.Row
        try:
            with db:
                yield db
        finally:
            db.close()

    def reserve(self, entry):
        """Réservation atomique : une même candidature ne part pas deux fois."""
        now = datetime.now().isoformat(timespec='seconds')
        with self.connect() as db:
            db.execute('BEGIN IMMEDIATE')
            old = self._duplicate(db, entry)
            if old and (old['sent_at'] or old['status'] in SUBMITTED_STATUSES | {'sending'}):
                return None
            if old:
                previous = json.loads(old['details'])
                entry = {**previous, **entry, 'origin': 'smtp', 'source': 'Email SMTP'}
                db.execute("UPDATE applications SET email=?, stage=?, details=?, status='sending', error='', updated_at=? WHERE id=?",
                           (entry['email'], entry['stage'], json.dumps(entry, ensure_ascii=False), now, old['id']))
                return old['id']
            cursor = db.execute('INSERT INTO applications (email, stage, details, status, updated_at, notes) VALUES (?,?,?,?,?,?)',
                                (entry['email'], entry['stage'], json.dumps(entry, ensure_ascii=False), 'sending', now, entry.get('notes', '')))
            return cursor.lastrowid

    @staticmethod
    def _duplicate(db, entry, exclude_id=None):
        # Email + poste ; sans email d'un côté, entreprise + poste.
        for row in db.execute('SELECT * FROM applications'):
            if row['id'] == exclude_id or row['stage'].strip().casefold() != entry['stage'].strip().casefold():
                continue
            email = entry.get('email', '')
            if email and row['email']:
                if email.casefold() == row['email'].casefold():
                    return row
            elif entry.get('entreprise', '').strip() and str(json.loads(row['details']).get('entreprise', '')).strip().casefold() == entry['entreprise'].strip().casefold():
                return row
        return None

    def create_external(self, data, origin='manual'):
        entry = external_entry(data, origin)
        with self.connect() as db:
            db.execute('BEGIN IMMEDIATE')
            if self._duplicate(db, entry):
                raise DuplicateApplication('Cette candidature figure déjà dans le suivi (même contact ou entreprise et même poste).')
            cursor = db.execute('''INSERT INTO applications
                (email, stage, details, status, sent_at, updated_at, follow_up_date, notes)
                VALUES (?,?,?,?,?,?,?,?)''', (entry['email'] or None, entry['stage'], json.dumps(entry, ensure_ascii=False),
                entry['status'], entry['date_envoi'], datetime.now().isoformat(timespec='seconds'), entry['follow_up_date'], entry['notes']))
            return cursor.lastrowid

    def finish(self, record_id, error='', follow_up_days=7):
        now = datetime.now()
        with self.connect() as db:
            db.execute('UPDATE applications SET status=?, sent_at=?, updated_at=?, follow_up_date=?, error=? WHERE id=?',
                       ('error' if error else 'sent', '' if error else now.isoformat(timespec='seconds'),
                        now.isoformat(timespec='seconds'), '' if error else (now + timedelta(days=follow_up_days)).date().isoformat(),
                        error, record_id))

    def all(self):
        with self.connect() as db:
            rows = db.execute('SELECT * FROM applications ORDER BY updated_at DESC, id DESC').fetchall()
        result = []
        for row in rows:
            item = {**json.loads(row['details']), **dict(row)}
            item.pop('details')
            item['email'] = item['email'] or ''
            item.setdefault('origin', 'smtp')
            item.setdefault('source', 'Email SMTP' if item['origin'] == 'smtp' else 'Tableur')
            item.setdefault('url', '')
            item['editable'] = bool(item['sent_at'] or item['origin'] in EXTERNAL_ORIGINS)
            item['resultats'] = STATUSES[item['status']]
            item['date_envoi'] = datetime.fromisoformat(item['sent_at']).strftime('%d/%m/%Y') if item['sent_at'] else ''
            item['due'] = bool(item['follow_up_date'] and item['follow_up_date'] <= datetime.now().date().isoformat()
                               and item['status'] in ('sent', 'follow_up'))
            result.append(item)
        return result

    def update(self, record_id, data):
        if not isinstance(data, dict):
            raise ValueError('Les informations de candidature doivent être un objet.')
        with self.connect() as db:
            db.execute('BEGIN IMMEDIATE')
            row = db.execute('SELECT * FROM applications WHERE id=?', (record_id,)).fetchone()
            if not row:
                raise LookupError('Candidature introuvable')
            details = json.loads(row['details'])
            if details.get('origin') in EXTERNAL_ORIGINS:
                entry = external_entry({**details, 'email': row['email'], 'date_envoi': row['sent_at'],
                                        'status': row['status'], 'notes': row['notes'],
                                        'follow_up_date': row['follow_up_date'], **data}, details['origin'])
                if self._duplicate(db, entry, exclude_id=record_id):
                    raise DuplicateApplication('Une autre candidature possède déjà ce contact ou cette entreprise et ce poste.')
                db.execute('''UPDATE applications SET email=?, stage=?, details=?, status=?, sent_at=?,
                    follow_up_date=?, notes=?, updated_at=? WHERE id=?''', (entry['email'] or None, entry['stage'],
                    json.dumps(entry, ensure_ascii=False), entry['status'], entry['date_envoi'], entry['follow_up_date'],
                    entry['notes'], datetime.now().isoformat(timespec='seconds'), record_id))
                return
            if not row['sent_at']:
                raise ValueError('Seules les candidatures envoyées peuvent changer de statut.')
            status = data.get('status', row['status'])
            if status not in ('sent', 'follow_up', 'interview', 'accepted', 'rejected'):
                raise ValueError('Statut invalide')
            for key in ('date_entretien', 'interlocuteur', 'coordonnees'):
                if key in data:
                    details[key] = str(data[key]).strip()[:1000]
            follow_up = str(data.get('follow_up_date', row['follow_up_date'])).strip()
            for value in (follow_up, data.get('date_entretien', '')):
                if value:
                    try:
                        datetime.strptime(value, '%Y-%m-%d')
                    except ValueError:
                        raise ValueError('Utilisez une date au format AAAA-MM-JJ')
            db.execute('UPDATE applications SET status=?, details=?, follow_up_date=?, notes=?, updated_at=? WHERE id=?',
                       (status, json.dumps(details, ensure_ascii=False), follow_up,
                        str(data.get('notes', row['notes']))[:10000], datetime.now().isoformat(timespec='seconds'), record_id))

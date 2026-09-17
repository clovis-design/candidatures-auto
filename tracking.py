"""Suivi local persistant. Aucun identifiant SMTP n'est enregistré."""
import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime, timedelta

STATUSES = {
    'sent': 'En attente de réponse',
    'follow_up': 'À relancer',
    'interview': 'Entretien',
    'accepted': 'Acceptée',
    'rejected': 'Refusée',
    'error': 'Échec d’envoi',
    'sending': 'Envoi en cours',
}


class TrackingStore:
    def __init__(self, path):
        self.path = str(path)
        with self.connect() as db:
            db.execute('''CREATE TABLE IF NOT EXISTS applications (
                id INTEGER PRIMARY KEY, email TEXT NOT NULL, stage TEXT NOT NULL,
                details TEXT NOT NULL, status TEXT NOT NULL, sent_at TEXT DEFAULT '',
                updated_at TEXT NOT NULL, follow_up_date TEXT DEFAULT '',
                notes TEXT DEFAULT '', error TEXT DEFAULT '',
                UNIQUE(email, stage))''')

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
            candidates = db.execute('SELECT * FROM applications WHERE email=?', (entry['email'],)).fetchall()
            old = next((r for r in candidates if r['stage'].casefold() == entry['stage'].casefold()), None)
            if old and (old['sent_at'] or old['status'] == 'sending'):
                return None
            if old:
                db.execute("UPDATE applications SET details=?, status='sending', error='', updated_at=? WHERE id=?",
                           (json.dumps(entry, ensure_ascii=False), now, old['id']))
                return old['id']
            cursor = db.execute('INSERT INTO applications (email, stage, details, status, updated_at, notes) VALUES (?,?,?,?,?,?)',
                                (entry['email'], entry['stage'], json.dumps(entry, ensure_ascii=False), 'sending', now, entry.get('notes', '')))
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
            item['resultats'] = STATUSES[item['status']]
            item['date_envoi'] = datetime.fromisoformat(item['sent_at']).strftime('%d/%m/%Y') if item['sent_at'] else ''
            item['due'] = bool(item['follow_up_date'] and item['follow_up_date'] <= datetime.now().date().isoformat()
                               and item['status'] in ('sent', 'follow_up'))
            result.append(item)
        return result

    def update(self, record_id, data):
        with self.connect() as db:
            row = db.execute('SELECT * FROM applications WHERE id=?', (record_id,)).fetchone()
            if not row:
                raise LookupError('Candidature introuvable')
            if not row['sent_at']:
                raise ValueError('Seules les candidatures envoyées peuvent changer de statut.')
            status = data.get('status', row['status'])
            if status not in ('sent', 'follow_up', 'interview', 'accepted', 'rejected'):
                raise ValueError('Statut invalide')
            details = json.loads(row['details'])
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

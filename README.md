# Candidatures Auto

Un espace local pour préparer des candidatures, envoyer automatiquement **CV + lettre personnalisée par email**, puis suivre les résultats dans l’application et dans un tableur Excel.

Le projet utilise Flask, JavaScript natif, SQLite, openpyxl et fpdf2. La génération de lettres repose sur des modèles éditables, sans service d’IA externe.

## Démarrer

Python 3.10 ou supérieur, avec le module `venv` :

```bash
./start.sh
```

Le script crée un environnement `.venv`, y installe les dépendances et lance l’application sur **http://127.0.0.1:5000**.

Ou, manuellement :

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python app.py
```

Variables facultatives : `PORT=5001`, `HOST=127.0.0.1`, `FLASK_DEBUG=1`, `TRACKING_DB=/chemin/vers/applications.sqlite3`. L’interface est conçue pour un usage local mono-utilisateur, sans authentification. L’envoi nécessite une connexion SMTP autorisée par votre fournisseur.

## Le parcours principal

1. **Profil et CV** : renseignez votre identité, le poste, votre parcours et déposez votre CV PDF/DOC/DOCX.
2. **Lettre** : générez un modèle éditable ou importez votre propre texte. Consultez l’aperçu PDF.
3. **Destinataires** : collez des emails ou importez un tableur. Ajustez l’entreprise, le poste et l’interlocuteur par ligne ; décochez les candidatures à exclure.
4. **Envoi** : choisissez un préréglage SMTP, testez la connexion, puis lancez la campagne. Chaque destinataire reçoit son propre email, son CV et une lettre PDF personnalisée. La progression affiche les résultats réels au fil des envois.
5. **Suivi** : retrouvez les candidatures dans **Mes candidatures**, modifiez leur statut, ajoutez un entretien ou des notes, puis exportez le suivi Excel.

### Personnalisation

Les mêmes variables fonctionnent dans la lettre, l’objet et le corps de l’email :

| Variable | Valeur |
| --- | --- |
| `{ENTREPRISE}` ou `{ENTREPRISE_NOM}` | Entreprise saisie, ou déduite du domaine email |
| `{POSTE}` ou `{STAGE}` | Poste de la ligne importée, sinon poste du profil |
| `{PRENOM}`, `{NOM}` | Identité du candidat |

Exemple d’objet : `Candidature {POSTE} — {ENTREPRISE} — {PRENOM} {NOM}`.

La casse et les espaces à l’intérieur des variables sont acceptés. L’option « Noms d’entreprises personnalisés » donne priorité aux noms saisis dans la liste ou le tableur. La lettre générée conserve `{POSTE}` pour s’adapter à chaque destinataire.

## Tableurs : importer, remplir et mettre à jour

### Alimenter directement le tableau de bord

Dans **Mes candidatures → Importer un tableur**, sélectionnez un fichier `.xlsx`, `.xlsm` ou `.csv` UTF-8. Ses candidatures sont ajoutées à l’historique SQLite et apparaissent immédiatement dans le tableau de bord, sans envoi d’email et sans configuration SMTP.

- Chaque ligne doit renseigner **Entreprise** et **Poste** (ou **Stage ciblé**). L’email est facultatif : les candidatures faites sur un site externe sont acceptées.
- Les dates de candidature, statuts, contacts, notes, entretiens et dates de relance sont conservés. Les colonnes **Source**, **Site** ou **Plateforme** et **Lien de l’offre** / **URL** sont également reconnues.
- Une ligne sans statut prend le statut « En attente de réponse » si elle a une date de candidature, sinon « À envoyer ». Un statut explicite comme « Entretien » peut être importé sans date d’envoi connue : aucune date n’est inventée.
- Statuts reconnus : À envoyer / Brouillon, En attente / Envoyée, À relancer, Entretien / RDV fixé, Acceptée, Refusée et Échec d’envoi. Un statut ou une date non reconnu est signalé avec le numéro de ligne pour correction.
- Un bilan détaille les ajouts, les doublons ignorés et les lignes invalides. Un nouvel import du même fichier ne recrée pas les mêmes candidatures et n’écrase pas les statuts ou notes déjà modifiés dans l’application.
- Le rapprochement utilise **email + poste**, ou **entreprise + poste** si l’email manque d’un côté. Des entreprises différentes sans email peuvent donc avoir le même intitulé de poste.

Exemple de suivi externe :

```csv
Entreprise;Poste;Date candidature;Statut;Site;Lien de l'offre;Notes
Atelier Numérique;Stage Python;15/09/2026;En attente;LinkedIn;https://example.org/offres/42;Candidature déposée en ligne
Studio Ouest;Alternance web;;À envoyer;Indeed;;Dossier à préparer
```

### Ajouter une candidature manuellement

Dans **Mes candidatures → Ajouter manuellement**, renseignez l’entreprise et le poste. Vous pouvez ajouter le site d’origine (LinkedIn, Indeed, site carrière…), le lien de l’offre, un email facultatif, les coordonnées du contact, la date, le statut, une relance et des notes.

Le formulaire propose la date du jour ; vous pouvez la corriger ou la vider si elle est inconnue. Choisir « À envoyer » efface la date d’envoi. **Mettre à jour** permet ensuite de corriger les informations de ces candidatures et de celles importées. Elles restent disponibles après redémarrage et figurent dans les exports Excel, y compris la source et le lien de l’offre.

### Import Excel / CSV pour préparer des envois

- Formats : **`.xlsx`, `.xlsm`** (lecture de la feuille active), **CSV UTF-8**, ou TXT pour une liste simple d’emails.
- Les en-têtes sont reconnus dans les **20 premières lignes**, quelle que soit la position des colonnes. Le modèle historique avec en-têtes en ligne 4 reste compatible.
- Une colonne `Email`, `Mail`, `Adresse email` ou `Coordonnées` est nécessaire. Une seule adresse email par ligne ; un téléphone peut figurer à côté de l’adresse.
- Autres colonnes reconnues : `Entreprise` / `Société`, `Poste` / `Stage ciblé`, `Interlocuteur`, `Type candidature`, `Date envoi`, `Date entretien`, `Résultats` / `Statut`, `Date relance`, `Notes`.
- Les lignes invalides et doublons email/poste sont signalés avec leur numéro. Deux postes distincts pour la même adresse restent deux candidatures distinctes.
- Limites : 10 Mo par requête, 10 000 lignes et 200 colonnes pour Excel, 500 candidatures par campagne.

Exemple CSV :

```csv
Entreprise;Email;Poste;Interlocuteur
Atelier Numérique;rh@atelier.example;Stage développement Python;Camille
Studio Ouest;contact@studio.example;Alternance développement web;Alex
```

### Trois usages du tableur

1. **Télécharger un modèle** : classeur vierge, en-têtes en ligne 4, données dès la ligne 5.
2. **Exporter cette liste en Excel** : export des destinataires sélectionnés, avec leurs informations importées. Une candidature non envoyée a une date d’envoi vide et le statut « À envoyer ».
3. **Exporter le suivi Excel**, dans Mes candidatures : export de l’historique persistant, avec dates d’envoi réelles, statuts actuels, interlocuteurs, entretiens, relances et notes. Les envois échoués restent sans date d’envoi.

Le bouton **Compléter mon tableur** permet de charger votre propre `.xlsx` et de télécharger une copie mise à jour. Le rapprochement utilise l’email et le poste, sans dépendre de l’ordre des lignes. Si le poste est absent, une correspondance email unique est requise ; sans email, une correspondance entreprise/poste unique est requise. Les cas ambigus sont laissés inchangés. Les colonnes de suivi manquantes sont ajoutées ; les autres cellules, formules, styles usuels et feuilles sont conservés par openpyxl. Cette fonction ne modifie pas votre fichier original sur disque et n’est pas une synchronisation Google Sheets.

## Historique, relances et doublons

- L’historique est enregistré dans **`data/applications.sqlite3`** et reste disponible après redémarrage. Sauvegardez ce fichier pour conserver votre suivi.
- La base existante est migrée automatiquement pour permettre les candidatures sans email, en conservant les identifiants et les données déjà enregistrées.
- Statuts modifiables après un envoi : en attente, à relancer, entretien, acceptée, refusée.
- Recherche par entreprise, adresse, poste, interlocuteur ou notes, et filtres par statut.
- Un rappel est fixé à 3, 7, 14 ou 30 jours après l’envoi, selon votre choix. Les candidatures en attente arrivées à échéance apparaissent dans **À relancer**. Ce sont des rappels dans le tableau de bord ; aucun email de relance n’est envoyé automatiquement.
- Un même couple **email / poste** déjà envoyé est ignoré lors d’une nouvelle campagne, même après redémarrage. Les doublons d’une même liste sont supprimés. Une réservation SQLite atomique protège aussi les envois concurrents.
- Une date d’envoi présente dans un tableur empêche un nouvel envoi de la ligne. Vous pouvez vider une date provenant d’un ancien brouillon dans l’aperçu éditable ; l’historique de l’application continue de protéger les envois réellement effectués.
- Un échec peut être chargé avec **Reprendre** : le destinataire est remis dans le formulaire pour un nouvel essai. Une campagne interrompue dont l’acceptation SMTP est incertaine reste réservée ; vérifiez l’envoi avant toute correction manuelle du suivi.
- Pour les envois de l’application, « Envoyée » signifie que le serveur SMTP a accepté le message, pas que le destinataire l’a lu ni que la livraison finale est garantie. Pour les candidatures externes, le statut repose sur les informations saisies ou importées ; leur dépôt n’est pas vérifié sur le site externe. Le compteur des candidatures envoyées inclut ces candidatures déclarées déposées, même si leur date exacte est inconnue.

Le brouillon (profil, texte, destinataires, réglages) est sauvegardé dans le navigateur. Le **mot de passe SMTP et le fichier CV ne sont pas sauvegardés** : renseignez-les à nouveau après rechargement. Les CV temporaires et les PDF créés pour l’envoi sont supprimés après traitement ; les PDF d’aperçu restent dans `generated/`.

## Configuration SMTP

Les préréglages remplissent l’hôte, le port et la sécurité. **Tester la connexion SMTP** vérifie la connexion et l’authentification sans envoyer de message.

| Fournisseur | Hôte | Port | Sécurité |
| --- | --- | --- | --- |
| Gmail | smtp.gmail.com | 587 | STARTTLS |
| Outlook / Microsoft 365 | smtp.office365.com | 587 | STARTTLS |
| OVH | ssl0.ovh.net | 465 | SSL |
| Test local | localhost | 1025 | Aucune |

Pour Gmail, activez la validation en deux étapes et créez un [mot de passe d’application](https://myaccount.google.com/apppasswords). Les comptes Microsoft doivent autoriser SMTP AUTH ; certaines organisations imposent OAuth, qui n’est pas pris en charge ici. Les identifiants sont transmis au serveur SMTP au moment du test ou de l’envoi, jamais stockés dans SQLite.

Le délai entre deux envois est réglable entre 0 et 60 secondes ; sa valeur initiale est de 1,5 seconde. Gardez la page ouverte pendant la campagne. L’envoi s’exécute dans la requête Flask, avec interrogation périodique de sa progression ; il n’y a pas de file de tâches planifiée. Utilisez une seule instance du serveur pour le suivi live des campagnes.

## Structure

```text
app.py                  Routes Flask, génération PDF, envoi SMTP
tracking.py             Historique SQLite, réservations et mises à jour
spreadsheets.py          Import Excel/CSV et mise à jour de classeurs
templates/index.html    Interface française
static/css/style.css    Styles adaptatifs ordinateur/mobile
static/js/app.js        Préparation, import et campagne
static/js/workspace.js  Tableau de bord, brouillon et suivi
tests/test_app.py       Tests d’intégration avec SMTP simulé
data/                   Base locale (ignorée par Git)
uploads/                CV temporaires
generated/              Aperçus PDF
```

## Tests

```bash
.venv/bin/python -m unittest discover -s tests -v
node --check static/js/app.js
node --check static/js/workspace.js
```

Les tests utilisent une base temporaire et un SMTP simulé : aucun email réel n’est envoyé. Ils vérifient la personnalisation et les pièces jointes, les erreurs et reprises, les doublons, les réservations concurrentes, les imports, la conservation des données Excel, les dates réelles et les modifications du suivi.

## API principale

| Méthode | Route | Usage |
| --- | --- | --- |
| POST | `/api/generate-letter` | Générer une lettre et son aperçu PDF |
| POST | `/api/preview-letter` | Prévisualiser un texte personnalisé |
| POST | `/api/parse-emails` | Analyser une liste de destinataires |
| POST | `/api/parse-excel` | Importer Excel/CSV (`file` multipart) |
| GET | `/api/excel-template` | Télécharger le modèle vierge |
| POST | `/api/generate-excel` | Exporter une liste (`entries` ou `emails_text`) |
| POST | `/api/test-smtp` | Tester la configuration SMTP |
| POST | `/api/send` | Envoyer une campagne multipart, dont `cv` |
| GET | `/api/campaigns/<id>` | Consulter la progression de la campagne |
| GET | `/api/applications` | Historique, statistiques et statuts |
| POST | `/api/applications` | Ajouter manuellement une candidature (entreprise et stage requis, email facultatif) |
| POST | `/api/applications/import` | Importer un tableur directement dans l’historique (`file` multipart) |
| PATCH | `/api/applications/<id>` | Modifier statut, dates, interlocuteur et notes |
| GET | `/api/applications/export` | Exporter l’historique Excel |
| POST | `/api/update-excel` | Compléter un `.xlsx` existant (`file` multipart) |

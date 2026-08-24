## frontend

### Rôle

Ce répertoire contient une interface web statique simple pour piloter le backend Flask de analyse_ao.
L'interface est volontairement limitée au **mode dropzone**.

### Fichiers
- index.html : structure HTML de l'interface dropzone, rapport et téléchargement.
- app.js : logique JavaScript d'upload, validation Excel, détection des lots, appels API et renommage avant téléchargement.
- styles.css : styles et mise en forme.
- Picto_Betrec02.svg : logo affiché dans l'interface.

### Fonctionnalités
- **Zone Référence BETREC (DCE Excel)** : dépôt d'un ou plusieurs classeurs Excel.
- **Zone Retours entreprises (ACT Excel)** : dépôt d'un ou plusieurs classeurs Excel.
- **Validation côté interface** : seuls les fichiers .xlsx, .xlsm et .xls sont acceptés pour le traitement.
- **Gestion multi-lots** : le lot est détecté dans le nom et peut être saisi à la volée lorsqu'il est absent, par exemple `lot 01`, `lot_02`, `L03`.
- **Restructuration** : envoi des fichiers déposés au backend via `POST /api/run/files`.
- **Téléchargement du XLSX résultat** : le bouton devient actif lorsque le backend fournit un `download_url` ou un `workbook_path`.
- **Renommage avant téléchargement** : une fenêtre demande le nom du fichier avant de lancer le téléchargement.

### API utilisée
- GET /api/health
- GET /api/config
- GET /api/last-run
- GET /api/download?path=...
- POST /api/run/files

Le contrat attendu dans la réponse data reste : status, workbook_path, download_url, warnings, errors.
Le frontend envoie aussi `dropped_file_lots` pour exposer les numéros de lots détectés, et conserve `act_enterprise_names` en compatibilité legacy si le backend l'attend encore.

### Conditions d'exécution
- Le backend Flask doit être démarré localement sur 127.0.0.1:8008.
- Il n'y a pas de bundling ni de build : ouvrir index.html dans un navigateur suffit si le backend est disponible.

### Points sensibles
- Après remplacement de app.js, le navigateur peut servir une version en cache : faire Ctrl + F5 si le comportement ne change pas.
- Les messages affichés sont destinés à l'utilisateur métier : français clair, sans jargon technique.
- Les numéros de lot sont lus dans les noms de fichiers. Le nom de l'entreprise reste obligatoire pour chaque ACT.
- Si le backend dépend réellement du nom d'entreprise métier, il faudra aligner le backend sur `dropped_file_lots`.

### À ne pas toucher
- Ne pas changer l'hôte/port du backend (127.0.0.1:8008) sans aligner `Lancer_AnalyseAO.bat` et `analyse_ao/api/server.py`.
- Ne pas réintroduire le mode numéro d'affaire dans l'interface frontend.
- Conserver la saisie du nom d'entreprise pour chaque ACT.

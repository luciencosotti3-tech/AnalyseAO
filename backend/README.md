# Backend Flask propre

- Adresse : `http://127.0.0.1:8008`
- Sante : `GET /api/health`
- Restructuration : `POST /api/run/files` (champs `dce_files`, `act_files`, `dropped_file_lots`, `act_enterprise_names`)
- Analyse d'offres : `POST /api/run/offer-analysis` (champ `workbook`, classeur `.xlsx` restructuré en entrée)
- Telechargement : `GET /api/download/<filename>` (nom proposé au navigateur via `?download_name=...`)
- Frontend servi directement par Flask : `GET /` et `GET /<path:filename>` renvoient les fichiers de `frontend/`

Le depot valide les fichiers Excel (`.xlsx`/`.xlsm` uniquement), les lots et les noms d'entreprise ACT.
`/api/run/files` construit la matrice DCE/ACT (`case_matrix.py`), génère le classeur multi-entreprises (`multi_company.py`) et renvoie `workbook_path` + `download_url`.
`/api/run/offer-analysis` prend en entrée le classeur restructuré et produit un rapport d'analyse (`offer_analysis.py`).

## Moteur Excel V1

Le module `excel_engine.py` réalise les premières opérations métier indépendamment de Flask :

- identification de plusieurs feuilles métier ;
- exclusion des pages sans données métier ;
- distinction BASE / OPTION-PSE / TRANCHE ;
- détection des marqueurs techniques `ART`, `CH3`, `STOT`, `TOTHT`, `TVA`, `TOTTTC` ;
- détection des blocs Quantité / PU / Montant répétés ;
- lecture facultative d'une colonne TVA par article ;
- conservation du détail DCE et sélection du cumul ;
- quantité entreprise facultative ;
- alignement exact par identifiant, puis référence, puis désignation + unité ;
- contrôles initiaux de cohérence.

Test manuel :

```powershell
python backend\inspect_offer.py --dce "DCE.xlsx" --act "ACT.xlsx" --entreprise "SMAC" --lot 05 --output "analyse_excel_v1.json"
```


## Sortie Excel V2

La commande `inspect_offer.py` produit désormais un classeur `.xlsx` avec :

- `BASE` ;
- `OPTIONS`, créée même si le DCE ne contient aucune option ;
- une feuille complémentaire par option ACT lorsque plusieurs feuilles d'option sont présentes.

Les options ACT sans équivalent DCE sont conservées avec le statut `OPTION AJOUTEE ACT`.
Les options DCE non chiffrées par l'entreprise restent visibles avec le statut `ABSENT ACT`.


## Alignement V3

L'alignement travaille en deux niveaux :

1. appariement des feuilles métier par périmètre, recouvrement des identifiants, références et nom ;
2. appariement des articles à l'intérieur de la paire de feuilles.

Priorité article : identifiant exact, référence, désignation normalisée, unité. Les rapprochements faibles ou ambigus sont signalés sans être masqués. Une feuille d'option uniquement présente dans l'ACT reste une option ACT autonome.


## Comparaison côte à côte V4

`compare_side_by_side.py` copie les tableaux visibles du DCE et de l'ACT dans une même feuille, conserve leurs cellules, styles, largeurs, hauteurs, formules et fusions, puis insère les lignes nécessaires pour aligner les articles sans aplatir les structures.

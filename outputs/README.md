# Conversion du planning vers Zoho Projects

Le script `convertir_zoho.py` convertit localement le deuxième onglet du planning au format du modèle d'import Zoho. Il ne se connecte à aucun service et ne crée aucune tâche dans Zoho.

## Utilisation

Depuis le dossier `outputs` :

Les colonnes Zoho et les champs personnalisés sont définis directement dans le script. Aucun fichier modèle n'est nécessaire. `XLSX-sample-taskimport.xlsx` est conservé comme référence du format standard Zoho.

```bash
python3 -m pip install -r requirements.txt
python3 convertir_zoho.py /chemin/planning.xlsx /chemin/import_zoho.xlsx
```

Le script produit un fichier Excel et un rapport texte. Les fichiers sources ne sont pas modifiés. Les classeurs et rapports générés dans `outputs` sont exclus de Git.

## Options

Exécuter `python3 convertir_zoho.py --help` pour toutes les options.

- `--sheet` : nom de l'onglet source (deuxième onglet par défaut).
- `--list-name` : nom de la liste cible.
- `--list-id` : identifiant de la liste existante, écrit dans `TASK LIST NAME` selon la documentation Zoho.
- `--owners` : fichier JSON associant les IDs source aux utilisateurs Zoho, par exemple `{"L01": "utilisateur"}`.
- `--status` : statut Zoho confirmé à appliquer.
- `--dependencies` : fichier JSON remplaçant toutes les dépendances des tâches spécifiées, par exemple `{"L18": ["L17 SS", "L05 FS"]}`.

## Règles de conversion

Les noms, identifiants et dates sont conservés. Les dépendances sont fin-début (FS), sauf l'exception L18/L17 début-début (SS) explicitement décrite dans la source. Les pilotes sont placés dans Responsible, les critères dans Prerequisites et les références / écarts dans la description. Charge et avancement restent vides. Les conflits de dates détectés sont signalés dans le rapport sans replanification.

Le tableau source doit commencer par une ligne `ID` / `Action` et conserver les intitulés de colonnes attendus. Une ligne sans ID marque sa fin. Le script vérifie les dates, les identifiants, les références et les cycles de dépendances.

## Import manuel

La colonne `Prerequisites` reprend intégralement `Critère d’acceptation / prérequis`. L'associer au champ personnalisé multiligne `Prerequisites` (nom API `prerequisites`) dans Zoho.

La colonne supplémentaire `Responsible` reprend intégralement `Pilote proposé / appuis`. L'associer au champ personnalisé RACI `Responsible` dans l'étape de mapping des champs personnalisés Zoho. Les mentions de plusieurs personnes et équipes restent telles quelles, sans attribution automatique à des comptes utilisateurs. Le champ `Accountable` n'est pas alimenté, car la source ne le distingue pas.

Dans Zoho, associer `TASK NAME`, `TASK ID`, `PREDECESSORS`, `TASK DESCRIPTION`, `TASK LIST NAME`, `START DATE` et `DUE DATE` aux champs correspondants. Le format des dates est `dd/MM/yyyy`. Les propriétaires restent non attribués en l'absence de mapping explicite.

Documentation : https://help.zoho.com/portal/en/kb/projects/settings-in-zoho-projects/import-data/articles/import-tasks-from-files

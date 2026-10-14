# zoho-projects-task-converter_PLF

Convertisseur Python local d'un **format de planning Excel précis** vers un fichier d'import de tâches Zoho Projects.

Le script lit le tableau de détail du planning et produit un classeur `.xlsx` avec les colonnes standard Zoho et deux champs personnalisés : `Responsible` et `Prerequisites`. Il produit également un rapport texte signalant les conflits de dates détectés.

**Aucun accès à Zoho ni identifiant de connexion n'est nécessaire.** Le script ne fait aucun appel réseau. L'import du résultat se fait manuellement dans Zoho Projects.

## Périmètre

Ce convertisseur attend les intitulés et la structure décrits ci-dessous. Il ne convertit pas automatiquement n'importe quel planning ou diagramme de Gantt. Les données sont extraites du tableau de détail, pas des barres graphiques.

Le fichier `outputs/XLSX-sample-taskimport.xlsx` est un exemple de référence du format Zoho. Le script ne le lit pas : les colonnes de sortie sont définies directement dans `ZOHO_COLUMNS`.

## Installation

Python 3 et `openpyxl` sont nécessaires. Depuis la racine du dépôt :

```bash
git clone https://github.com/BigMentor-FR/zoho-projects-task-converter_PLF.git
cd zoho-projects-task-converter_PLF
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r outputs/requirements.txt
```

Sous Windows, activer l'environnement avec `.venv\Scripts\Activate.ps1` dans PowerShell.

## Conversion

```bash
python3 outputs/convertir_zoho.py /chemin/planning.xlsx /chemin/import_zoho.xlsx \
  --list-name "Liste cible"
```

Préciser le nom de liste cible pour chaque utilisation. Par défaut, le script lit **le deuxième onglet** du classeur. Pour en sélectionner un autre :

```bash
python3 outputs/convertir_zoho.py /chemin/planning.xlsx /chemin/import_zoho.xlsx \
  --sheet "Détail" --list-name "Liste cible"
```

Deux fichiers sont produits : `import_zoho.xlsx` et `import_zoho.rapport.txt`. Le planning source reste intact. La sortie ne peut pas avoir le même chemin que la source. Si les fichiers de sortie existent déjà, ils sont remplacés.

## Format du planning source

La ligne d'en-tête peut se trouver après des lignes de titre. Elle doit commencer par `ID` en colonne A et `Action` en colonne B. Les autres colonnes requises sont repérées par leur intitulé exact et peuvent changer de position.

| Colonne source requise | Contenu attendu | Destination |
| --- | --- | --- |
| `ID` | Identifiant unique : lettres suivies de chiffres, par exemple `T01` | `TASK ID` |
| `Action` | Nom non vide, 255 caractères maximum | `TASK NAME` |
| `Début proposé` | Date Excel réelle | `START DATE` |
| `Fin proposée` | Date Excel réelle, au moins égale au début | `DUE DATE` |
| `Pilote proposé / appuis` | Personnes ou équipes, conservées telles quelles | `Responsible` |
| `Prédécesseurs` | IDs séparés par des virgules, ou vide / `—` / `-` | `PREDECESSORS` |
| `Statut à renseigner` | Colonne attendue, sans reprise automatique de sa valeur | Aucun, sauf statut fourni via `--status` |
| `Critère d’acceptation / prérequis` | Texte libre | `Prerequisites` |
| `Source / écart à l’ébauche` | Références et contexte utiles | `TASK DESCRIPTION` |

Les tâches doivent être contiguës : **la première ligne sans ID termine la lecture**. Les tableaux annexes placés après cette ligne ne sont pas importés. Les colonnes supplémentaires, dont les durées calendaires, ne sont pas converties en charge.

Le classeur est lu avec les valeurs enregistrées des cellules. Si des dates sont calculées par formule, enregistrer préalablement le classeur dans un tableur qui recalcule ces formules.

## Correspondance avec Zoho

Dans l'assistant d'import, associer les colonnes suivantes aux champs correspondants :

| Colonne générée | Champ Zoho |
| --- | --- |
| `TASK NAME` | Nom de tâche |
| `TASK ID` | ID de tâche, utilisé pour les liens entre les tâches importées |
| `PREDECESSORS` | Prédécesseurs |
| `TASK DESCRIPTION` | Description de tâche |
| `TASK LIST NAME` | Nom de la liste des tâches |
| `START DATE` | Date de début |
| `DUE DATE` | Date d'échéance |
| `Responsible` | Champ personnalisé texte RACI `Responsible` (API : `responsible`) |
| `Prerequisites` | Champ personnalisé multiligne `Prerequisites` (API : `prerequisites`) |

Les champs personnalisés doivent exister dans la disposition de tâches utilisée. L'association automatique des colonnes n'est pas garantie : vérifier le mapping avant de continuer. Choisir le format de date `dd/MM/yyyy`.

Les propriétaires restent vides sans mapping explicite. Le champ RACI `Responsible` est du texte : il n'attribue pas automatiquement la tâche à un compte utilisateur. `Accountable`, la charge, la priorité et l'avancement ne sont pas déduits du planning. La description contient uniquement la source / l'écart à l'ébauche, sans recopier les champs dédiés.

### Liste existante

Pour cibler une liste existante avec son identifiant :

```bash
python3 outputs/convertir_zoho.py /chemin/planning.xlsx /chemin/import_zoho.xlsx \
  --list-id 123456789
```

Le script écrit cet identifiant comme texte dans `TASK LIST NAME`, conformément à la procédure Zoho pour l'import dans une liste existante. `TASK LIST ID` reste vide. Sans identifiant, le nom choisi via `--list-name` est utilisé ; vérifier la liste obtenue dans l'assistant.

## Dépendances et dates

Un prédécesseur source `T01` devient `T01 FS` : fin de T01 avant début du successeur. Plusieurs IDs produisent plusieurs liens, par exemple `T01 FS, T02 FS`.

Le script conserve une règle spécifique au format d'origine : pour `L18`, il ajoute `L17 SS` lorsque la référence à « début de L17 » figure dans la colonne source / écart et que L17 existe. `SS` signifie début-début. Pour adapter les liens, utiliser `--dependencies`.

Les dates ne sont pas recalculées. Un début le même jour que la fin d'un prédécesseur FS, ou avant cette fin, est signalé dans le rapport. Le même jour peut nécessiter une vérification des heures et du calendrier Zoho. Les liaisons SS dont le successeur commence avant le prédécesseur sont aussi signalées. Les contraintes de dates FF et SF ne sont pas vérifiées par ce rapport.

Les IDs dupliqués, prédécesseurs inconnus, autodépendances, cycles, noms invalides et dates invalides arrêtent la conversion.

## Options avancées

```bash
python3 outputs/convertir_zoho.py --help
```

| Option | Fonction |
| --- | --- |
| `--sheet NOM` | Sélectionner l'onglet source |
| `--list-name NOM` | Définir le nom de liste cible |
| `--list-id ID` | Cibler une liste existante ; prioritaire sur le nom |
| `--owners FICHIER.json` | Fournir des propriétaires par ID de tâche |
| `--status STATUT` | Appliquer un statut Zoho confirmé à toutes les tâches |
| `--dependencies FICHIER.json` | Remplacer les dépendances des tâches indiquées |

Exemple de fichier de propriétaires, avec valeurs fictives :

```json
{
  "T01": "Utilisateur exemple",
  "T02": "Autre utilisateur"
}
```

Ces valeurs sont écrites dans `TASK OWNER` et doivent être associées aux utilisateurs du projet dans Zoho.

Exemple de fichier de dépendances :

```json
{
  "T03": ["T01 FS", "T02 SS"],
  "T04": []
}
```

Chaque entrée remplace **toutes** les dépendances de la tâche concernée, y compris la règle spéciale éventuelle. Une liste vide supprime ses liens. Les types acceptés sont `FS`, `SS`, `FF` et `SF`. Les décalages tels que `+ 2d` ne sont pas pris en charge par cette option.

## Fichiers du dépôt

- `outputs/convertir_zoho.py` : script de conversion.
- `outputs/requirements.txt` : dépendance Python.
- `outputs/XLSX-sample-taskimport.xlsx` : référence du format standard, facultative.

Les plannings métier et résultats de conversion ne font pas partie du dépôt. Pour éviter de les ajouter à Git, conserver ses fichiers de travail hors du dépôt ou dans `work/`, et les résultats dans `outputs/` : ces emplacements sont ignorés, sauf les fichiers explicitement versionnés. Ne pas publier ses fichiers de mapping s'ils contiennent des données internes.

## Documentation Zoho

[Import de tâches et syntaxe des dépendances](https://help.zoho.com/portal/en/kb/projects/settings-in-zoho-projects/import-data/articles/import-tasks-from-files).

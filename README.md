# Convertisseur CSV pour Affacturage

Application de conversion de fichiers Excel vers des fichiers CSV formatés pour l'affacturage. Elle permet à partir d'un journal de facturation, de ressortir des fichiers distincts et construits correctement pour un traitement par FactoFrance.

## 📋 Description

Cette application permet de convertir des fichiers Excel contenant des factures en fichiers CSV formatés selon les spécifications de l'affacturage FactoFrance. Elle génère automatiquement :
- Fichiers de **Balance** (FBA) : liste des factures
- Fichiers de **Tiers** (TIE) : informations clients
- Séparation automatique **France (1A)** / **Étranger (1B)**

## 🗂️ Structure du projet

```
CSV-MAM/
├── interface.py              # Interface graphique principale (tkinter)
├── traitement.py             # Fonctions de traitement des données
├── requirements.txt          # Dépendances Python
├── burographic.ico           # Icône de l'application
├── datas/                    # Données de référence
│   ├── clients_siret.csv     # Base clients avec SIRET et adresses
│   └── codes_pays.csv        # Correspondance pays → code ISO
└── dist/                     # Exécutable compilé
    └── Convertisseur-CSV.exe
```

## 🔧 Fonctionnement du code

### 1. **interface.py** - Interface graphique

**Rôle :** Gère l'interface utilisateur et orchestre le flux de traitement.

**Flux d'exécution :**
```python
1. Sélection du fichier Excel source
2. Validation du fichier
3. Lecture et chargement dans un DataFrame pandas
4. Génération du DataFrame Balance
5. Séparation clients FR/étranger
6. Génération des DataFrames Tiers
7. Export des fichiers CSV
```

**Fonctions principales :**
- `choisir_fichier()` : Ouvre un dialogue de sélection de fichier
- `lancer_conversion()` : Lance le processus complet de conversion

### 2. **traitement.py** - Logique métier

#### Fonctions principales

**`convertir_fichier(chemin_fichier, sheet_name=0)`**
- Lit un fichier Excel et retourne un DataFrame pandas
- Gère les erreurs d'encodage et de format
- Retourne : `(succès: bool, résultat: DataFrame|str)`

**`generate_balance_file(df_source)`**
- Génère le fichier Balance à partir des données sources
- Ajoute lignes de début (000000) et fin (999999)
- Mappe les codes règlement : T→TRT, C→CHE, V→VIR, A→AVO
- Calcule montants : positifs (VIR/CHE/TRT), négatifs (AVO)
- Arrondit montants à 2 décimales
- Formate dates au format DD/MM/YYYY

**Structure Balance :**
```
Code vendeur cédant | Date fichier | Code client | N° pièce | Date pièce | Devise | Montant | Date échéance | Type | Mode règlement | N° commande
```

**`separer_clients_par_pays(df_balance, df_clients)`**
- Sépare un DataFrame Balance en deux : clients FR et clients étrangers
- Compare le champ `Pays` du fichier clients_siret.csv
- Retourne : `(df_balance_fr, df_balance_etranger)`

**`generate_tiers_file(df_balance)`**
- Génère le fichier Tiers à partir d'un DataFrame Balance
- Déduplique automatiquement les clients
- Charge les données depuis `clients_siret.csv` et `codes_pays.csv`
- Tronque les champs selon les longueurs max :
  - SIRET : 14 caractères
  - Raison sociale : 40 caractères
  - Voie : 40 caractères
  - Code postal : 6 caractères
  - Ville : 34 caractères
- Gère les valeurs NaN (convertit en chaîne vide)
- Retourne : `(df_tiers, clients_non_identifies)`

**Structure Tiers :**
```
Code vendeur cédant | Code client | SIRET | Sigle | Raison sociale | N° voie | Complément | CP | Ville | Code pays ISO
```

**`export_dataframe_to_csv(df_source, type, suffixe='1A', dossier_destination=None)`**
- Exporte un DataFrame en fichier CSV
- Génère nom de fichier : `{TYPE}SS{CEDANT}{SUFFIXE}.{JOUR_ANNEE}`
  - Exemple : `FBASS0123451A.346` (346e jour de l'année)
- Format : séparateur `;`, encodage `utf-8-sig`, sans en-têtes
- Nombres : format `%.2f` (2 décimales obligatoires)

**`get_resource_path(relative_path)`**
- Résout les chemins de fichiers pour PyInstaller
- En développement : chemin relatif normal
- En .exe : utilise `sys._MEIPASS` (dossier temporaire)

### 3. Fichiers de données

**datas/clients_siret.csv**
```csv
Code;Nom;Voie;Complement;CP;Ville;Pays;SIRET;Raison sociale
12050;CAZENAVE;PLACE GERE BELESTEN;AEROPOLE;64121;SERRES;FRANCE;31095537200027;CAZENAVE
```

**datas/codes_pays.csv**
```csv
Pays;ISO
FRANCE;FR
ESPAGNE;ES
ITALIE;IT
```

## 🚀 Installation et utilisation

### Prérequis
- Python 3.8+
- pip

### Installation des dépendances
```bash
pip install -r requirements.txt
```

### Lancement en développement
```bash
python interface.py
```

### Utilisation
1. Cliquez sur "📁 Parcourir..." pour sélectionner un fichier Excel
2. Cliquez sur "Lancer la conversion"
3. Les fichiers CSV sont générés dans le répertoire du projet :
   - `FBASS0123451A.001` : Balance clients français
   - `TIESS0123451A.001` : Tiers clients français
   - `FBASS0123451B.001` : Balance clients étrangers (si présents)
   - `TIESS0123451B.001` : Tiers clients étrangers (si présents)

## 📦 Compilation en exécutable

### Avec PyInstaller
```bash
# Installation
pip install pyinstaller

# Compilation
pyinstaller --onefile --windowed --name "Convertisseur-CSV" --add-data "datas;datas" --add-data "burographic.ico;." --icon="burographic.ico" interface.py
```

L'exécutable sera généré dans `dist/Convertisseur-CSV.exe`

### Options de compilation
- `--onefile` : Un seul fichier .exe
- `--windowed` : Sans console (interface graphique uniquement)
- `--add-data "datas;datas"` : Inclut le dossier des données
- `--icon="burographic.ico"` : Icône de l'application

## 🔍 Détails techniques

### Format des fichiers CSV de sortie

**Séparateur :** Point-virgule (`;`)  
**Encodage :** UTF-8 avec BOM (`utf-8-sig`)  
**En-têtes :** Aucun (fichiers sans ligne d'en-tête)  
**Nombres :** Format `%.2f` (ex: `1234.50`)  
**Dates :** Format `DD/MM/YYYY`

### Logique de séparation FR/Étranger

```python
if pays.upper() == 'FRANCE':
    → Fichiers 1A (français)
else:
    → Fichiers 1B (étrangers)
```

### Gestion des erreurs

- **Fichiers manquants** : Messages d'erreur explicites
- **Clients non identifiés** : Warning avec liste des codes manquants
- **Valeurs NaN** : Converties en chaînes vides
- **Erreurs d'encodage** : Gérées automatiquement avec `utf-8-sig`

## 🔒 Conventions de nommage

**Fichiers Balance :**
```
FBA + SS + {CODE_CEDANT} + {1A|1B} + . + {JOUR_ANNEE}
Exemple : FBASS0123451A.346
```

**Fichiers Tiers :**
```
TIE + SS + {CODE_CEDANT} + {1A|1B} + . + {JOUR_ANNEE}
Exemple : TIESS0123451A.346
```

**Lignes spéciales :**
- `000000` : Ligne de début (DEB)
- `999999` : Ligne de fin (FIN)

## 📚 Dépendances

- **pandas** : Manipulation de données tabulaires
- **openpyxl** : Lecture de fichiers Excel (.xlsx)
- **tkinter** : Interface graphique (inclus avec Python)
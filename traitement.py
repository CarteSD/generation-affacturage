"""
Module de traitement des fichiers
Contient les fonctions de conversion et de traitement des données
"""

import os
import sys
import shutil
from pathlib import Path
from re import match
from unittest import case


def get_resource_path(relative_path):
    """
    Obtient le chemin absolu d'une ressource, compatible avec PyInstaller.
    """
    try:
        # PyInstaller crée un dossier temp et stocke le chemin dans _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)


def get_data_file_path(filename):
    """
    Obtient le chemin d'un fichier de données modifiable par l'utilisateur.
    Les fichiers sont stockés dans Documents/CSV-MAM/config.
    Si le fichier n'existe pas, il est copié depuis le bundle.
    
    Args:
        filename (str): Nom du fichier (ex: 'clients_siret.csv')
    
    Returns:
        str: Chemin complet du fichier de données
    """
    # Dossier de configuration dans Documents
    config_dir = os.path.join(os.path.expanduser("~"), "Documents", "CSV-MAM", "config")
    
    # Créer le dossier s'il n'existe pas
    os.makedirs(config_dir, exist_ok=True)
    
    # Chemin du fichier dans le dossier config
    data_file_path = os.path.join(config_dir, filename)
    
    # Si le fichier n'existe pas, le copier depuis le bundle
    if not os.path.exists(data_file_path):
        source_path = get_resource_path(os.path.join('datas', filename))
        if os.path.exists(source_path):
            shutil.copy2(source_path, data_file_path)
    
    return data_file_path

def convertir_fichier(chemin_fichier, sheet_name=0):
    """
    Lit un fichier Excel et retourne un DataFrame pandas.

    Args:
        chemin_fichier (str): Chemin du fichier source.
        sheet_name (int|str): Index ou nom de la feuille à lire (par défaut 0).

    Returns:
        tuple: (succès: bool, résultat: DataFrame|str)
            - Si succès=True, résultat est un DataFrame pandas
            - Si succès=False, résultat est un message d'erreur
    """
    try:
        try:
            import pandas as pd
        except Exception as e:
            msg = "Le package 'pandas' (et 'openpyxl') n'est pas installé. Installez-le avec: pip install pandas openpyxl"
            return False, msg

        if not os.path.exists(chemin_fichier):
            return False, "Le fichier n'existe pas"

        extension = Path(chemin_fichier).suffix.lower()
        if extension not in ['.xlsx', '.xls', '.xlsm']:
            return False, "Le fichier doit être un fichier Excel (.xlsx, .xls ou .xlsm)"

        # Lire le fichier Excel
        df = pd.read_excel(chemin_fichier, sheet_name=sheet_name, engine="openpyxl")

        return True, df
    except Exception as e:
        msg = f"Erreur lors de la lecture : {str(e)}"
        return False, msg


def valider_fichier(chemin_fichier):
    """
    Valide si le fichier peut être traité
    
    Args:
        chemin_fichier (str): Chemin du fichier à valider
    
    Returns:
        tuple: (valide, message)
    """
    if not os.path.exists(chemin_fichier):
        return False, "Le fichier n'existe pas"
    
    extension = Path(chemin_fichier).suffix.lower()
    if extension not in ['.xlsx', '.xls', '.xlsm', '.csv']:
        return False, f"Format non supporté : {extension}"
    
    taille = os.path.getsize(chemin_fichier)
    if taille == 0:
        return False, "Le fichier est vide"
    
    return True, "Fichier valide"


def separer_clients_par_pays(df_balance, df_clients):
    """
    Sépare un DataFrame Balance en clients français et étrangers.
    
    Args:
        df_balance (DataFrame): DataFrame Balance.
        df_clients (DataFrame): DataFrame des informations clients.
    
    Returns:
        tuple: (df_balance_fr, df_balance_etranger)
    """
    import pandas as pd
    
    # Récupérer les lignes de début/fin
    ligne_debut = df_balance[df_balance['Code client'] == '000000']
    ligne_fin = df_balance[df_balance['Code client'] == '999999']
    
    # Récupérer les lignes de données (sans début/fin)
    df_data = df_balance[(df_balance['Code client'] != '000000') & (df_balance['Code client'] != '999999')]
    
    # Séparer français et étranger
    lignes_fr = []
    lignes_etranger = []
    
    for _, row in df_data.iterrows():
        code_client = row['Code client']
        
        # Vérifier si le client existe dans df_clients
        if code_client in df_clients['Code'].values:
            client_info = df_clients[df_clients['Code'] == code_client].iloc[0]
            pays = str(client_info.get('Pays', 'FRANCE')).upper()
            
            if pays == 'FRANCE':
                lignes_fr.append(row)
            else:
                lignes_etranger.append(row)
        else:
            # Par défaut, considérer comme français si non trouvé
            lignes_fr.append(row)
    
    # Créer les DataFrames
    if lignes_fr:
        df_fr = pd.concat([ligne_debut, pd.DataFrame(lignes_fr), ligne_fin], ignore_index=True)
    else:
        df_fr = pd.DataFrame(columns=df_balance.columns)
    
    if lignes_etranger:
        df_etranger = pd.concat([ligne_debut, pd.DataFrame(lignes_etranger), ligne_fin], ignore_index=True)
    else:
        df_etranger = pd.DataFrame(columns=df_balance.columns)
    
    return df_fr, df_etranger


def generate_balance_file(df_source):
    """
    Génère un fichier de balance à partir du DataFrame source.
    
    Args:
        df_source (DataFrame): DataFrame source.
    
    Returns:
        DataFrame: DataFrame de la balance générée.
    """
    try:
        import pandas as pd
    except Exception as e:
        msg = "Le package 'pandas' (et 'openpyxl') n'est pas installé. Installez-le avec: pip install pandas openpyxl"
        return False, msg
    
    # Créer une liste pour accumuler les lignes
    lignes = []

    # Insérer la première ligne manuellement
    lignes.append(['000000', pd.Timestamp.now().strftime('%d/%m/%Y'), '', '', pd.Timestamp.now().strftime('%d/%m/%Y'), 'EUR', 0, 0, 'DEB', '', ''])

    # Définir les constantes pour les colonnes
    CODE_VENDEUR_CEDANT = '012345'
    DATE_FICHIER = pd.Timestamp.now().strftime('%d/%m/%Y')
    DEVISE_FICHIER = 'EUR'
    NUERO_COMMANDE = ''

    # Parcourir les lignes du df source et remplir le df balance
    for _, row in df_source.iterrows():
        codeClient = row.get('Client')
        reglement = row.get('Règlement')
        numPiece = row.get('N°Fact.')
        datePiece = row.get('Date').strftime('%d/%m/%Y')
        dateEcheance = row.get('Echéance').strftime('%d/%m/%Y')
        typePiece = 'FAC' # Mettre FAC par défaut

        premiereLettreCodeReglement = str(reglement)[0]

        # Faire un switch case sur la première lettre du code règlement
        match premiereLettreCodeReglement:
            case 'C':
                codeReglement = 'CHE'
            case 'V':
                codeReglement = 'VIR'
            case 'A':
                codeReglement = '' # Supprimer le code règlement (formalisme demandé)
                typePiece = 'AVO' # Remplacer le type de pièce par AVO
        
        if typePiece == 'AVO':
            montantDevise = round(-abs(row.get('Montant T.T.C.')), 2)
        else:
            montantDevise = round(abs(row.get('Montant T.T.C.')), 2)

        lignes.append([
            CODE_VENDEUR_CEDANT,
            DATE_FICHIER,
            codeClient,
            numPiece,
            datePiece,
            DEVISE_FICHIER,
            montantDevise,
            dateEcheance,
            typePiece,
            codeReglement,
            NUERO_COMMANDE
        ])

    # Insérer la dernière ligne manuellement
    lignes.append(['999999', pd.Timestamp.now().strftime('%d/%m/%Y'), '', '', pd.Timestamp.now().strftime('%d/%m/%Y'), 'EUR', 0, 0, 'FIN', '', ''])

    # Ajout du nom des colonnes
    colonnes = [
        'Code vendeur cédant',
        'Date du fichier',
        'Code client',
        'N° de la pièce',
        'Date de la pièce',
        'Devise du fichier',
        'Montant en devise',
        'Date d\'échéance',
        'Type de la pièce',
        'Mode de règlement',
        'Numéro de la commande'
    ]
    df_balance = pd.DataFrame(lignes, columns=colonnes)
    
    return df_balance

def generate_tiers_file(df_balance):
    """
    Génère un fichier de tiers à partir du DataFrame Balance.
    
    Args:
        df_balance (DataFrame): DataFrame Balance.
    
    Returns:
        DataFrame: DataFrame des tiers généré.
        set: Ensemble des clients non identifiés.
    """
    try:
        import pandas as pd
    except Exception as e:
        msg = "Le package 'pandas' (et 'openpyxl') n'est pas installé. Installez-le avec: pip install pandas openpyxl"
        return False, msg
    
    df_tiers = pd.DataFrame()

    lignes = []

    # Insérer la première ligne manuellement
    lignes.append(['000000', 'DEB', '32038969500026', 'MONTAGE ET ASSEMBLAGE MECANIQUE', 'MONTAGE ET ASSEMBLAGE MECANIQUE', '23 RUE MELVILLE-LYNCH', 'PARC D\'ACTIVITE MAIGNON', '64100', 'BAYONNE', 'FR'])
    
    # Définir les constantes pour les colonnes
    CODE_VENDEUR_CEDANT = '012345'

    # Récupérer les données utiles
    df_clients = pd.read_csv(get_data_file_path('clients_siret.csv'), sep=';', encoding='utf-8-sig')
    df_codes_pays = pd.read_csv(get_data_file_path('codes_pays.csv'), sep=';', encoding='utf-8-sig')

    # Déclarer les clients non identifiés
    clients_non_identifies = set()
    clients_traites = set()  # Pour éviter les doublons

    # Parcourir les lignes du df balance et remplir le df tiers
    for _, row in df_balance.iterrows():
        code_client = row['Code client']
        
        # Ignorer les lignes de début/fin
        if code_client in ['000000', '999999']:
            continue
        
        # Éviter les doublons
        if code_client in clients_traites:
            continue
        
        clients_traites.add(code_client)
        
        if not code_client in df_clients['Code'].values:
            clients_non_identifies.add(str(code_client))
            continue
        
        client_info = df_clients[df_clients['Code'] == row['Code client']].iloc[0]
        
        # Fonction pour gérer les NaN
        def safe_str(val, max_len=None):
            if pd.isna(val):
                return ''
            result = str(val)
            return result[:max_len] if max_len else result
        
        lignes.append([
            CODE_VENDEUR_CEDANT,
            safe_str(client_info['Code']),
            safe_str(client_info['SIRET'], 14),
            safe_str(client_info['Raison sociale'], 40),
            safe_str(client_info['Raison sociale'], 40),
            safe_str(client_info['Voie'], 40),
            safe_str(client_info['Complement'], 40),
            safe_str(client_info['CP'], 6),
            safe_str(client_info['Ville'], 34),
            df_codes_pays.loc[df_codes_pays['Pays'] == client_info['Pays'], 'ISO'].values[0] if len(df_codes_pays.loc[df_codes_pays['Pays'] == client_info['Pays'], 'ISO'].values) > 0 else 'FR'
        ])

    # Insérer la dernière ligne manuellement
    lignes.append(['999999', 'FIN', '32038969500026', 'MONTAGE ET ASSEMBLAGE MECANIQUE', 'MONTAGE ET ASSEMBLAGE MECANIQUE', '23 RUE MELVILLE-LYNCH', 'PARC D\'ACTIVITE MAIGNON', '64100', 'BAYONNE', 'FR'])

    # Ajout du nom des colonnes
    colonnes = [
        'Code vendeur cédant',
        'Code client',
        'Identifiant du tiers',
        'Sigle du tiers',
        'Raison sociale',
        'N° et nom de la voie',
        'Complément d\'adresse',
        'Code postal',
        'Ville',
        'Code pays'
    ]

    df_tiers = pd.DataFrame(lignes, columns=colonnes)

    return df_tiers, clients_non_identifies

def export_dataframe_to_csv(df_source, type, suffixe='1A', dossier_destination=None):
    """
    Exporte le DataFrame source en fichier CSV.
    
    Args:
        df_source (DataFrame): DataFrame source.
        type (str): Type de fichier ('balance' ou 'tiers').
        suffixe (str): Suffixe du fichier ('1A' pour français, '1B' pour étranger).
        dossier_destination (str): Chemin du dossier de destination (optionnel).
    
    Returns:
        tuple: (succès: bool, message: str)
    """
    try:
        import pandas as pd
    except Exception as e:
        msg = "Le package 'pandas' (et 'openpyxl') n'est pas installé. Installez-le avec: pip install pandas openpyxl"
        return False, msg
    
    if type == 'balance':
        nom_fichier = "FBA"
        nom_fichier += "SS"
        nom_fichier += "012345"
        nom_fichier += suffixe
        nom_fichier += "."
        nom_fichier += f"{pd.Timestamp.now().timetuple().tm_yday:03d}"
    elif type == 'tiers':
        nom_fichier = "TIE"
        nom_fichier += "SS"
        nom_fichier += "012345"
        nom_fichier += suffixe
        nom_fichier += "."
        nom_fichier += f"{pd.Timestamp.now().timetuple().tm_yday:03d}"
    
    # Construire le chemin complet avec le dossier de destination
    if dossier_destination:
        chemin_complet = os.path.join(dossier_destination, nom_fichier)
    else:
        chemin_complet = nom_fichier

    try:
        df_source.to_csv(chemin_complet, index=False, header=False, sep=';', decimal=',', encoding='cp850', lineterminator='\r\n', float_format='%.2f')
        return True, f"Fichier exporté avec succès : {chemin_complet}"
    except Exception as e:
        msg = f"Erreur lors de l'exportation : {str(e)}"
        return False, msg
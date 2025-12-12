"""
Module de traitement des fichiers
Contient les fonctions de conversion et de traitement des données
"""

import os
from pathlib import Path
from re import match
from unittest import case

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
        typePiece = 'FAC'

        premiereLettreCodeReglement = str(reglement)[0]

        # Faire un switch case sur la première lettre du code règlement
        match premiereLettreCodeReglement:
            case 'T':
                codeReglement = 'TRT'
            case 'C':
                codeReglement = 'CHE'
            case 'V':
                codeReglement = 'VIR'
            case 'A':
                codeReglement = 'AVO'
        
        if codeReglement == 'AVO':
            montantDevise = round(-abs(row.get('Montant T.T.C.')), 2)
        elif codeReglement in ['VIR', 'CHE', 'TRT']:
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

    return df_tiers

def export_dataframe_to_csv(df_source, type):
    """
    Exporte le DataFrame source en fichier CSV.
    
    Args:
        df_source (DataFrame): DataFrame source.
        type (str): Type de fichier ('balance' ou 'ecritures').
    
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
        nom_fichier += "1A"
        nom_fichier += "."
        nom_fichier += f"{pd.Timestamp.now().timetuple().tm_yday:03d}"

    try:
        df_source.to_csv(nom_fichier, index=False, header=False, sep=';', encoding='utf-8-sig')
        return True, f"Fichier exporté avec succès : {nom_fichier}"
    except Exception as e:
        msg = f"Erreur lors de l'exportation : {str(e)}"
        return False, msg
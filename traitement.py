"""
Module de traitement des fichiers
Contient les fonctions de conversion et de traitement des données
"""

import os
from pathlib import Path

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
    
        # Créer un df avec les colonnes nécessaires
    df_balance = ['Code vendeur cédant', 'Date du fichier', 'Code client', 'N° de la pièce', 'Date de la pièce', 'Devise du fichier', 'Montant en devise', 'Date d\'échéance', 'Type de la pièce', 'Mode de règlement', 'Numéro de la commande']

    # Insérer la première ligne manuellement
    df_balance.loc[len(df_balance)] = ['000000', pd.Timestamp.now().strftime('%d/%m/%Y'), '', '', pd.Timestamp.now().strftime('%d/%m/%Y'), 'EUR', 0, 0, 'DEB', '', '']

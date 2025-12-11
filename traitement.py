"""
Module de traitement des fichiers
Contient les fonctions de conversion et de traitement des données
"""

import os
from pathlib import Path

def convertir_fichier(chemin_fichier, chemin_sortie=None, sheet_name=0):
    """
    Convertit un fichier Excel (.xlsx/.xls/.xlsm) en CSV.

    Args:
        chemin_fichier (str): Chemin du fichier source.
        chemin_sortie (str|None): Chemin du fichier CSV de sortie. Si None, remplace l'extension par .csv.
        sheet_name (int|str): Index ou nom de la feuille à lire (par défaut 0).

    Returns:
        tuple: (succès: bool, message: str)
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

        if chemin_sortie is None:
            chemin_sortie = str(Path(chemin_fichier).with_suffix('.csv'))

        # Lire et écrire
        df = pd.read_excel(chemin_fichier, sheet_name=sheet_name, engine="openpyxl")
        df.to_csv(chemin_sortie, index=False)

        return True, f"Fichier converti avec succès en : {chemin_sortie}"
    except Exception as e:
        msg = f"Erreur lors de la conversion : {str(e)}"
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

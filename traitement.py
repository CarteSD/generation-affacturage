"""
Module de traitement des fichiers
Contient les fonctions de conversion et de traitement des données
"""

import os
from pathlib import Path


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

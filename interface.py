import tkinter as tk
from tkinter import filedialog, messagebox
import os
from traitement import valider_fichier, convertir_fichier, generate_balance_file, generate_tiers_file, export_dataframe_to_csv, separer_clients_par_pays
import pandas as pd


class ConversionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Convertisseur CSV")
        self.root.geometry("500x200")
        self.root.resizable(False, False)
        
        # Définir l'icône de la fenêtre
        try:
            if os.path.exists("burographic.ico"):
                self.root.iconbitmap("burographic.ico")
        except Exception:
            pass  # Ignorer si l'icône n'est pas trouvée
        
        self.fichier_selectionne = None
        self.dataframe = None  # Stocke le DataFrame en mémoire
        self.dossier_destination = None  # Dossier d'export
        
        # Frame principal
        main_frame = tk.Frame(root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Label pour le fichier
        self.label_fichier = tk.Label(
            main_frame, 
            text="Aucun fichier sélectionné",
            wraplength=450,
            justify=tk.LEFT
        )
        self.label_fichier.pack(pady=(0, 15))
        
        # Bouton pour sélectionner le fichier
        btn_parcourir = tk.Button(
            main_frame,
            text="📁 Parcourir...",
            command=self.choisir_fichier,
            width=20,
            height=2
        )
        btn_parcourir.pack(pady=5)
        
        # Bouton pour lancer la conversion
        self.btn_convertir = tk.Button(
            main_frame,
            text="Lancer la conversion",
            command=self.lancer_conversion,
            width=20,
            height=2,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 10, "bold"),
            state=tk.DISABLED
        )
        self.btn_convertir.pack(pady=5)
    
    def choisir_fichier(self):
        """Ouvre la boîte de dialogue pour choisir un fichier"""
        fichier = filedialog.askopenfilename(
            title="Sélectionner un fichier",
            filetypes=[
                ("Fichiers Excel", "*.xsls;*.xlsx;*.xlsm"),
                ("Tous les fichiers", "*.*")
            ]
        )
        
        if fichier:
            self.fichier_selectionne = fichier
            nom_fichier = os.path.basename(fichier)
            self.label_fichier.config(text=f"Fichier sélectionné : {nom_fichier}")
            self.btn_convertir.config(state=tk.NORMAL)
    
    def lancer_conversion(self):
        """Lance la conversion du fichier sélectionné"""
        if not self.fichier_selectionne:
            messagebox.showwarning("Attention", "Aucun fichier sélectionné")
            return
        
        # Demander le dossier de destination
        self.dossier_destination = filedialog.askdirectory(
            title="Choisir le dossier de destination pour les fichiers CSV",
            initialdir=os.path.dirname(self.fichier_selectionne)
        )
        
        if not self.dossier_destination:
            messagebox.showwarning("Attention", "Aucun dossier de destination sélectionné")
            return
        
        # Valider le fichier
        valide, message_validation = valider_fichier(self.fichier_selectionne)
        if not valide:
            messagebox.showerror("Erreur", message_validation)
            return
        
        # Convertir le fichier
        success, resultat = convertir_fichier(self.fichier_selectionne)
        
        if success:
            self.dataframe = resultat
            nb_lignes, nb_colonnes = self.dataframe.shape
            messagebox.showinfo(
                "Succès", 
                f"Fichier chargé avec succès !\n\nLignes : {nb_lignes}\nColonnes : {nb_colonnes}"
            )
        else:
            messagebox.showerror("Erreur", resultat)
            return
        
        # Générer le Dataframe Balance
        df_balance = generate_balance_file(self.dataframe)

        # Charger les données clients pour la séparation
        df_clients = pd.read_csv('datas/clients_siret.csv', sep=';', encoding='utf-8-sig')
        
        # Séparer les clients français et étrangers
        df_balance_fr, df_balance_etranger = separer_clients_par_pays(df_balance, df_clients)
        
        fichiers_exportes = []
        tous_clients_non_identifies = set()
        
        # Traiter les clients français (1A)
        if not df_balance_fr.empty and len(df_balance_fr) > 2:  # Plus que juste les lignes début/fin
            
            # Exporter Balance FR
            success, message = export_dataframe_to_csv(df_balance_fr, "balance", "1A", self.dossier_destination)
            if success:
                fichiers_exportes.append(message)
            else:
                messagebox.showerror("Erreur", message)
                return
            
            # Générer et exporter Tiers FR
            df_tiers_fr, clients_non_identifies_fr = generate_tiers_file(df_balance_fr.copy())
            tous_clients_non_identifies.update(clients_non_identifies_fr)
            
            success, message = export_dataframe_to_csv(df_tiers_fr, "tiers", "1A", self.dossier_destination)
            if success:
                fichiers_exportes.append(message)
            else:
                messagebox.showerror("Erreur", message)
                return
        
        # Traiter les clients étrangers (1B)
        if not df_balance_etranger.empty and len(df_balance_etranger) > 2:  # Plus que juste les lignes début/fin
            
            # Exporter Balance étranger
            success, message = export_dataframe_to_csv(df_balance_etranger, "balance", "1B", self.dossier_destination)
            if success:
                fichiers_exportes.append(message)
            else:
                messagebox.showerror("Erreur", message)
                return
            
            # Générer et exporter Tiers étranger
            df_tiers_etranger, clients_non_identifies_etr = generate_tiers_file(df_balance_etranger.copy())
            tous_clients_non_identifies.update(clients_non_identifies_etr)
            
            success, message = export_dataframe_to_csv(df_tiers_etranger, "tiers", "1B", self.dossier_destination)
            if success:
                fichiers_exportes.append(message)
            else:
                messagebox.showerror("Erreur", message)
                return
        
        # Afficher les clients non identifiés s'il y en a
        # Filtrer les valeurs vides
        clients_valides = {c for c in tous_clients_non_identifies if c and str(c).strip() and str(c) not in ['000000', '999999']}
        if clients_valides:
            messagebox.showwarning(
                "Clients non identifiés",
                f"Les clients suivants n'ont pas été identifiés :\n{', '.join(sorted(clients_valides))}\n\nVeuillez vérifier les codes clients."
            )
        
        # Afficher un message de succès global
        message_final = "Conversion terminée avec succès !\n\nFichiers exportés :\n" + "\n".join(fichiers_exportes)
        messagebox.showinfo("Succès", message_final)

def main():
    root = tk.Tk()
    app = ConversionApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

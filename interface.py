import tkinter as tk
from tkinter import filedialog, messagebox
import os
from traitement import valider_fichier, convertir_fichier, generate_balance_file, generate_tiers_file, export_dataframe_to_csv


class ConversionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Convertisseur CSV")
        self.root.geometry("500x200")
        self.root.resizable(False, False)
        
        self.fichier_selectionne = None
        self.dataframe = None  # Stocke le DataFrame en mémoire
        
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
        
        # Valider le fichier
        valide, message_validation = valider_fichier(self.fichier_selectionne)
        if not valide:
            messagebox.showerror("Erreur", message_validation)
            return
        
        # Convertir le fichier
        succes, resultat = convertir_fichier(self.fichier_selectionne)
        
        if succes:
            self.dataframe = resultat
            nb_lignes, nb_colonnes = self.dataframe.shape
            messagebox.showinfo(
                "Succès", 
                f"Fichier chargé avec succès !\n\nLignes : {nb_lignes}\nColonnes : {nb_colonnes}"
            )
        else:
            messagebox.showerror("Erreur", resultat)
            return
        
        # Générer le dataframe Balance
        df_balance = generate_balance_file(self.dataframe)
        print("DataFrame Balance généré :")
        print(df_balance)

        # Générer le dataframe Tiers à partir de la Balance
        df_tiers, clients_non_identifies = generate_tiers_file(df_balance.copy())
        print("DataFrame Tiers généré :")
        print(df_tiers)
        if clients_non_identifies:
            messagebox.showwarning(
                "Clients non identifiés",
                f"Les clients suivants n'ont pas été identifiés :\n{', '.join(clients_non_identifies)}"
            )

        # Exporter le DataFrame Balance
        success, message = export_dataframe_to_csv(df_balance, "balance")
        if success:
            messagebox.showinfo("Succès", message)
        else:
            messagebox.showerror("Erreur", message)

        success, message = export_dataframe_to_csv(df_tiers, "tiers")
        if success:
            messagebox.showinfo("Succès", message)
        else:
            messagebox.showerror("Erreur", message)

def main():
    root = tk.Tk()
    app = ConversionApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

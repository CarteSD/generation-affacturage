import tkinter as tk
from tkinter import filedialog, messagebox
import os
from traitement import valider_fichier


class ConversionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Convertisseur CSV")
        self.root.geometry("500x200")
        self.root.resizable(False, False)
        
        self.fichier_selectionne = None
        
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
        
        messagebox.showinfo("Succès", "Le fichier est correctement validé et prêt pour la conversion.")
        

def main():
    root = tk.Tk()
    app = ConversionApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

from bot_eshop_core import process_excel


class CMSDetectorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CMS Detector – Excel")
        self.geometry("520x200")
        self.resizable(False, False)

        self.input_file = None

        tk.Label(self, text="Sélectionne un fichier Excel contenant des emails",
                 font=("Arial", 12)).pack(pady=15)

        tk.Button(self, text="📂 Choisir un fichier Excel",
                  command=self.choose_file, width=30).pack(pady=5)

        self.label_file = tk.Label(self, text="Aucun fichier sélectionné", fg="gray")
        self.label_file.pack(pady=5)

        tk.Button(self, text="🚀 Lancer la détection",
                  command=self.run_detection, width=30).pack(pady=15)

    def choose_file(self):
        file = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file:
            self.input_file = file
            self.label_file.config(text=os.path.basename(file), fg="black")

    def run_detection(self):
        if not self.input_file:
            messagebox.showwarning("Attention", "Veuillez sélectionner un fichier Excel.")
            return

        try:
            output = process_excel(self.input_file)
            messagebox.showinfo(
                "Terminé",
                f"Fichier généré avec succès :\n{output}"
            )
        except Exception as e:
            messagebox.showerror("Erreur", str(e))


if __name__ == "__main__":
    app = CMSDetectorApp()
    app.mainloop()

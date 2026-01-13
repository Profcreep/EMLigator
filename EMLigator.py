import sys
import os
import glob
import tempfile
from pathlib import Path
from tkinter import filedialog, messagebox
import unicodedata
import customtkinter as ctk

# Import du module de conversion
import outlookmsgfile  # Assure-toi que ce module est installé/prêt

# ---------------------------
# Icône de la fenêtre
# ---------------------------
icon_path = Path(__file__).parent / "resources" / "EMLigator.ico"

# ---------------------------
# Nom de fichier sûr
# ---------------------------
def safe_filename(filename):
    filename = unicodedata.normalize('NFKD', filename).encode('ascii', 'ignore').decode('ascii')
    filename = "".join(c if c.isalnum() or c in " -_" else "_" for c in filename)
    return filename[:100] if len(filename) > 100 else filename

# ---------------------------
# Conversion avec convert‑outlook‑msg‑file
# ---------------------------
def convert_msg_to_eml(msg_path, output_dir=None):
    msg_path = Path(msg_path).resolve()
    if not msg_path.exists():
        raise FileNotFoundError(f"Fichier introuvable : {msg_path}")

    if output_dir is None:
        output_dir = msg_path.parent.resolve()
    else:
        output_dir = Path(output_dir).resolve()
        output_dir.mkdir(parents=True, exist_ok=True)

    # Définir le nom final
    eml_name = safe_filename(msg_path.stem) + ".eml"
    eml_path = output_dir / eml_name

    try:
        # Charge le .msg avec outlookmsgfile
        eml_message = outlookmsgfile.load(str(msg_path))

        # Génère le contenu MIME (.eml)
        with open(eml_path, "wb") as f:
            f.write(eml_message.as_bytes())

        return eml_path

    except Exception as e:
        raise RuntimeError(f"Impossible de convertir {msg_path.name} : {e}")

# ---------------------------
# Drag & Drop
# ---------------------------
def handle_drag_drop(files):
    converted = []
    for f in files:
        p = Path(f).resolve()
        if p.suffix.lower() == ".msg":
            try:
                eml_file = convert_msg_to_eml(p)
                converted.append(eml_file)
                print(f"Converti : {eml_file}")
            except Exception as e:
                print(f"Erreur pour {p}: {e}")
    if converted:
        messagebox.showinfo("EMLigator", f"Conversion terminée !\n{len(converted)} fichier(s) converti(s)")
    else:
        messagebox.showwarning("EMLigator", "Aucun fichier .msg valide trouvé.")

# ---------------------------
# GUI
# ---------------------------
class EMLigatorApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("EMLigator")
        self.geometry("500x300")
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")

        if icon_path.exists():
            self.iconbitmap(str(icon_path))

        ctk.CTkLabel(self, text="EMLigator", font=("Arial", 24)).pack(pady=20)
        ctk.CTkButton(self, text="Convertir un fichier", command=self.single_file).pack(pady=10)
        ctk.CTkButton(self, text="Convertir un dossier entier", command=self.batch_folder).pack(pady=10)
        ctk.CTkLabel(self, text="Fait par Profcreep", font=("Arial", 12), text_color="gray").pack(side="bottom", pady=10)

    def single_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("MSG files", "*.msg")])
        if file_path:
            try:
                eml_file = convert_msg_to_eml(file_path)
                messagebox.showinfo("EMLigator", f"Converti en : {eml_file}")
            except Exception as e:
                messagebox.showerror("EMLigator", f"Erreur : {e}")

    def batch_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            converted = []
            for f in glob.glob(os.path.join(folder, "*.msg")):
                try:
                    eml_file = convert_msg_to_eml(f)
                    converted.append(eml_file)
                except Exception as e:
                    print(f"Erreur pour {f}: {e}")
            messagebox.showinfo("EMLigator", f"Conversion terminée !\n{len(converted)} fichier(s) converti(s)")

# ---------------------------
# Main
# ---------------------------
if __name__ == "__main__":
    if len(sys.argv) > 1:
        handle_drag_drop(sys.argv[1:])
    else:
        app = EMLigatorApp()
        app.mainloop()

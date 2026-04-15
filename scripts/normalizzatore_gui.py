import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
import threading

class NormalizzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("🔧 Normalizzatore di Nomi (File e Cartelle)")
        self.root.geometry("900x700")
        self.root.resizable(True, True)

        # Variabili
        self.selected_path = tk.StringVar()
        self.log_text = None
        self.is_processing = False

        self.setup_ui()

    def setup_ui(self):
        """Crea l'interfaccia grafica"""

        # ===== HEADER =====
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill=tk.X, padx=15, pady=15)

        title = ttk.Label(header_frame, text="🔧 Normalizzatore di Nomi", 
                         font=("Arial", 16, "bold"))
        title.pack(anchor=tk.W)

        subtitle = ttk.Label(header_frame, 
                            text="Converti nomi: 'Alberto Sozio' → 'alberto-sozio' | 'image 1.jpg' → 'image-1.jpg'",
                            font=("Arial", 10), foreground="gray")
        subtitle.pack(anchor=tk.W, pady=(5, 0))

        # ===== SELEZIONE CARTELLA =====
        path_frame = ttk.LabelFrame(self.root, text="📁 Cartella da Normalizzare", padding=10)
        path_frame.pack(fill=tk.X, padx=15, pady=10)

        path_display = ttk.Entry(path_frame, textvariable=self.selected_path, 
                                state="readonly", font=("Arial", 10))
        path_display.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        browse_btn = ttk.Button(path_frame, text="📂 Sfoglia", 
                               command=self.browse_folder)
        browse_btn.pack(side=tk.LEFT)

        use_current_btn = ttk.Button(path_frame, text="📍 Usa Corrente", 
                                    command=self.use_current_folder)
        use_current_btn.pack(side=tk.LEFT, padx=(5, 0))

        # ===== OPZIONI =====
        options_frame = ttk.LabelFrame(self.root, text="⚙️ Opzioni", padding=10)
        options_frame.pack(fill=tk.X, padx=15, pady=10)

        self.recursive_var = tk.BooleanVar(value=True)
        recursive_check = ttk.Checkbutton(options_frame, text="Elabora sottocartelle (ricorsivo)",
                                         variable=self.recursive_var)
        recursive_check.pack(anchor=tk.W)

        self.show_details_var = tk.BooleanVar(value=True)
        details_check = ttk.Checkbutton(options_frame, text="Mostra dettagli operazioni",
                                       variable=self.show_details_var)
        details_check.pack(anchor=tk.W, pady=(5, 0))

        # ===== LOG =====
        log_frame = ttk.LabelFrame(self.root, text="📋 Log Operazioni", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, width=100,
                                                 font=("Courier", 9), wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Colori per il log
        self.log_text.tag_config("info", foreground="#0066CC")
        self.log_text.tag_config("file", foreground="#009900")
        self.log_text.tag_config("folder", foreground="#FF8C00")
        self.log_text.tag_config("success", foreground="#00AA00", font=("Courier", 9, "bold"))
        self.log_text.tag_config("error", foreground="#CC0000", font=("Courier", 9, "bold"))
        self.log_text.tag_config("warning", foreground="#FF6600")

        # ===== PULSANTI AZIONE =====
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill=tk.X, padx=15, pady=15)

        self.start_btn = ttk.Button(button_frame, text="▶️ Avvia Normalizzazione",
                                   command=self.start_normalization)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 10))

        clear_btn = ttk.Button(button_frame, text="🗑️ Pulisci Log",
                              command=self.clear_log)
        clear_btn.pack(side=tk.LEFT, padx=(0, 10))

        quit_btn = ttk.Button(button_frame, text="❌ Esci",
                             command=self.root.quit)
        quit_btn.pack(side=tk.LEFT)

        # Info iniziale
        self.log_add("✅ Pronto. Seleziona una cartella e clicca 'Avvia Normalizzazione'", "info")

    def log_add(self, text, tag="info"):
        """Aggiunge testo al log"""
        self.log_text.insert(tk.END, text + "\n", tag)
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def clear_log(self):
        """Pulisce il log"""
        self.log_text.delete(1.0, tk.END)

    def browse_folder(self):
        """Apre dialog per selezionare cartella"""
        folder = filedialog.askdirectory(title="Seleziona cartella da normalizzare")
        if folder:
            self.selected_path.set(folder)
            self.log_add(f"📁 Cartella selezionata: {folder}", "info")

    def use_current_folder(self):
        """Usa la cartella corrente"""
        current = os.getcwd()
        self.selected_path.set(current)
        self.log_add(f"📍 Cartella corrente: {current}", "info")

    def normalize_name(self, name):
        """Converte il nome in minuscolo con trattini"""
        if '.' in name and not name.startswith('.'):
            name_part, ext = name.rsplit('.', 1)
            normalized = name_part.lower().replace(' ', '-')
            return f"{normalized}.{ext}"
        else:
            return name.lower().replace(' ', '-')

    def normalize_folder(self, root_path):
        """Normalizza ricorsivamente cartella e file"""
        root_path = Path(root_path)
        renamed = {'files': 0, 'folders': 0}

        try:
            # Processa dal basso verso l'alto per evitare problemi
            for dirpath, dirnames, filenames in os.walk(root_path, topdown=False):
                current_dir = Path(dirpath)

                # Rinomina file
                for filename in filenames:
                    old_path = current_dir / filename
                    new_name = self.normalize_name(filename)

                    if new_name != filename:
                        new_path = current_dir / new_name
                        if self.show_details_var.get():
                            self.log_add(f"  📄 {filename} → {new_name}", "file")
                        old_path.rename(new_path)
                        renamed['files'] += 1

                # Rinomina cartelle
                for dirname in dirnames:
                    old_path = current_dir / dirname
                    new_name = self.normalize_name(dirname)

                    if new_name != dirname:
                        new_path = current_dir / new_name
                        if self.show_details_var.get():
                            self.log_add(f"  📁 {dirname} → {new_name}", "folder")
                        old_path.rename(new_path)
                        renamed['folders'] += 1

            return renamed, None

        except Exception as e:
            return renamed, str(e)

    def start_normalization(self):
        """Avvia la normalizzazione in thread separato"""
        if not self.selected_path.get():
            messagebox.showwarning("Attenzione", "Seleziona una cartella!")
            return

        path = Path(self.selected_path.get())
        if not path.exists():
            messagebox.showerror("Errore", f"Cartella non trovata: {path}")
            return

        if not path.is_dir():
            messagebox.showerror("Errore", "Il percorso selezionato non è una cartella")
            return

        # Chiedi conferma
        if messagebox.askyesno("Conferma", 
                              f"Normalizzare i nomi in:\n{path}\n\nContinuare?"):
            # Disabilita il bottone durante l'operazione
            self.start_btn.config(state=tk.DISABLED)
            self.is_processing = True

            # Esegui in thread separato per non bloccare l'interfaccia
            thread = threading.Thread(target=self._do_normalization, args=(path,))
            thread.daemon = True
            thread.start()

    def _do_normalization(self, path):
        """Esegue la normalizzazione (in thread)"""
        try:
            self.log_add("\n" + "="*70, "info")
            self.log_add("⏳ Elaborazione in corso...", "warning")
            self.log_add("="*70, "info")

            renamed, error = self.normalize_folder(path)

            if error:
                self.log_add(f"\n❌ ERRORE: {error}", "error")
            else:
                self.log_add("\n" + "="*70, "success")
                self.log_add("✅ OPERAZIONE COMPLETATA", "success")
                self.log_add("="*70, "success")
                self.log_add(f"📁 Cartelle rinominate: {renamed['folders']}", "success")
                self.log_add(f"📄 File rinominati: {renamed['files']}", "success")
                self.log_add(f"📊 Totale: {renamed['folders'] + renamed['files']}", "success")
                self.log_add("="*70 + "\n", "success")

                messagebox.showinfo("Successo", 
                                  f"Operazione completata!\n\n"
                                  f"Cartelle: {renamed['folders']}\n"
                                  f"File: {renamed['files']}\n"
                                  f"Totale: {renamed['folders'] + renamed['files']}")

        except Exception as e:
            self.log_add(f"\n❌ ERRORE CRITICO: {str(e)}", "error")
            messagebox.showerror("Errore", f"Si è verificato un errore:\n{str(e)}")

        finally:
            self.is_processing = False
            self.start_btn.config(state=tk.NORMAL)

def main():
    root = tk.Tk()
    app = NormalizzerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

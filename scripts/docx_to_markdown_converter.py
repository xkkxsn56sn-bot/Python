#!/usr/bin/env python3
"""
DOCX to Markdown Converter - Applicazione per macOS
Converte documenti DOCX in formato Markdown usando Pandoc
"""

import os
import sys
import subprocess
import threading
from pathlib import Path
from typing import List, Optional
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime


class DOCXtoMarkdownConverter:
    """Interfaccia grafica per il convertitore DOCX → Markdown"""

    def __init__(self, root):
        self.root = root
        self.root.title("DOCX → Markdown Converter")
        self.root.geometry("800x600")

        # Variabili
        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.extract_media = tk.BooleanVar(value=True)
        self.standalone = tk.BooleanVar(value=True)
        self.wrap_none = tk.BooleanVar(value=True)
        self.files_found = []
        self.converting = False

        self.setup_ui()

    def setup_ui(self):
        """Configura l'interfaccia utente"""

        # Stile
        style = ttk.Style()
        style.theme_use('aqua')

        # Frame principale
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configurazione griglia
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # Titolo
        title = ttk.Label(main_frame, text="📝 DOCX → Markdown Converter", 
                         font=('Helvetica', 16, 'bold'))
        title.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # --- SEZIONE INPUT ---
        input_frame = ttk.LabelFrame(main_frame, text="📥 Cartella File DOCX", padding="10")
        input_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        input_frame.columnconfigure(0, weight=1)

        ttk.Entry(input_frame, textvariable=self.input_dir, width=50).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(input_frame, text="Sfoglia...", 
                  command=self.browse_input).grid(row=0, column=1)
        ttk.Button(input_frame, text="🔍 Scansiona", 
                  command=self.scan_files).grid(row=0, column=2, padx=(5, 0))

        # --- SEZIONE OUTPUT ---
        output_frame = ttk.LabelFrame(main_frame, text="📤 Cartella Destinazione Markdown", padding="10")
        output_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        output_frame.columnconfigure(0, weight=1)

        ttk.Entry(output_frame, textvariable=self.output_dir, width=50).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(output_frame, text="Sfoglia...", 
                  command=self.browse_output).grid(row=0, column=1)

        # --- OPZIONI CONVERSIONE ---
        options_frame = ttk.LabelFrame(main_frame, text="⚙️ Opzioni Pandoc", padding="10")
        options_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        ttk.Checkbutton(options_frame, text="Estrai immagini in cartella separata (--extract-media)", 
                       variable=self.extract_media).grid(row=0, column=0, sticky=tk.W, pady=2)

        ttk.Checkbutton(options_frame, text="Documento standalone (--standalone)", 
                       variable=self.standalone).grid(row=1, column=0, sticky=tk.W, pady=2)

        ttk.Checkbutton(options_frame, text="Nessun wrap delle righe (--wrap=none)", 
                       variable=self.wrap_none).grid(row=2, column=0, sticky=tk.W, pady=2)

        # --- LISTA FILE ---
        files_frame = ttk.LabelFrame(main_frame, text="📄 File Trovati", padding="10")
        files_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        files_frame.rowconfigure(0, weight=1)
        files_frame.columnconfigure(0, weight=1)

        # Treeview per lista file
        self.file_tree = ttk.Treeview(files_frame, columns=('path', 'size'), 
                                      height=8, selectmode='extended')
        self.file_tree.heading('#0', text='Nome File')
        self.file_tree.heading('path', text='Percorso')
        self.file_tree.heading('size', text='Dimensione')
        self.file_tree.column('#0', width=200)
        self.file_tree.column('path', width=350)
        self.file_tree.column('size', width=100)

        scrollbar = ttk.Scrollbar(files_frame, orient=tk.VERTICAL, 
                                 command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=scrollbar.set)

        self.file_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # --- CONSOLE LOG ---
        log_frame = ttk.LabelFrame(main_frame, text="📋 Log Conversione", padding="10")
        log_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, 
                                                  state='disabled', wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # --- BARRA PROGRESSO ---
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        progress_frame.columnconfigure(0, weight=1)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                           maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E))

        self.status_label = ttk.Label(progress_frame, text="Pronto")
        self.status_label.grid(row=1, column=0, sticky=tk.W)

        # --- PULSANTI AZIONE ---
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=3, pady=10)

        self.convert_btn = ttk.Button(button_frame, text="▶️ Avvia Conversione", 
                                     command=self.start_conversion)
        self.convert_btn.grid(row=0, column=0, padx=5)

        ttk.Button(button_frame, text="🗑️ Pulisci Log", 
                  command=self.clear_log).grid(row=0, column=1, padx=5)

        ttk.Button(button_frame, text="❌ Esci", 
                  command=self.root.quit).grid(row=0, column=2, padx=5)

        # Configura resize
        main_frame.rowconfigure(4, weight=1)
        main_frame.rowconfigure(5, weight=1)

        # Log iniziale
        self.log("✓ Applicazione avviata")
        self.check_dependencies()

    def log(self, message: str):
        """Aggiunge messaggio al log"""
        self.log_text.configure(state='normal')
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state='disabled')

    def browse_input(self):
        """Seleziona cartella input"""
        directory = filedialog.askdirectory(title="Seleziona cartella file DOCX")
        if directory:
            self.input_dir.set(directory)
            self.log(f"📁 Cartella input: {directory}")

    def browse_output(self):
        """Seleziona cartella output"""
        directory = filedialog.askdirectory(title="Seleziona cartella destinazione")
        if directory:
            self.output_dir.set(directory)
            self.log(f"📁 Cartella output: {directory}")

    def check_dependencies(self):
        """Verifica dipendenze installate"""
        self.log("🔍 Verifica dipendenze...")

        # Controlla Pandoc
        try:
            result = subprocess.run(['pandoc', '--version'], 
                                  capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                version = result.stdout.split('\n')[0]
                self.log(f"✓ {version}")
            else:
                self.log("⚠️  Pandoc non trovato")
                self.show_install_instructions()
        except FileNotFoundError:
            self.log("❌ Pandoc NON installato")
            self.show_install_instructions()

    def show_install_instructions(self):
        """Mostra istruzioni installazione"""
        msg = """Pandoc è necessario per la conversione DOCX → Markdown.

Installazione su macOS:

1. Con Homebrew (consigliato):
   brew install pandoc

2. Download diretto:
   https://pandoc.org/installing.html

Dopo l'installazione, riavvia l'applicazione."""

        messagebox.showwarning("Dipendenza mancante", msg)

    def scan_files(self):
        """Scansiona cartella per file DOCX"""
        input_path = self.input_dir.get()

        if not input_path:
            messagebox.showwarning("Attenzione", "Seleziona prima una cartella input")
            return

        if not os.path.exists(input_path):
            messagebox.showerror("Errore", f"Cartella non trovata: {input_path}")
            return

        self.log(f"🔍 Scansione cartella: {input_path}")

        # Pulisci lista precedente
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        self.files_found = []

        # Cerca file .docx (escludi file temporanei che iniziano con ~$)
        path = Path(input_path)
        docx_files = [f for f in path.glob('*.docx') if not f.name.startswith('~$')]
        docx_files += [f for f in path.glob('**/*.docx') if not f.name.startswith('~$')]

        for file in docx_files:
            size = file.stat().st_size
            size_str = self.format_size(size)

            self.file_tree.insert('', tk.END, text=file.name, 
                                 values=(str(file.parent), size_str))
            self.files_found.append(file)

        self.log(f"✓ Trovati {len(self.files_found)} file DOCX")

        if len(self.files_found) == 0:
            messagebox.showinfo("Nessun file", 
                              "Nessun file .docx trovato nella cartella selezionata")

    def format_size(self, size: int) -> str:
        """Formatta dimensione file"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.1f} {unit}"
            size /= 1024.0
        return f"{size:.1f} TB"

    def start_conversion(self):
        """Avvia processo di conversione"""
        if self.converting:
            messagebox.showwarning("Attenzione", "Conversione già in corso")
            return

        if not self.files_found:
            messagebox.showwarning("Attenzione", "Nessun file da convertire. Scansiona prima la cartella.")
            return

        output_path = self.output_dir.get()
        if not output_path:
            messagebox.showwarning("Attenzione", "Seleziona una cartella di destinazione")
            return

        # Crea cartella output se non esiste
        Path(output_path).mkdir(parents=True, exist_ok=True)

        # Avvia conversione in thread separato
        self.converting = True
        self.convert_btn.configure(state='disabled')

        thread = threading.Thread(target=self.convert_files, daemon=True)
        thread.start()

    def convert_files(self):
        """Esegue conversione file"""
        output_path = Path(self.output_dir.get())
        total = len(self.files_found)
        success = 0
        failed = 0

        self.log("\n🚀 Avvio conversione DOCX → Markdown...")
        self.status_label.configure(text="Conversione in corso...")

        # Crea cartella media se necessario
        if self.extract_media.get():
            media_dir = output_path / "media"
            media_dir.mkdir(exist_ok=True)

        for idx, file in enumerate(self.files_found, 1):
            self.log(f"\n[{idx}/{total}] Conversione: {file.name}")

            output_file = output_path / f"{file.stem}.md"

            try:
                # Costruisci comando Pandoc
                cmd = [
                    'pandoc',
                    str(file),
                    '-o', str(output_file)
                ]

                # Aggiungi opzioni se selezionate
                if self.extract_media.get():
                    cmd.extend(['--extract-media', str(output_path / 'media')])

                if self.wrap_none.get():
                    cmd.append('--wrap=none')

                if self.standalone.get():
                    cmd.append('--standalone')

                # Esegui conversione
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

                if result.returncode == 0:
                    self.log(f"  ✓ Salvato: {output_file.name}")
                    success += 1
                else:
                    self.log(f"  ✗ Errore: {result.stderr[:100]}")
                    failed += 1

            except subprocess.TimeoutExpired:
                self.log(f"  ✗ Timeout: file troppo grande o complesso")
                failed += 1
            except Exception as e:
                self.log(f"  ✗ Errore: {str(e)}")
                failed += 1

            # Aggiorna progress bar
            progress = (idx / total) * 100
            self.progress_var.set(progress)
            self.root.update_idletasks()

        # Riepilogo finale
        self.log("\n" + "="*60)
        self.log("📊 RIEPILOGO CONVERSIONE")
        self.log("="*60)
        self.log(f"File totali:     {total}")
        self.log(f"Convertiti:      {success} ✓")
        self.log(f"Falliti:         {failed} ✗")
        self.log(f"Percentuale:     {(success/total*100):.1f}%")
        self.log(f"Output salvato:  {output_path}")
        if self.extract_media.get():
            self.log(f"Immagini:        {output_path / 'media'}")
        self.log("="*60)

        self.status_label.configure(text=f"Completato: {success}/{total} file convertiti")
        self.convert_btn.configure(state='normal')
        self.converting = False

        # Mostra messaggio finale
        msg = f"Convertiti {success} file su {total}\n\nFile Markdown salvati in:\n{output_path}"
        if self.extract_media.get():
            msg += f"\n\nImmagini estratte in:\n{output_path / 'media'}"

        messagebox.showinfo("Conversione completata", msg)

    def clear_log(self):
        """Pulisce il log"""
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state='disabled')
        self.log("✓ Log pulito")


def main():
    """Funzione principale"""
    root = tk.Tk()
    app = DOCXtoMarkdownConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
WPS to Markdown Converter - Soluzione Completa per macOS
Conversione in due fasi: WPS → DOCX → Markdown
Usa LibreOffice per WPS→DOCX e Pandoc per DOCX→MD
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


class WPStoMarkdownConverter:
    """Interfaccia grafica per il convertitore completo WPS → DOCX → Markdown"""

    def __init__(self, root):
        self.root = root
        self.root.title("WPS → Markdown Converter (Completo)")
        self.root.geometry("900x700")

        # Variabili
        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.keep_docx = tk.BooleanVar(value=False)
        self.files_found = []
        self.converting = False
        self.libreoffice_path = self.find_libreoffice()
        self.pandoc_available = self.check_pandoc()

        self.setup_ui()

    def find_libreoffice(self):
        """Trova il percorso di LibreOffice"""
        possible_paths = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "/usr/local/bin/soffice",
            "/opt/homebrew/bin/soffice"
        ]

        for path in possible_paths:
            if os.path.exists(path):
                return path

        try:
            result = subprocess.run(['which', 'soffice'], 
                                  capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                return result.stdout.strip()
        except:
            pass

        return None

    def check_pandoc(self):
        """Verifica se Pandoc è installato"""
        try:
            result = subprocess.run(['pandoc', '--version'], 
                                  capture_output=True, text=True, timeout=5)
            return result.returncode == 0
        except FileNotFoundError:
            return False

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
        title = ttk.Label(main_frame, text="🔄 WPS → Markdown Converter", 
                         font=('Helvetica', 16, 'bold'))
        title.grid(row=0, column=0, columnspan=3, pady=(0, 10))

        subtitle = ttk.Label(main_frame, text="Conversione completa: WPS → DOCX → Markdown", 
                            font=('Helvetica', 10), foreground='gray')
        subtitle.grid(row=1, column=0, columnspan=3, pady=(0, 15))

        # --- SEZIONE INPUT ---
        input_frame = ttk.LabelFrame(main_frame, text="📥 Cartella File WPS", padding="10")
        input_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        input_frame.columnconfigure(0, weight=1)

        ttk.Entry(input_frame, textvariable=self.input_dir, width=50).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(input_frame, text="Sfoglia...", 
                  command=self.browse_input).grid(row=0, column=1)
        ttk.Button(input_frame, text="🔍 Scansiona", 
                  command=self.scan_files).grid(row=0, column=2, padx=(5, 0))

        # --- SEZIONE OUTPUT ---
        output_frame = ttk.LabelFrame(main_frame, text="📤 Cartella Destinazione Markdown", padding="10")
        output_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        output_frame.columnconfigure(0, weight=1)

        ttk.Entry(output_frame, textvariable=self.output_dir, width=50).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(output_frame, text="Sfoglia...", 
                  command=self.browse_output).grid(row=0, column=1)

        # Checkbox per mantenere DOCX
        ttk.Checkbutton(output_frame, text="Mantieni file DOCX intermedi", 
                       variable=self.keep_docx).grid(row=1, column=0, sticky=tk.W, pady=(5, 0))

        # --- LISTA FILE ---
        files_frame = ttk.LabelFrame(main_frame, text="📄 File Trovati", padding="10")
        files_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        files_frame.rowconfigure(0, weight=1)
        files_frame.columnconfigure(0, weight=1)

        # Treeview per lista file
        self.file_tree = ttk.Treeview(files_frame, columns=('path', 'size'), 
                                      height=6, selectmode='extended')
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

        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, 
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

        self.convert_btn = ttk.Button(button_frame, text="▶️ Avvia Conversione Completa", 
                                     command=self.start_conversion)
        self.convert_btn.grid(row=0, column=0, padx=5)

        ttk.Button(button_frame, text="🗑️ Pulisci Log", 
                  command=self.clear_log).grid(row=0, column=1, padx=5)

        ttk.Button(button_frame, text="❌ Esci", 
                  command=self.root.quit).grid(row=0, column=2, padx=5)

        # Configura resize
        main_frame.rowconfigure(4, weight=1)
        main_frame.rowconfigure(5, weight=2)

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
        directory = filedialog.askdirectory(title="Seleziona cartella file WPS")
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

        # Controlla LibreOffice
        if self.libreoffice_path:
            self.log(f"✓ LibreOffice trovato: {self.libreoffice_path}")
        else:
            self.log("❌ LibreOffice NON installato")

        # Controlla Pandoc
        if self.pandoc_available:
            try:
                result = subprocess.run(['pandoc', '--version'], 
                                      capture_output=True, text=True, timeout=5)
                version = result.stdout.split('\n')[0]
                self.log(f"✓ {version}")
            except:
                pass
        else:
            self.log("❌ Pandoc NON installato")

        # Mostra avviso se manca qualcosa
        if not self.libreoffice_path or not self.pandoc_available:
            self.show_install_instructions()

    def show_install_instructions(self):
        """Mostra istruzioni installazione"""
        missing = []
        if not self.libreoffice_path:
            missing.append("LibreOffice")
        if not self.pandoc_available:
            missing.append("Pandoc")

        msg = f"""Dipendenze mancanti: {', '.join(missing)}

Installazione su macOS:

LibreOffice (per WPS → DOCX):
  brew install --cask libreoffice

Pandoc (per DOCX → Markdown):
  brew install pandoc

Dopo l'installazione, riavvia l'applicazione."""

        messagebox.showwarning("Dipendenze mancanti", msg)

    def scan_files(self):
        """Scansiona cartella per file WPS"""
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

        # Cerca file .wps
        path = Path(input_path)
        wps_files = list(path.glob('*.wps')) + list(path.glob('**/*.wps'))

        for file in wps_files:
            size = file.stat().st_size
            size_str = self.format_size(size)

            self.file_tree.insert('', tk.END, text=file.name, 
                                 values=(str(file.parent), size_str))
            self.files_found.append(file)

        self.log(f"✓ Trovati {len(self.files_found)} file WPS")

        if len(self.files_found) == 0:
            messagebox.showinfo("Nessun file", 
                              "Nessun file .wps trovato nella cartella selezionata")

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

        if not self.libreoffice_path:
            messagebox.showerror("Errore", "LibreOffice non installato. Impossibile procedere.")
            return

        if not self.pandoc_available:
            messagebox.showerror("Errore", "Pandoc non installato. Impossibile procedere.")
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
        """Esegue conversione completa file (WPS → DOCX → MD)"""
        output_path = Path(self.output_dir.get())
        temp_dir = output_path / "temp_docx"
        temp_dir.mkdir(exist_ok=True)
        media_dir = output_path / "media"
        media_dir.mkdir(exist_ok=True)

        total = len(self.files_found)
        success = 0
        failed = 0

        self.log("\n" + "="*60)
        self.log("🚀 Avvio conversione completa WPS → DOCX → Markdown")
        self.log("="*60)
        self.status_label.configure(text="Conversione in corso...")

        for idx, file in enumerate(self.files_found, 1):
            self.log(f"\n[{idx}/{total}] File: {file.name}")

            # Determina cartella di destinazione per DOCX
            if self.keep_docx.get():
                docx_dir = output_path
            else:
                docx_dir = temp_dir

            docx_file = docx_dir / f"{file.stem}.docx"
            md_file = output_path / f"{file.stem}.md"

            try:
                # FASE 1: WPS → DOCX con LibreOffice
                self.log(f"  [1/2] Conversione WPS → DOCX...")

                cmd_libreoffice = [
                    self.libreoffice_path,
                    '--headless',
                    '--convert-to', 'docx',
                    '--outdir', str(docx_dir),
                    str(file)
                ]

                result1 = subprocess.run(cmd_libreoffice, capture_output=True, 
                                        text=True, timeout=120)

                if result1.returncode != 0 or not docx_file.exists():
                    self.log(f"  ✗ Errore LibreOffice: {result1.stderr[:100]}")
                    failed += 1
                    continue

                self.log(f"  ✓ DOCX creato: {docx_file.name}")

                # FASE 2: DOCX → Markdown con Pandoc
                self.log(f"  [2/2] Conversione DOCX → Markdown...")

                cmd_pandoc = [
                    'pandoc',
                    str(docx_file),
                    '-o', str(md_file),
                    '--extract-media', str(media_dir),
                    '--wrap=none',
                    '--standalone'
                ]

                result2 = subprocess.run(cmd_pandoc, capture_output=True, 
                                        text=True, timeout=120)

                if result2.returncode == 0 and md_file.exists():
                    self.log(f"  ✓ Markdown creato: {md_file.name}")
                    success += 1

                    # Rimuovi DOCX temporaneo se richiesto
                    if not self.keep_docx.get():
                        docx_file.unlink()
                else:
                    self.log(f"  ✗ Errore Pandoc: {result2.stderr[:100]}")
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

        # Pulizia cartella temporanea
        if not self.keep_docx.get() and temp_dir.exists():
            try:
                temp_dir.rmdir()
            except:
                pass

        # Riepilogo finale
        self.log("\n" + "="*60)
        self.log("📊 RIEPILOGO CONVERSIONE COMPLETA")
        self.log("="*60)
        self.log(f"File totali:         {total}")
        self.log(f"Convertiti (MD):     {success} ✓")
        self.log(f"Falliti:             {failed} ✗")
        self.log(f"Percentuale:         {(success/total*100):.1f}%")
        self.log(f"Output Markdown:     {output_path}")
        if self.keep_docx.get():
            self.log(f"File DOCX:           Mantenuti in {output_path}")
        self.log(f"Immagini estratte:   {media_dir}")
        self.log("="*60)

        self.status_label.configure(text=f"Completato: {success}/{total} file convertiti")
        self.convert_btn.configure(state='normal')
        self.converting = False

        # Mostra messaggio finale
        msg = f"""Conversione completata!

File Markdown creati: {success}/{total}

Salvati in:
{output_path}

Immagini estratte in:
{media_dir}"""

        if self.keep_docx.get():
            msg += f"\n\nFile DOCX intermedi mantenuti in:\n{output_path}"

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
    app = WPStoMarkdownConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()

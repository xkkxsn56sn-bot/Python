#!/usr/bin/env python3
from __future__ import annotations

import subprocess
import threading
import tkinter as tk
from pathlib import Path
from queue import Empty, Queue
from tkinter import filedialog, messagebox, scrolledtext, ttk

from to_pdf_converter import SUPPORTED_EXTENSIONS, convert_all, discover_input_files


class ToPDFConverterGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Batch Converter to PDF")
        self.root.geometry("1180x900")
        self.root.minsize(1040, 820)
        self.root.resizable(True, True)

        self.input_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.resource_path = tk.StringVar()
        self.report_name = tk.StringVar(value="to-pdf-report.json")
        self.pdf_engine = tk.StringVar(value="xelatex")
        self.preset = tk.StringVar(value="balanced")
        self.recursive = tk.BooleanVar(value=True)
        self.overwrite = tk.BooleanVar(value=True)
        self.prefer_pandoc = tk.BooleanVar(value=True)
        self.toc = tk.BooleanVar(value=False)
        self.number_sections = tk.BooleanVar(value=False)

        self.file_paths: list[Path] = []
        self.last_report_path: Path | None = None
        self.last_failed_results: list[dict[str, object]] = []
        self.is_running = False
        self.progress_queue: Queue[tuple[str, object]] = Queue()

        self.setup_ui()
        self.root.after(100, self.process_queue)

    def setup_ui(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("aqua")
        except tk.TclError:
            pass

        default_label = tk.Label(self.root)
        self.default_text_color = default_label.cget("fg")
        default_label.destroy()
        self.default_background_color = self.root.cget("bg")

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main_frame = ttk.Frame(self.root, padding=12)
        main_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.columnconfigure(2, weight=1)
        main_frame.columnconfigure(3, weight=1)
        main_frame.rowconfigure(4, weight=1)

        title = ttk.Label(main_frame, text="Batch Converter to PDF", font=("Helvetica", 17, "bold"))
        title.grid(row=0, column=0, columnspan=4, sticky=tk.W, pady=(0, 6))

        subtitle = ttk.Label(
            main_frame,
            text="Formati supportati: " + ", ".join(sorted(SUPPORTED_EXTENSIONS)),
            foreground="gray",
        )
        subtitle.grid(row=1, column=0, columnspan=4, sticky=tk.W, pady=(0, 12))

        input_frame = ttk.LabelFrame(main_frame, text="Input", padding=10)
        input_frame.grid(row=2, column=0, columnspan=4, sticky=(tk.E, tk.W), pady=5)
        input_frame.columnconfigure(0, weight=1)

        ttk.Entry(input_frame, textvariable=self.input_path).grid(row=0, column=0, sticky=(tk.E, tk.W), padx=(0, 6))
        ttk.Button(input_frame, text="File", command=self.browse_file).grid(row=0, column=1)
        ttk.Button(input_frame, text="Cartella", command=self.browse_folder).grid(row=0, column=2, padx=(6, 0))
        ttk.Button(input_frame, text="Scansiona", command=self.scan_input).grid(row=0, column=3, padx=(6, 0))

        output_frame = ttk.LabelFrame(main_frame, text="Output", padding=10)
        output_frame.grid(row=3, column=0, columnspan=4, sticky=(tk.E, tk.W), pady=5)
        output_frame.columnconfigure(0, weight=1)

        ttk.Entry(output_frame, textvariable=self.output_dir).grid(row=0, column=0, sticky=(tk.E, tk.W), padx=(0, 6))
        ttk.Button(output_frame, text="Sfoglia", command=self.browse_output).grid(row=0, column=1)

        content_pane = ttk.Panedwindow(main_frame, orient=tk.VERTICAL)
        content_pane.grid(row=4, column=0, columnspan=4, sticky=(tk.N, tk.S, tk.E, tk.W), pady=5)

        top_frame = ttk.Frame(content_pane)
        bottom_frame = ttk.Frame(content_pane)
        top_frame.columnconfigure(0, weight=1)
        top_frame.rowconfigure(0, weight=1)
        bottom_frame.columnconfigure(0, weight=1)
        bottom_frame.rowconfigure(0, weight=1)

        content_pane.add(top_frame, weight=3)
        content_pane.add(bottom_frame, weight=2)

        top_pane = ttk.Panedwindow(top_frame, orient=tk.HORIZONTAL)
        top_pane.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

        options_container = ttk.Frame(top_pane)
        right_frame = ttk.Frame(top_pane)
        options_container.columnconfigure(0, weight=1)
        options_container.rowconfigure(0, weight=1)
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=3, minsize=260)
        right_frame.rowconfigure(1, weight=1, minsize=120)

        top_pane.add(options_container, weight=2)
        top_pane.add(right_frame, weight=3)

        options_frame = ttk.LabelFrame(options_container, text="Opzioni", padding=10)
        options_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), padx=(0, 6))
        options_frame.columnconfigure(1, weight=1)

        preset_frame = ttk.LabelFrame(options_frame, text="Preset", padding=8)
        preset_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.E, tk.W), pady=(0, 10))
        preset_frame.columnconfigure(0, weight=1)
        preset_frame.columnconfigure(1, weight=1)
        preset_frame.columnconfigure(2, weight=1)
        ttk.Radiobutton(
            preset_frame,
            text="Bilanciato",
            value="balanced",
            variable=self.preset,
            command=self.apply_selected_preset,
        ).grid(row=0, column=0, sticky=tk.W)
        ttk.Radiobutton(
            preset_frame,
            text="Qualita stampa",
            value="print_quality",
            variable=self.preset,
            command=self.apply_selected_preset,
        ).grid(row=0, column=1, sticky=tk.W, padx=(12, 0))
        ttk.Radiobutton(
            preset_frame,
            text="Conversione veloce",
            value="fast",
            variable=self.preset,
            command=self.apply_selected_preset,
        ).grid(row=0, column=2, sticky=tk.W, padx=(12, 0))

        self.preset_hint_container = tk.Frame(preset_frame, bg=self.default_background_color, highlightthickness=0, bd=0)
        self.preset_hint_container.grid(row=1, column=0, columnspan=3, sticky=(tk.E, tk.W), pady=(8, 0))
        self.preset_hint_container.grid_columnconfigure(0, weight=1)

        self.preset_hint = tk.Label(
            self.preset_hint_container,
            text="",
            fg=self.default_text_color,
            bg=self.default_background_color,
            font=("Helvetica", 11),
            justify=tk.LEFT,
            wraplength=520,
            anchor="w",
            padx=0,
            pady=0,
        )
        self.preset_hint.grid(row=0, column=0, sticky=(tk.E, tk.W))
        self.preset_hint_container.bind("<Configure>", self.on_preset_hint_resize)

        ttk.Checkbutton(options_frame, text="Ricorsivo", variable=self.recursive).grid(row=1, column=0, sticky=tk.W)
        ttk.Checkbutton(options_frame, text="Sovrascrivi PDF esistenti", variable=self.overwrite).grid(row=2, column=0, sticky=tk.W)
        ttk.Checkbutton(options_frame, text="Preferisci Pandoc per MD/EPUB", variable=self.prefer_pandoc).grid(row=3, column=0, sticky=tk.W)
        ttk.Checkbutton(options_frame, text="Indice (TOC)", variable=self.toc).grid(row=4, column=0, sticky=tk.W)
        ttk.Checkbutton(options_frame, text="Numerazione sezioni", variable=self.number_sections).grid(row=5, column=0, sticky=tk.W)

        ttk.Label(options_frame, text="PDF engine").grid(row=6, column=0, sticky=tk.W, pady=(10, 0))
        ttk.Entry(options_frame, textvariable=self.pdf_engine).grid(row=6, column=1, sticky=(tk.E, tk.W), pady=(10, 0))

        ttk.Label(options_frame, text="Resource path").grid(row=7, column=0, sticky=tk.W, pady=(8, 0))
        ttk.Entry(options_frame, textvariable=self.resource_path).grid(row=7, column=1, sticky=(tk.E, tk.W), pady=(8, 0))
        ttk.Button(options_frame, text="Sfoglia", command=self.browse_resource_path).grid(row=7, column=2, padx=(6, 0), pady=(8, 0))

        ttk.Label(options_frame, text="Report JSON").grid(row=8, column=0, sticky=tk.W, pady=(8, 0))
        ttk.Entry(options_frame, textvariable=self.report_name).grid(row=8, column=1, sticky=(tk.E, tk.W), pady=(8, 0))

        files_frame = ttk.LabelFrame(right_frame, text="File trovati", padding=10)
        files_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), pady=(0, 5))
        files_frame.columnconfigure(0, weight=1)
        files_frame.rowconfigure(0, weight=1)

        self.file_tree = ttk.Treeview(files_frame, columns=("folder", "type"), height=16)
        self.file_tree.heading("#0", text="Nome file")
        self.file_tree.heading("folder", text="Cartella")
        self.file_tree.heading("type", text="Tipo")
        self.file_tree.column("#0", width=220)
        self.file_tree.column("folder", width=330)
        self.file_tree.column("type", width=80, anchor=tk.CENTER)
        self.file_tree.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

        files_scroll = ttk.Scrollbar(files_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        files_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.file_tree.configure(yscrollcommand=files_scroll.set)

        failures_frame = ttk.LabelFrame(right_frame, text="File falliti", padding=10)
        failures_frame.grid(row=1, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), pady=(5, 0))
        failures_frame.columnconfigure(0, weight=1)
        failures_frame.rowconfigure(0, weight=1)

        self.failure_tree = ttk.Treeview(failures_frame, columns=("type", "reason"), height=5)
        self.failure_tree.heading("#0", text="File")
        self.failure_tree.heading("type", text="Tipo")
        self.failure_tree.heading("reason", text="Errore")
        self.failure_tree.column("#0", width=220)
        self.failure_tree.column("type", width=80, anchor=tk.CENTER)
        self.failure_tree.column("reason", width=330)
        self.failure_tree.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

        failures_scroll = ttk.Scrollbar(failures_frame, orient=tk.VERTICAL, command=self.failure_tree.yview)
        failures_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.failure_tree.configure(yscrollcommand=failures_scroll.set)

        log_frame = ttk.LabelFrame(bottom_frame, text="Log conversione", padding=10)
        log_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, state="disabled", wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=5, column=0, columnspan=4, sticky=(tk.E, tk.W), pady=(8, 0))
        progress_frame.columnconfigure(0, weight=1)

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.E, tk.W))

        self.status_label = ttk.Label(progress_frame, text="Pronto")
        self.status_label.grid(row=1, column=0, sticky=tk.W, pady=(4, 0))

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=4, pady=(12, 0), sticky=tk.W)

        self.start_button = ttk.Button(button_frame, text="Avvia conversione", command=self.start_conversion)
        self.start_button.grid(row=0, column=0)
        ttk.Button(button_frame, text="Apri output", command=self.open_output_folder).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(button_frame, text="Apri report", command=self.open_report_file).grid(row=0, column=2, padx=(8, 0))
        ttk.Button(button_frame, text="Pulisci log", command=self.clear_log).grid(row=0, column=3, padx=(8, 0))
        ttk.Button(button_frame, text="Esci", command=self.root.quit).grid(row=0, column=4, padx=(8, 0))

        # Keep the options and files sections visible by default while allowing manual resize.
        self.root.after(200, lambda: content_pane.sashpos(0, 460))
        self.root.after(250, lambda: top_pane.sashpos(0, 470))

        self.apply_selected_preset()
        self.log("Pronto. Seleziona un file o una cartella e avvia la scansione.")

    def clear_failure_list(self) -> None:
        for item in self.failure_tree.get_children():
            self.failure_tree.delete(item)

    def on_preset_hint_resize(self, event: tk.Event) -> None:
        wraplength = max(event.width - 4, 220)
        self.preset_hint.config(wraplength=wraplength)

    def populate_failure_list(self) -> None:
        self.clear_failure_list()
        for result in self.last_failed_results:
            source = Path(str(result.get("source", "")))
            suffix = source.suffix.lower()
            stderr = str(result.get("stderr", "")).strip()
            stdout = str(result.get("stdout", "")).strip()
            reason = stderr or stdout or "Errore sconosciuto"
            reason = " ".join(reason.splitlines())
            if len(reason) > 160:
                reason = reason[:157] + "..."

            self.failure_tree.insert(
                "",
                tk.END,
                text=source.name or str(result.get("source", "")),
                values=(suffix, reason),
            )

    def log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")

    def browse_file(self) -> None:
        filename = filedialog.askopenfilename(
            title="Seleziona un file",
            filetypes=[("Documenti supportati", "*.docx *.wps *.epub *.txt *.rtf *.md")],
        )
        if filename:
            self.input_path.set(filename)
            self.log(f"File selezionato: {filename}")

    def browse_folder(self) -> None:
        folder = filedialog.askdirectory(title="Seleziona cartella input")
        if folder:
            self.input_path.set(folder)
            self.log(f"Cartella selezionata: {folder}")

    def browse_output(self) -> None:
        folder = filedialog.askdirectory(title="Seleziona cartella output")
        if folder:
            self.output_dir.set(folder)
            self.log(f"Cartella output: {folder}")

    def browse_resource_path(self) -> None:
        folder = filedialog.askdirectory(title="Seleziona resource path per Pandoc")
        if folder:
            self.resource_path.set(folder)

    def apply_selected_preset(self) -> None:
        preset = self.preset.get()

        if preset == "print_quality":
            self.prefer_pandoc.set(True)
            self.pdf_engine.set("xelatex")
            self.toc.set(False)
            self.number_sections.set(False)
            self.preset_hint.config(
                text=(
                    "Qualita stampa: privilegia Pandoc con xelatex\n"
                    "per una resa tipografica piu stabile."
                )
            )
            return

        if preset == "fast":
            self.prefer_pandoc.set(False)
            self.pdf_engine.set("xelatex")
            self.toc.set(False)
            self.number_sections.set(False)
            self.preset_hint.config(
                text=(
                    "Conversione veloce: usa LibreOffice come prima scelta\n"
                    "quando possibile e mantiene opzioni minime."
                )
            )
            return

        self.prefer_pandoc.set(True)
        self.pdf_engine.set("xelatex")
        self.toc.set(False)
        self.number_sections.set(False)
        self.preset_hint.config(text="Bilanciato: impostazione consigliata\nper uso generale.")

    def open_output_folder(self) -> None:
        output_value = self.output_dir.get().strip()
        if not output_value:
            messagebox.showwarning("Attenzione", "Seleziona prima una cartella di output")
            return

        output_path = Path(output_value).expanduser()
        if not output_path.exists():
            messagebox.showwarning("Attenzione", f"Cartella output non trovata: {output_path}")
            return

        try:
            subprocess.run(["open", str(output_path)], check=True)
            self.log(f"Aperta cartella output: {output_path}")
        except Exception as exc:
            messagebox.showerror("Errore", f"Impossibile aprire la cartella output: {exc}")
            self.log(f"Errore apertura output: {exc}")

    def open_report_file(self) -> None:
        if self.last_report_path is None:
            messagebox.showwarning("Attenzione", "Nessun report disponibile. Esegui prima una conversione.")
            return

        if not self.last_report_path.exists():
            messagebox.showwarning("Attenzione", f"Report non trovato: {self.last_report_path}")
            return

        try:
            subprocess.run(["open", str(self.last_report_path)], check=True)
            self.log(f"Aperto report: {self.last_report_path}")
        except Exception as exc:
            messagebox.showerror("Errore", f"Impossibile aprire il report: {exc}")
            self.log(f"Errore apertura report: {exc}")

    def scan_input(self) -> None:
        path_value = self.input_path.get().strip()
        if not path_value:
            messagebox.showwarning("Attenzione", "Seleziona prima un file o una cartella input")
            return

        try:
            file_paths = discover_input_files(Path(path_value), recursive=self.recursive.get())
        except Exception as exc:
            messagebox.showerror("Errore", str(exc))
            self.log(f"Errore scansione: {exc}")
            return

        self.file_paths = file_paths
        self.last_failed_results = []
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        self.clear_failure_list()

        for file_path in file_paths:
            self.file_tree.insert(
                "",
                tk.END,
                text=file_path.name,
                values=(str(file_path.parent), file_path.suffix.lower()),
            )

        self.status_label.config(text=f"Trovati {len(file_paths)} file")
        self.log(f"Scansione completata: {len(file_paths)} file compatibili")

    def validate(self) -> bool:
        if self.is_running:
            return False

        if not self.input_path.get().strip():
            messagebox.showwarning("Attenzione", "Seleziona un input")
            return False

        if not self.output_dir.get().strip():
            messagebox.showwarning("Attenzione", "Seleziona una cartella di output")
            return False

        if not self.report_name.get().strip():
            messagebox.showwarning("Attenzione", "Inserisci un nome report JSON")
            return False

        if not self.file_paths:
            self.scan_input()
            if not self.file_paths:
                return False

        return True

    def start_conversion(self) -> None:
        if not self.validate():
            return

        self.is_running = True
        self.last_report_path = None
        self.last_failed_results = []
        self.clear_failure_list()
        self.start_button.config(state=tk.DISABLED)
        self.progress_var.set(0)
        self.status_label.config(text="Conversione in corso...")
        self.log("Avvio conversione batch...")

        worker = threading.Thread(target=self.run_conversion, daemon=True)
        worker.start()

    def progress_callback(self, current: int, total: int, message: str) -> None:
        self.progress_queue.put(("progress", (current, total, message)))

    def run_conversion(self) -> None:
        try:
            report = convert_all(
                input_path=Path(self.input_path.get().strip()),
                output_dir=Path(self.output_dir.get().strip()),
                recursive=self.recursive.get(),
                overwrite=self.overwrite.get(),
                prefer_pandoc=self.prefer_pandoc.get(),
                pdf_engine=self.pdf_engine.get().strip() or "xelatex",
                toc=self.toc.get(),
                number_sections=self.number_sections.get(),
                resource_path=Path(self.resource_path.get().strip()) if self.resource_path.get().strip() else None,
                report_name=self.report_name.get().strip(),
                progress_callback=self.progress_callback,
            )
            self.progress_queue.put(("done", report))
        except Exception as exc:
            self.progress_queue.put(("error", str(exc)))

    def process_queue(self) -> None:
        try:
            while True:
                event, payload = self.progress_queue.get_nowait()
                if event == "progress":
                    current, total, message = payload
                    percent = 100 if total == 0 else (current / total) * 100
                    self.progress_var.set(percent)
                    self.status_label.config(text=message)
                    self.log(message)
                elif event == "done":
                    report = payload
                    summary = report["summary"]
                    self.last_report_path = Path(report["output_dir"]) / self.report_name.get().strip()
                    self.last_failed_results = [result for result in report.get("results", []) if result.get("status") == "failed"]
                    self.populate_failure_list()
                    self.progress_var.set(100)
                    self.status_label.config(
                        text=f"Completato: {summary['converted']} convertiti, {summary['skipped']} saltati, {summary['failed']} falliti"
                    )
                    self.log(
                        "Conversione completata: "
                        f"{summary['converted']} convertiti, {summary['skipped']} saltati, {summary['failed']} falliti"
                    )
                    self.log(f"Report JSON: {self.last_report_path}")
                    if self.last_failed_results:
                        self.log("Dettagli errori disponibili nella tabella 'File falliti'.")
                    messagebox.showinfo(
                        "Completato",
                        "Conversione completata.\n\n"
                        f"Convertiti: {summary['converted']}\n"
                        f"Saltati: {summary['skipped']}\n"
                        f"Falliti: {summary['failed']}\n\n"
                        f"Report: {self.last_report_path}",
                    )
                    self.is_running = False
                    self.start_button.config(state=tk.NORMAL)
                elif event == "error":
                    self.progress_var.set(0)
                    self.status_label.config(text="Errore")
                    self.log(f"Errore: {payload}")
                    messagebox.showerror("Errore", str(payload))
                    self.is_running = False
                    self.start_button.config(state=tk.NORMAL)
        except Empty:
            pass

        self.root.after(100, self.process_queue)


def main() -> None:
    root = tk.Tk()
    app = ToPDFConverterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
#!/usr/bin/env python3
from __future__ import annotations

import shlex
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, scrolledtext, ttk


ROOT_DIR = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = ROOT_DIR / "scripts"


class ScriptLauncherGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Launcher GUI Python Scripts")
        self.root.geometry("980x720")
        self.root.minsize(900, 620)

        self.script_paths: list[Path] = []
        self.setup_ui()
        self.refresh_scripts()

    def setup_ui(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("aqua")
        except tk.TclError:
            pass

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        main_frame = ttk.Frame(self.root, padding=12)
        main_frame.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)

        title = ttk.Label(
            main_frame,
            text="Launcher GUI per script Python",
            font=("Helvetica", 16, "bold"),
        )
        title.grid(row=0, column=0, sticky=tk.W, pady=(0, 10))

        subtitle = ttk.Label(
            main_frame,
            text="Avvia gli script dalla cartella scripts con un’unica interfaccia.",
            foreground="gray",
        )
        subtitle.grid(row=1, column=0, sticky=tk.W, pady=(0, 12))

        toolbar = ttk.Frame(main_frame)
        toolbar.grid(row=2, column=0, sticky=(tk.E, tk.W), pady=(0, 8))
        toolbar.columnconfigure(1, weight=1)

        ttk.Label(toolbar, text="Cerca:").grid(row=0, column=0, padx=(0, 6))
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(toolbar, textvariable=self.search_var)
        search_entry.grid(row=0, column=1, sticky=(tk.E, tk.W))
        search_entry.bind("<KeyRelease>", lambda _event: self.refresh_scripts())

        ttk.Button(toolbar, text="Aggiorna", command=self.refresh_scripts).grid(row=0, column=2, padx=(8, 0))

        content = ttk.Frame(main_frame)
        content.grid(row=3, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))
        content.columnconfigure(0, weight=1)
        content.rowconfigure(0, weight=1)

        self.script_tree = ttk.Treeview(content, columns=("path",), height=10)
        self.script_tree.heading("#0", text="Script")
        self.script_tree.heading("path", text="Percorso")
        self.script_tree.column("#0", width=260)
        self.script_tree.column("path", width=520)
        self.script_tree.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

        tree_scroll = ttk.Scrollbar(content, orient=tk.VERTICAL, command=self.script_tree.yview)
        tree_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.script_tree.configure(yscrollcommand=tree_scroll.set)
        self.script_tree.bind("<<TreeviewSelect>>", self.on_script_selected)
        self.script_tree.bind("<Double-1>", lambda _event: self.run_selected_script())

        details_frame = ttk.LabelFrame(content, text="Dettagli script", padding=8)
        details_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.E, tk.W), pady=(8, 0))
        details_frame.columnconfigure(0, weight=1)

        self.details_title_var = tk.StringVar(value="Seleziona uno script")
        self.details_path_var = tk.StringVar(value="")
        ttk.Label(details_frame, textvariable=self.details_title_var, font=("Helvetica", 12, "bold")).grid(
            row=0, column=0, sticky=tk.W
        )
        ttk.Label(details_frame, textvariable=self.details_path_var, foreground="gray", wraplength=700).grid(
            row=1, column=0, sticky=tk.W, pady=(4, 0)
        )

        footer = ttk.Frame(main_frame)
        footer.grid(row=4, column=0, sticky=(tk.E, tk.W), pady=(8, 0))
        footer.columnconfigure(0, weight=1)

        ttk.Label(footer, text="Argomenti:").grid(row=0, column=0, sticky=tk.W)
        self.args_var = tk.StringVar()
        ttk.Entry(footer, textvariable=self.args_var).grid(row=0, column=1, sticky=(tk.E, tk.W), padx=(6, 8))

        self.run_button = ttk.Button(footer, text="Avvia script", command=self.run_selected_script)
        self.run_button.grid(row=0, column=2)

        log_frame = ttk.LabelFrame(main_frame, text="Log", padding=8)
        log_frame.grid(row=5, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), pady=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, state="disabled", wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

        self.log("Launcher pronto. Seleziona uno script e avvialo.")

    def refresh_scripts(self) -> None:
        query = self.search_var.get().strip().lower()
        for item in self.script_tree.get_children():
            self.script_tree.delete(item)

        self.script_paths = []
        if not SCRIPTS_DIR.exists():
            self.log("Cartella scripts non trovata")
            return

        for script_path in sorted(SCRIPTS_DIR.glob("*.py")):
            if script_path.name in {"gui_launcher.py", "__init__.py"}:
                continue
            display_name = script_path.stem
            relative_path = script_path.relative_to(ROOT_DIR)
            if query and query not in display_name.lower() and query not in str(relative_path).lower():
                continue
            self.script_tree.insert("", tk.END, text=display_name, values=(str(relative_path),))
            self.script_paths.append(script_path)

        if not self.script_paths:
            self.log("Nessuno script disponibile nella cartella scripts")

    def get_selected_script(self) -> Path | None:
        selected_items = self.script_tree.selection()
        if not selected_items:
            return None
        index = self.script_tree.index(selected_items[0])
        if 0 <= index < len(self.script_paths):
            return self.script_paths[index]
        return None

    def on_script_selected(self, _event: tk.Event | None = None) -> None:
        script_path = self.get_selected_script()
        if script_path is None:
            self.details_title_var.set("Seleziona uno script")
            self.details_path_var.set("")
            return

        self.details_title_var.set(script_path.stem)
        self.details_path_var.set(str(script_path.relative_to(ROOT_DIR)))

    def run_selected_script(self) -> None:
        script_path = self.get_selected_script()
        if script_path is None:
            messagebox.showwarning("Selezione richiesta", "Seleziona prima uno script dalla lista")
            return

        args_text = self.args_var.get().strip()
        args = shlex.split(args_text) if args_text else []
        command = [sys.executable, str(script_path), *args]
        self.log(f"Avvio: {' '.join(command)}")

        try:
            subprocess.Popen(command, cwd=str(ROOT_DIR))
            self.log("Script avviato in background")
        except OSError as exc:
            self.log(f"Errore nell'avvio: {exc}")
            messagebox.showerror("Errore", f"Impossibile avviare lo script:\n{exc}")

    def log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")


def main() -> None:
    root = tk.Tk()
    ScriptLauncherGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
# Run with: .venv/bin/python scripts/zip_compressor_gui.py
import os
import threading
import zipfile
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


class ZipCompressorGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("ZIP Compressor")
        self.root.geometry("800x620")
        self.root.minsize(640, 500)
        self.root.resizable(True, True)

        self.selected_files: list[Path] = []
        self.output_path = tk.StringVar()
        self.compression_level = tk.IntVar(value=6)
        self.is_running = False

        self.setup_ui()

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
        main_frame.rowconfigure(2, weight=1)

        # ===== HEADER =====
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, sticky=tk.EW, pady=(0, 10))

        ttk.Label(header_frame, text="ZIP Compressor", font=("Arial", 16, "bold")).pack(anchor=tk.W)
        ttk.Label(
            header_frame,
            text="Select files or folders to compress into a ZIP archive.",
            font=("Arial", 10),
            foreground="gray",
        ).pack(anchor=tk.W, pady=(4, 0))

        # ===== FILE SELECTION =====
        files_frame = ttk.LabelFrame(main_frame, text="Files & Folders to Compress", padding=10)
        files_frame.grid(row=1, column=0, sticky=tk.EW, pady=(0, 10))
        files_frame.columnconfigure(0, weight=1)

        btn_row = ttk.Frame(files_frame)
        btn_row.grid(row=0, column=0, sticky=tk.EW, pady=(0, 8))

        ttk.Button(btn_row, text="Add Files", command=self.add_files).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="Add Folder", command=self.add_folder).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="Remove Selected", command=self.remove_selected).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_row, text="Clear All", command=self.clear_all).pack(side=tk.LEFT)

        list_frame = ttk.Frame(files_frame)
        list_frame.grid(row=1, column=0, sticky=tk.EW)
        list_frame.columnconfigure(0, weight=1)

        self.file_listbox = tk.Listbox(
            list_frame,
            selectmode=tk.EXTENDED,
            height=8,
            font=("Menlo", 10),
            relief=tk.FLAT,
            borderwidth=1,
            highlightthickness=1,
        )
        self.file_listbox.grid(row=0, column=0, sticky=tk.EW)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky=tk.NS)
        self.file_listbox.configure(yscrollcommand=scrollbar.set)

        # ===== OUTPUT & OPTIONS =====
        options_frame = ttk.LabelFrame(main_frame, text="Output & Options", padding=10)
        options_frame.grid(row=2, column=0, sticky=tk.EW, pady=(0, 10))
        options_frame.columnconfigure(1, weight=1)

        ttk.Label(options_frame, text="Output ZIP file:").grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        ttk.Entry(options_frame, textvariable=self.output_path, font=("Arial", 10)).grid(
            row=0, column=1, sticky=tk.EW, padx=(0, 8)
        )
        ttk.Button(options_frame, text="Browse", command=self.browse_output).grid(row=0, column=2, sticky=tk.W)

        ttk.Label(options_frame, text="Compression level:").grid(row=1, column=0, sticky=tk.W, padx=(0, 8), pady=(10, 0))
        level_frame = ttk.Frame(options_frame)
        level_frame.grid(row=1, column=1, sticky=tk.W, pady=(10, 0))

        ttk.Scale(
            level_frame,
            from_=0,
            to=9,
            orient=tk.HORIZONTAL,
            variable=self.compression_level,
            length=200,
            command=lambda v: self.compression_level.set(int(float(v))),
        ).pack(side=tk.LEFT)
        self.level_label = ttk.Label(level_frame, text="6  (balanced)", width=18)
        self.level_label.pack(side=tk.LEFT, padx=(10, 0))
        self.compression_level.trace_add("write", self._update_level_label)

        # ===== LOG =====
        log_frame = ttk.LabelFrame(main_frame, text="Log", padding=10)
        log_frame.grid(row=3, column=0, sticky=(tk.N, tk.S, tk.E, tk.W), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)

        self.log = scrolledtext.ScrolledText(
            log_frame, height=8, state=tk.DISABLED, font=("Menlo", 10), relief=tk.FLAT
        )
        self.log.grid(row=0, column=0, sticky=(tk.N, tk.S, tk.E, tk.W))

        # ===== ACTIONS =====
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=4, column=0, sticky=tk.EW)

        self.progress = ttk.Progressbar(action_frame, mode="indeterminate")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 12))

        self.compress_btn = ttk.Button(action_frame, text="Compress", command=self.start_compression)
        self.compress_btn.pack(side=tk.LEFT)

    # ------------------------------------------------------------------
    # UI helpers
    # ------------------------------------------------------------------

    def _update_level_label(self, *_) -> None:
        level = self.compression_level.get()
        descriptions = {
            0: "no compression",
            1: "fastest",
            2: "fast",
            3: "fast",
            4: "moderate",
            5: "moderate",
            6: "balanced",
            7: "good",
            8: "high",
            9: "maximum",
        }
        self.level_label.configure(text=f"{level}  ({descriptions[level]})")

    def _log(self, message: str) -> None:
        self.log.configure(state=tk.NORMAL)
        self.log.insert(tk.END, message + "\n")
        self.log.see(tk.END)
        self.log.configure(state=tk.DISABLED)

    def _refresh_listbox(self) -> None:
        self.file_listbox.delete(0, tk.END)
        for path in self.selected_files:
            self.file_listbox.insert(tk.END, str(path))

    # ------------------------------------------------------------------
    # File management
    # ------------------------------------------------------------------

    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(title="Select Files to Compress")
        for p in paths:
            path = Path(p)
            if path not in self.selected_files:
                self.selected_files.append(path)
        self._refresh_listbox()

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select Folder to Compress")
        if folder:
            path = Path(folder)
            if path not in self.selected_files:
                self.selected_files.append(path)
            self._refresh_listbox()

    def remove_selected(self) -> None:
        indices = list(self.file_listbox.curselection())
        for i in reversed(indices):
            self.selected_files.pop(i)
        self._refresh_listbox()

    def clear_all(self) -> None:
        self.selected_files.clear()
        self._refresh_listbox()

    def browse_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save ZIP File As",
            defaultextension=".zip",
            filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")],
        )
        if path:
            self.output_path.set(path)

    # ------------------------------------------------------------------
    # Compression
    # ------------------------------------------------------------------

    def start_compression(self) -> None:
        if self.is_running:
            return

        if not self.selected_files:
            messagebox.showwarning("No files selected", "Please add at least one file or folder to compress.")
            return

        output = self.output_path.get().strip()
        if not output:
            messagebox.showwarning("No output path", "Please specify a destination ZIP file.")
            return

        self.is_running = True
        self.compress_btn.configure(state=tk.DISABLED)
        self.progress.start(10)

        thread = threading.Thread(target=self._compress_worker, args=(output,), daemon=True)
        thread.start()

    def _compress_worker(self, output: str) -> None:
        output_path = Path(output)
        level = self.compression_level.get()

        try:
            self._log(f"Starting compression → {output_path.name}")
            self._log(f"Compression level: {level}")

            with zipfile.ZipFile(
                output_path,
                mode="w",
                compression=zipfile.ZIP_DEFLATED,
                compresslevel=level,
            ) as zf:
                total_added = 0
                for entry in self.selected_files:
                    if entry.is_dir():
                        for root, _, files in os.walk(entry):
                            for filename in files:
                                file_path = Path(root) / filename
                                arcname = file_path.relative_to(entry.parent)
                                zf.write(file_path, arcname)
                                self._log(f"  + {arcname}")
                                total_added += 1
                    else:
                        zf.write(entry, entry.name)
                        self._log(f"  + {entry.name}")
                        total_added += 1

            size_mb = output_path.stat().st_size / (1024 * 1024)
            self._log(f"\nDone! {total_added} item(s) compressed.")
            self._log(f"Output size: {size_mb:.2f} MB → {output_path}")
            self.root.after(0, lambda: messagebox.showinfo("Success", f"ZIP created successfully!\n{output_path}"))

        except Exception as exc:
            self._log(f"\nError: {exc}")
            self.root.after(0, lambda: messagebox.showerror("Error", str(exc)))

        finally:
            self.root.after(0, self._compression_finished)

    def _compression_finished(self) -> None:
        self.is_running = False
        self.compress_btn.configure(state=tk.NORMAL)
        self.progress.stop()


def main() -> None:
    root = tk.Tk()
    app = ZipCompressorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

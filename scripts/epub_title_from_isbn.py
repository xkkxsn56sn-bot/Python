#!/usr/bin/env python3
from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except Exception:
    tk = None
    filedialog = None
    messagebox = None
    ttk = None


CONTAINER_PATH = "META-INF/container.xml"


def get_opf_path(epub_path: Path) -> str:
    with zipfile.ZipFile(epub_path, "r") as archive:
        container_data = archive.read(CONTAINER_PATH)

    root = ET.fromstring(container_data)
    ns = {"c": "urn:oasis:names:tc:opendocument:xmlns:container"}
    node = root.find("c:rootfiles/c:rootfile", ns)
    if node is None:
        raise ValueError("File EPUB non valido: impossibile trovare il package document (OPF)")

    opf_path = node.attrib.get("full-path")
    if not opf_path:
        raise ValueError("File EPUB non valido: full-path mancante in container.xml")

    return opf_path


def extract_epub_title(epub_path: Path) -> str:
    opf_path = get_opf_path(epub_path)
    with zipfile.ZipFile(epub_path, "r") as archive:
        opf_data = archive.read(opf_path)

    root = ET.fromstring(opf_data)
    namespaces = {
        "dc": "http://purl.org/dc/elements/1.1/",
        "opf": "http://www.idpf.org/2007/opf",
    }

    title_node = root.find(".//dc:title", namespaces)
    if title_node is None or not (title_node.text and title_node.text.strip()):
        title_node = root.find(".//opf:metadata/dc:title", namespaces)

    if title_node is None or not (title_node.text and title_node.text.strip()):
        raise ValueError("Titolo non trovato nei metadati EPUB")

    return title_node.text.strip()


def sanitize_filename(value: str, fallback: str = "senza_titolo") -> str:
    cleaned = value.strip()
    cleaned = re.sub(r"[\\/:*?\"<>|]", "_", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    cleaned = cleaned.rstrip(".")
    return cleaned or fallback


def ensure_unique_path(path: Path) -> Path:
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    index = 2

    while True:
        candidate = parent / f"{stem} ({index}){suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def process_epub_directory(input_dir: Path, destination_name: str, recursive: bool = True) -> dict:
    input_dir = input_dir.expanduser().resolve()
    destination_name = destination_name.strip()

    if not destination_name:
        raise ValueError("Il nome della cartella di destinazione non puo essere vuoto")
    if "/" in destination_name or "\\" in destination_name:
        raise ValueError("Inserisci solo il nome cartella, senza percorso")

    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"Cartella input non trovata: {input_dir}")

    output_dir = (input_dir / destination_name).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    iterator = input_dir.rglob("*") if recursive else input_dir.glob("*")
    files = sorted(
        path
        for path in iterator
        if path.is_file()
        and path.suffix.lower() == ".epub"
        and output_dir not in path.parents
    )

    results: list[dict] = []
    for source in files:
        try:
            title = extract_epub_title(source)
            safe_title = sanitize_filename(title)
            relative_parent = source.relative_to(input_dir).parent
            destination_parent = output_dir / relative_parent
            destination_parent.mkdir(parents=True, exist_ok=True)
            destination = ensure_unique_path(destination_parent / f"{safe_title}.epub")
            shutil.copy2(source, destination)
            results.append(
                {
                    "source": str(source),
                    "title": title,
                    "saved_as": str(destination),
                    "status": "ok",
                }
            )
        except Exception as exc:
            results.append(
                {
                    "source": str(source),
                    "status": "error",
                    "error": str(exc),
                }
            )

    return {
        "input_dir": str(input_dir),
        "output_dir": str(output_dir),
        "total_found": len(files),
        "renamed": sum(1 for item in results if item["status"] == "ok"),
        "failed": sum(1 for item in results if item["status"] == "error"),
        "results": results,
    }


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("EPUB Renamer da Titolo")
        self.root.geometry("860x540")

        self.last_output_dir: Path | None = None

        frame = ttk.Frame(root, padding=12)
        frame.pack(fill="both", expand=True)

        self.base_dir_var = tk.StringVar(value=str(Path.cwd()))
        ttk.Label(frame, text="Cartella sorgente").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(0, 6))
        ttk.Entry(frame, textvariable=self.base_dir_var, width=60).grid(row=0, column=1, sticky="ew", pady=(0, 6))
        ttk.Button(frame, text="Sfoglia", command=self.choose_source_folder).grid(row=0, column=2, sticky="ew", padx=(8, 0), pady=(0, 6))

        self.destination_name_var = tk.StringVar(value="epub_rinominati")
        ttk.Label(frame, text="Nome cartella destinazione").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(0, 6))
        ttk.Entry(frame, textvariable=self.destination_name_var, width=60).grid(row=1, column=1, columnspan=2, sticky="ew", pady=(0, 6))

        self.recursive_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="Includi sottocartelle", variable=self.recursive_var).grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(2, 8)
        )

        ttk.Button(frame, text="Estrai titolo e rinomina", command=self.run_lookup).grid(
            row=3, column=0, columnspan=2, sticky="ew", pady=(4, 10)
        )
        self.open_output_button = ttk.Button(
            frame,
            text="Apri cartella output",
            command=self.open_output_folder,
            state="disabled",
        )
        self.open_output_button.grid(row=3, column=2, sticky="ew", padx=(8, 0), pady=(4, 10))

        self.output = tk.Text(frame, height=16, wrap="word")
        self.output.grid(row=4, column=0, columnspan=3, sticky="nsew")

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(4, weight=1)

    def choose_source_folder(self) -> None:
        selected = filedialog.askdirectory(title="Seleziona cartella sorgente EPUB")
        if selected:
            self.base_dir_var.set(selected)

    def run_lookup(self) -> None:
        destination_name = self.destination_name_var.get().strip()
        source_dir_text = self.base_dir_var.get().strip()

        if not destination_name:
            messagebox.showerror("Errore", "Inserisci il nome della cartella di destinazione")
            return
        if not source_dir_text:
            messagebox.showerror("Errore", "Seleziona la cartella sorgente")
            return

        try:
            result = process_epub_directory(
                input_dir=Path(source_dir_text),
                destination_name=destination_name,
                recursive=self.recursive_var.get(),
            )
            self.last_output_dir = Path(result["output_dir"])
            self.open_output_button.configure(state="normal")
            self.output.delete("1.0", "end")
            self.output.insert("1.0", json.dumps(result, ensure_ascii=False, indent=2))
        except Exception as exc:
            messagebox.showerror("Errore", str(exc))

    def open_output_folder(self) -> None:
        if self.last_output_dir is None:
            messagebox.showerror("Errore", "Cartella output non disponibile")
            return

        try:
            self.last_output_dir.mkdir(parents=True, exist_ok=True)
            subprocess.run(["open", str(self.last_output_dir)], check=False)
        except Exception as exc:
            messagebox.showerror("Errore", f"Impossibile aprire la cartella output: {exc}")


def launch_gui() -> int:
    if tk is None:
        print("Tkinter non disponibile in questo ambiente", file=sys.stderr)
        return 1

    root = tk.Tk()
    App(root)
    root.mainloop()
    return 0


def main() -> int:
    return launch_gui()


if __name__ == "__main__":
    raise SystemExit(main())

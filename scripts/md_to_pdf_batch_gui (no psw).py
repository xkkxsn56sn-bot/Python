#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import threading
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
from pathlib import Path
from queue import Queue, Empty
from typing import Callable, Optional

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except Exception:
    tk = None
    filedialog = None
    messagebox = None
    ttk = None


@dataclass
class ConversionResult:
    source: str
    target: str
    status: str
    returncode: Optional[int] = None
    stderr: str = ""
    stdout: str = ""


ProgressCallback = Optional[Callable[[int, int, str], None]]


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def check_dependency(name: str) -> bool:
    return shutil.which(name) is not None


def find_markdown_files(input_dir: Path, recursive: bool) -> list[Path]:
    patterns = ("*.md", "*.markdown")
    files: list[Path] = []
    for pattern in patterns:
        files.extend(input_dir.rglob(pattern) if recursive else input_dir.glob(pattern))
    return sorted({p.resolve() for p in files})


def build_command(
    src: Path,
    dst: Path,
    input_format: str,
    pdf_engine: str,
    toc: bool,
    number_sections: bool,
    metadata_file: Optional[Path],
    resource_path: Optional[Path],
) -> list[str]:
    cmd = [
        "pandoc",
        str(src),
        "-f",
        input_format,
        "-s",
        "-o",
        str(dst),
        "--pdf-engine",
        pdf_engine,
    ]
    if toc:
        cmd.append("--toc")
    if number_sections:
        cmd.append("--number-sections")
    if metadata_file:
        cmd.extend(["--metadata-file", str(metadata_file)])
    if resource_path:
        cmd.extend(["--resource-path", str(resource_path)])
    return cmd


def convert_all(
    input_dir: Path,
    output_dir: Path,
    input_format: str,
    pdf_engine: str,
    overwrite: bool,
    recursive: bool,
    toc: bool,
    number_sections: bool,
    metadata_file: Optional[Path],
    resource_path: Optional[Path],
    report_name: str,
    progress_callback: ProgressCallback = None,
) -> dict:
    input_dir = input_dir.resolve()
    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"Input directory not found: {input_dir}")

    if not check_dependency("pandoc"):
        raise RuntimeError("pandoc is not installed or not available in PATH")

    if not check_dependency(pdf_engine):
        raise RuntimeError(f"PDF engine '{pdf_engine}' is not installed or not available in PATH")

    md_files = find_markdown_files(input_dir, recursive)
    results: list[ConversionResult] = []
    total = len(md_files)

    for index, src in enumerate(md_files, start=1):
        rel = src.relative_to(input_dir)
        dst = (output_dir / rel).with_suffix(".pdf")
        dst.parent.mkdir(parents=True, exist_ok=True)

        if progress_callback:
            progress_callback(index - 1, total, f"Preparing {src.name}")

        if dst.exists() and not overwrite:
            results.append(ConversionResult(str(src), str(dst), "skipped"))
            if progress_callback:
                progress_callback(index, total, f"Skipped {src.name}")
            continue

        cmd = build_command(
            src=src,
            dst=dst,
            input_format=input_format,
            pdf_engine=pdf_engine,
            toc=toc,
            number_sections=number_sections,
            metadata_file=metadata_file,
            resource_path=resource_path,
        )

        run = subprocess.run(cmd, capture_output=True, text=True)

        if run.returncode == 0:
            results.append(
                ConversionResult(
                    source=str(src),
                    target=str(dst),
                    status="converted",
                    returncode=0,
                    stderr=run.stderr.strip(),
                    stdout=run.stdout.strip(),
                )
            )
            if progress_callback:
                progress_callback(index, total, f"Converted {src.name}")
        else:
            results.append(
                ConversionResult(
                    source=str(src),
                    target=str(dst),
                    status="failed",
                    returncode=run.returncode,
                    stderr=run.stderr.strip(),
                    stdout=run.stdout.strip(),
                )
            )
            if progress_callback:
                progress_callback(index, total, f"Failed {src.name}")

    report = {
        "started_at": utc_now(),
        "input_dir": str(input_dir),
        "output_dir": str(output_dir),
        "input_format": input_format,
        "pdf_engine": pdf_engine,
        "recursive": recursive,
        "overwrite": overwrite,
        "toc": toc,
        "number_sections": number_sections,
        "metadata_file": str(metadata_file.resolve()) if metadata_file else None,
        "resource_path": str(resource_path.resolve()) if resource_path else None,
        "total_found": len(md_files),
        "results": [asdict(r) for r in results],
        "summary": {
            "converted": sum(1 for r in results if r.status == "converted"),
            "skipped": sum(1 for r in results if r.status == "skipped"),
            "failed": sum(1 for r in results if r.status == "failed"),
        },
        "finished_at": utc_now(),
    }

    report_path = output_dir / report_name
    report_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
    return report


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Batch-convert Markdown files to PDF using Pandoc + xelatex with recursive directory scan, JSON reporting, and optional GUI."
    )
    p.add_argument("input_dir", nargs="?", help="Directory containing Markdown files")
    p.add_argument("output_dir", nargs="?", help="Directory where PDFs and JSON report will be written")
    p.add_argument("--input-format", default="gfm", help="Pandoc input format (default: gfm)")
    p.add_argument("--pdf-engine", default="xelatex", help="Pandoc PDF engine (default: xelatex)")
    p.add_argument("--no-recursive", action="store_true", help="Do not scan subdirectories")
    p.add_argument("--no-overwrite", action="store_true", help="Skip PDFs that already exist")
    p.add_argument("--toc", action="store_true", help="Add table of contents")
    p.add_argument("--number-sections", action="store_true", help="Number sections in output PDFs")
    p.add_argument("--metadata-file", type=Path, default=None, help="Optional Pandoc metadata YAML file")
    p.add_argument("--resource-path", type=Path, default=None, help="Optional resource path for images/includes")
    p.add_argument("--report-name", default="conversion-report.json", help="JSON report filename")
    p.add_argument("--gui", action="store_true", help="Launch GUI")
    return p.parse_args()


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Markdown to PDF Batch Converter")
        self.root.geometry("820x620")
        self.root.minsize(760, 560)

        self.queue: Queue = Queue()
        self.worker: Optional[threading.Thread] = None

        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.input_format = tk.StringVar(value="gfm")
        self.pdf_engine = tk.StringVar(value="xelatex")
        self.metadata_file = tk.StringVar()
        self.resource_path = tk.StringVar()
        self.report_name = tk.StringVar(value="conversion-report.json")
        self.recursive = tk.BooleanVar(value=True)
        self.overwrite = tk.BooleanVar(value=True)
        self.toc = tk.BooleanVar(value=False)
        self.number_sections = tk.BooleanVar(value=False)
        self.progress_value = tk.DoubleVar(value=0)
        self.status_text = tk.StringVar(value="Ready")

        self._build_ui()
        self.root.after(100, self._poll_queue)

    def _build_ui(self):
        frm = ttk.Frame(self.root, padding=16)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        row = 0
        ttk.Label(frm, text="Input folder").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(frm, textvariable=self.input_dir).grid(row=row, column=1, sticky="ew", padx=8)
        ttk.Button(frm, text="Browse", command=self._pick_input_dir).grid(row=row, column=2, sticky="ew")

        row += 1
        ttk.Label(frm, text="Output folder").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(frm, textvariable=self.output_dir).grid(row=row, column=1, sticky="ew", padx=8)
        ttk.Button(frm, text="Browse", command=self._pick_output_dir).grid(row=row, column=2, sticky="ew")

        row += 1
        ttk.Label(frm, text="Input format").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(frm, textvariable=self.input_format).grid(row=row, column=1, sticky="ew", padx=8)
        ttk.Label(frm, text="Example: gfm").grid(row=row, column=2, sticky="w")

        row += 1
        ttk.Label(frm, text="PDF engine").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(frm, textvariable=self.pdf_engine).grid(row=row, column=1, sticky="ew", padx=8)
        ttk.Label(frm, text="Example: xelatex").grid(row=row, column=2, sticky="w")

        row += 1
        ttk.Label(frm, text="Metadata file").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(frm, textvariable=self.metadata_file).grid(row=row, column=1, sticky="ew", padx=8)
        ttk.Button(frm, text="Browse", command=self._pick_metadata_file).grid(row=row, column=2, sticky="ew")

        row += 1
        ttk.Label(frm, text="Resource path").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(frm, textvariable=self.resource_path).grid(row=row, column=1, sticky="ew", padx=8)
        ttk.Button(frm, text="Browse", command=self._pick_resource_dir).grid(row=row, column=2, sticky="ew")

        row += 1
        ttk.Label(frm, text="Report name").grid(row=row, column=0, sticky="w", pady=6)
        ttk.Entry(frm, textvariable=self.report_name).grid(row=row, column=1, sticky="ew", padx=8)

        row += 1
        opts = ttk.LabelFrame(frm, text="Options", padding=12)
        opts.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(12, 8))
        ttk.Checkbutton(opts, text="Recursive scan", variable=self.recursive).grid(row=0, column=0, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(opts, text="Overwrite existing PDFs", variable=self.overwrite).grid(row=0, column=1, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(opts, text="Add table of contents", variable=self.toc).grid(row=1, column=0, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(opts, text="Number sections", variable=self.number_sections).grid(row=1, column=1, sticky="w", padx=8, pady=4)

        row += 1
        deps = ttk.LabelFrame(frm, text="Dependencies", padding=12)
        deps.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(4, 8))
        ttk.Button(deps, text="Check pandoc and engine", command=self._check_dependencies).grid(row=0, column=0, sticky="w")
        self.dep_label = ttk.Label(deps, text="Not checked yet")
        self.dep_label.grid(row=0, column=1, sticky="w", padx=12)

        row += 1
        ttk.Button(frm, text="Start conversion", command=self._start_conversion).grid(row=row, column=0, columnspan=3, sticky="ew", pady=(8, 10))

        row += 1
        ttk.Progressbar(frm, variable=self.progress_value, maximum=100).grid(row=row, column=0, columnspan=3, sticky="ew", pady=(0, 6))

        row += 1
        ttk.Label(frm, textvariable=self.status_text).grid(row=row, column=0, columnspan=3, sticky="w", pady=(0, 8))

        row += 1
        ttk.Label(frm, text="Log").grid(row=row, column=0, columnspan=3, sticky="w")

        row += 1
        self.log = tk.Text(frm, height=18, wrap="word")
        self.log.grid(row=row, column=0, columnspan=3, sticky="nsew")
        frm.rowconfigure(row, weight=1)

        scrollbar = ttk.Scrollbar(frm, orient="vertical", command=self.log.yview)
        scrollbar.grid(row=row, column=3, sticky="ns")
        self.log.configure(yscrollcommand=scrollbar.set)

    def _log(self, text: str):
        self.log.insert("end", text + "\n")
        self.log.see("end")

    def _pick_input_dir(self):
        path = filedialog.askdirectory(title="Select input folder")
        if path:
            self.input_dir.set(path)
            if not self.resource_path.get():
                self.resource_path.set(path)

    def _pick_output_dir(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.output_dir.set(path)

    def _pick_metadata_file(self):
        path = filedialog.askopenfilename(
            title="Select metadata file",
            filetypes=[("YAML files", "*.yaml *.yml"), ("All files", "*.*")],
        )
        if path:
            self.metadata_file.set(path)

    def _pick_resource_dir(self):
        path = filedialog.askdirectory(title="Select resource folder")
        if path:
            self.resource_path.set(path)

    def _check_dependencies(self):
        pandoc_ok = check_dependency("pandoc")
        engine = self.pdf_engine.get().strip() or "xelatex"
        engine_ok = check_dependency(engine)
        status = f"pandoc: {'OK' if pandoc_ok else 'missing'} | {engine}: {'OK' if engine_ok else 'missing'}"
        self.dep_label.config(text=status)
        self._log(status)

    def _start_conversion(self):
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("Conversion running", "A conversion is already in progress.")
            return

        input_dir = self.input_dir.get().strip()
        output_dir = self.output_dir.get().strip()

        if not input_dir or not output_dir:
            messagebox.showerror("Missing paths", "Please select both input and output folders.")
            return

        self.progress_value.set(0)
        self.status_text.set("Starting conversion...")
        self._log("Starting conversion...")

        def progress(done: int, total: int, message: str):
            self.queue.put(("progress", done, total, message))

        def worker():
            try:
                report = convert_all(
                    input_dir=Path(input_dir),
                    output_dir=Path(output_dir),
                    input_format=self.input_format.get().strip() or "gfm",
                    pdf_engine=self.pdf_engine.get().strip() or "xelatex",
                    overwrite=self.overwrite.get(),
                    recursive=self.recursive.get(),
                    toc=self.toc.get(),
                    number_sections=self.number_sections.get(),
                    metadata_file=Path(self.metadata_file.get()) if self.metadata_file.get().strip() else None,
                    resource_path=Path(self.resource_path.get()) if self.resource_path.get().strip() else None,
                    report_name=self.report_name.get().strip() or "conversion-report.json",
                    progress_callback=progress,
                )
                self.queue.put(("done", report))
            except Exception as e:
                self.queue.put(("error", str(e)))

        self.worker = threading.Thread(target=worker, daemon=True)
        self.worker.start()

    def _poll_queue(self):
        try:
            while True:
                item = self.queue.get_nowait()
                kind = item[0]

                if kind == "progress":
                    _, done, total, message = item
                    percent = 0 if total == 0 else (done / total) * 100
                    self.progress_value.set(percent)
                    self.status_text.set(message)
                    self._log(message)

                elif kind == "done":
                    report = item[1]
                    summary = report["summary"]
                    msg = (
                        f"Finished. Converted: {summary['converted']}, "
                        f"Skipped: {summary['skipped']}, Failed: {summary['failed']}"
                    )
                    self.progress_value.set(100)
                    self.status_text.set(msg)
                    self._log(msg)
                    self._log(f"JSON report: {Path(report['output_dir']) / self.report_name.get().strip()}")
                    messagebox.showinfo("Completed", msg)

                elif kind == "error":
                    err = item[1]
                    self.status_text.set(f"Error: {err}")
                    self._log(f"Error: {err}")
                    messagebox.showerror("Conversion error", err)

        except Empty:
            pass

        self.root.after(100, self._poll_queue)


def run_gui() -> int:
    if tk is None:
        print("ERROR: Tkinter is not available in this Python installation.", file=sys.stderr)
        return 1

    root = tk.Tk()
    App(root)
    root.mainloop()
    return 0


def run_cli(args: argparse.Namespace) -> int:
    if not args.input_dir or not args.output_dir:
        print("ERROR: input_dir and output_dir are required unless --gui is used", file=sys.stderr)
        return 1

    try:
        report = convert_all(
            input_dir=Path(args.input_dir),
            output_dir=Path(args.output_dir),
            input_format=args.input_format,
            pdf_engine=args.pdf_engine,
            overwrite=not args.no_overwrite,
            recursive=not args.no_recursive,
            toc=args.toc,
            number_sections=args.number_sections,
            metadata_file=args.metadata_file,
            resource_path=args.resource_path,
            report_name=args.report_name,
        )
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1

    print(json.dumps(report["summary"], indent=2, ensure_ascii=False))
    return 0


def main() -> int:
    args = parse_args()
    if args.gui or (not args.input_dir and not args.output_dir):
        return run_gui()
    return run_cli(args)


if __name__ == "__main__":
    raise SystemExit(main())
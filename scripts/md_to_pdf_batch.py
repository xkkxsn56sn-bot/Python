#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tkinter as tk
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Optional


@dataclass
class ConversionResult:
    source: str
    target: str
    status: str
    returncode: Optional[int] = None
    stderr: str = ""
    stdout: str = ""


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

    for src in md_files:
        rel = src.relative_to(input_dir)
        dst = (output_dir / rel).with_suffix(".pdf")
        dst.parent.mkdir(parents=True, exist_ok=True)

        if dst.exists() and not overwrite:
            results.append(ConversionResult(str(src), str(dst), "skipped"))
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
        description="Batch-convert Markdown files to PDF using Pandoc + xelatex with recursive directory scan and JSON reporting."
    )
    p.add_argument("input_dir", nargs="?", default=str(Path.cwd()), help="Directory containing Markdown files")
    p.add_argument("output_dir", nargs="?", default=str((Path.cwd() / "pdf-output").resolve()), help="Directory where PDFs and JSON report will be written")
    p.add_argument("--input-format", default="gfm", help="Pandoc input format (default: gfm)")
    p.add_argument("--pdf-engine", default="xelatex", help="Pandoc PDF engine (default: xelatex)")
    p.add_argument("--no-recursive", action="store_true", help="Do not scan subdirectories")
    p.add_argument("--no-overwrite", action="store_true", help="Skip PDFs that already exist")
    p.add_argument("--toc", action="store_true", help="Add table of contents")
    p.add_argument("--number-sections", action="store_true", help="Number sections in output PDFs")
    p.add_argument("--metadata-file", type=Path, default=None, help="Optional Pandoc metadata YAML file")
    p.add_argument("--resource-path", type=Path, default=None, help="Optional resource path for images/includes")
    p.add_argument("--report-name", default="conversion-report.json", help="JSON report filename")
    p.add_argument("--gui", action="store_true", help="Launch the desktop GUI")
    return p.parse_args()


def main() -> int:
    args = parse_args()

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


class MarkdownToPDFGUI:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Markdown → PDF Batch")
        self.root.geometry("900x620")
        self.root.minsize(820, 560)

        self.input_dir = tk.StringVar(value=str(Path.cwd()))
        self.output_dir = tk.StringVar(value=str((Path.cwd() / "pdf-output").resolve()))
        self.input_format = tk.StringVar(value="gfm")
        self.pdf_engine = tk.StringVar(value="xelatex")
        self.recursive = tk.BooleanVar(value=True)
        self.overwrite = tk.BooleanVar(value=True)
        self.toc = tk.BooleanVar(value=False)
        self.number_sections = tk.BooleanVar(value=False)
        self.report_name = tk.StringVar(value="conversion-report.json")
        self.metadata_file = tk.StringVar(value="")
        self.resource_path = tk.StringVar(value="")

        self._build_ui()

    def _build_ui(self) -> None:
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=tk.BOTH, expand=True)
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)
        main.columnconfigure(2, weight=1)
        main.rowconfigure(9, weight=1)

        ttk.Label(main, text="Input directory").grid(row=0, column=0, sticky="w")
        ttk.Entry(main, textvariable=self.input_dir).grid(row=1, column=0, columnspan=2, sticky="ew", padx=(0, 8))
        ttk.Button(main, text="Browse", command=self._pick_input).grid(row=1, column=2, sticky="ew")

        ttk.Label(main, text="Output directory").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(main, textvariable=self.output_dir).grid(row=3, column=0, columnspan=2, sticky="ew", padx=(0, 8))
        ttk.Button(main, text="Browse", command=self._pick_output).grid(row=3, column=2, sticky="ew")

        ttk.Checkbutton(main, text="Scansiona ricorsivamente", variable=self.recursive).grid(row=4, column=0, sticky="w", pady=(8, 0))
        ttk.Checkbutton(main, text="Sovrascrivi PDF esistenti", variable=self.overwrite).grid(row=4, column=1, sticky="w", pady=(8, 0))
        ttk.Checkbutton(main, text="Aggiungi TOC", variable=self.toc).grid(row=4, column=2, sticky="w", pady=(8, 0))
        ttk.Checkbutton(main, text="Numerazione sezioni", variable=self.number_sections).grid(row=5, column=0, sticky="w", pady=(4, 0))

        ttk.Label(main, text="Input format").grid(row=6, column=0, sticky="w", pady=(10, 0))
        ttk.Entry(main, textvariable=self.input_format).grid(row=7, column=0, sticky="ew", padx=(0, 8))
        ttk.Label(main, text="PDF engine").grid(row=6, column=1, sticky="w", pady=(10, 0))
        ttk.Entry(main, textvariable=self.pdf_engine).grid(row=7, column=1, sticky="ew", padx=(0, 8))
        ttk.Label(main, text="Report JSON").grid(row=6, column=2, sticky="w", pady=(10, 0))
        ttk.Entry(main, textvariable=self.report_name).grid(row=7, column=2, sticky="ew")

        ttk.Button(main, text="Run conversion", command=self._run).grid(row=8, column=0, columnspan=3, sticky="ew", pady=(12, 8))

        self.log_text = scrolledtext.ScrolledText(main, height=16, state="disabled", wrap=tk.WORD)
        self.log_text.grid(row=9, column=0, columnspan=3, sticky="nsew")

    def _pick_input(self) -> None:
        directory = filedialog.askdirectory(title="Select input directory")
        if directory:
            self.input_dir.set(directory)

    def _pick_output(self) -> None:
        directory = filedialog.askdirectory(title="Select output directory")
        if directory:
            self.output_dir.set(directory)

    def _log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def _run(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state="disabled")
        self._log("Starting conversion...")
        try:
            report = convert_all(
                input_dir=Path(self.input_dir.get().strip()),
                output_dir=Path(self.output_dir.get().strip()),
                input_format=self.input_format.get().strip(),
                pdf_engine=self.pdf_engine.get().strip(),
                overwrite=self.overwrite.get(),
                recursive=self.recursive.get(),
                toc=self.toc.get(),
                number_sections=self.number_sections.get(),
                metadata_file=Path(self.metadata_file.get().strip()).expanduser().resolve() if self.metadata_file.get().strip() else None,
                resource_path=Path(self.resource_path.get().strip()).expanduser().resolve() if self.resource_path.get().strip() else None,
                report_name=self.report_name.get().strip(),
            )
            self._log(json.dumps(report["summary"], indent=2, ensure_ascii=False))
            messagebox.showinfo("Done", "Conversion completed")
        except Exception as exc:  # noqa: BLE001
            self._log(f"ERROR: {exc}")
            messagebox.showerror("Error", str(exc))

    def run(self) -> None:
        self.root.mainloop()


def main() -> int:
    if len(sys.argv) == 1:
        MarkdownToPDFGUI().run()
        return 0

    args = parse_args()

    if args.gui:
        MarkdownToPDFGUI().run()
        return 0

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


if __name__ == "__main__":
    raise SystemExit(main())
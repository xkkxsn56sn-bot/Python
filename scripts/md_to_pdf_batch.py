#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
from pathlib import Path
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
    p.add_argument("input_dir", help="Directory containing Markdown files")
    p.add_argument("output_dir", help="Directory where PDFs and JSON report will be written")
    p.add_argument("--input-format", default="gfm", help="Pandoc input format (default: gfm)")
    p.add_argument("--pdf-engine", default="xelatex", help="Pandoc PDF engine (default: xelatex)")
    p.add_argument("--no-recursive", action="store_true", help="Do not scan subdirectories")
    p.add_argument("--no-overwrite", action="store_true", help="Skip PDFs that already exist")
    p.add_argument("--toc", action="store_true", help="Add table of contents")
    p.add_argument("--number-sections", action="store_true", help="Number sections in output PDFs")
    p.add_argument("--metadata-file", type=Path, default=None, help="Optional Pandoc metadata YAML file")
    p.add_argument("--resource-path", type=Path, default=None, help="Optional resource path for images/includes")
    p.add_argument("--report-name", default="conversion-report.json", help="JSON report filename")
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


if __name__ == "__main__":
    raise SystemExit(main())
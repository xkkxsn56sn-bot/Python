#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable


SUPPORTED_EXTENSIONS = {".docx", ".wps", ".epub", ".txt", ".rtf", ".md"}
PANDOC_EXTENSIONS = {".epub", ".md"}
MARKDOWN_EXTENSIONS = {".md"}
ProgressCallback = Callable[[int, int, str], None] | None
MACOS_SOFFICE_CANDIDATES = (
    Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
    Path("/Applications/LibreOffice.app/Contents/MacOS/soffice.bin"),
)


@dataclass
class ConversionResult:
    source: str
    target: str
    status: str
    engine: str | None = None
    returncode: int | None = None
    stdout: str = ""
    stderr: str = ""


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def check_dependency(name: str) -> bool:
    return shutil.which(name) is not None


def resolve_soffice_executable() -> str | None:
    soffice_path = shutil.which("soffice")
    if soffice_path:
        return soffice_path

    for candidate in MACOS_SOFFICE_CANDIDATES:
        if candidate.exists() and candidate.is_file() and candidate.stat().st_mode & 0o111:
            return str(candidate)

    return None


def discover_input_files(input_path: Path, recursive: bool) -> list[Path]:
    if input_path.is_file():
        if input_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise ValueError(f"Unsupported file type: {input_path.suffix}")
        return [input_path.resolve()]

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input path not found: {input_path}")

    walker = input_path.rglob if recursive else input_path.glob
    files = [
        path.resolve()
        for path in walker("*")
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS
    ]
    return sorted(files)


def is_pandoc_candidate(source: Path) -> bool:
    return source.suffix.lower() in PANDOC_EXTENSIONS


def resolve_output_path(source: Path, input_path: Path, output_dir: Path) -> Path:
    input_path = input_path.resolve()

    if input_path.is_file():
        return (output_dir / source.stem).with_suffix(".pdf")

    relative_path = source.resolve().relative_to(input_path)
    return (output_dir / relative_path).with_suffix(".pdf")


def build_pandoc_command(
    source: Path,
    destination: Path,
    pdf_engine: str,
    toc: bool,
    number_sections: bool,
    resource_path: Path | None,
) -> list[str]:
    input_format = "gfm" if source.suffix.lower() in MARKDOWN_EXTENSIONS else source.suffix.lower().lstrip(".")
    command = [
        "pandoc",
        str(source),
        "-f",
        input_format,
        "-s",
        "-o",
        str(destination),
        "--pdf-engine",
        pdf_engine,
    ]

    if toc:
        command.append("--toc")
    if number_sections:
        command.append("--number-sections")
    if resource_path:
        command.extend(["--resource-path", str(resource_path)])

    return command


def convert_with_pandoc(
    source: Path,
    destination: Path,
    pdf_engine: str,
    toc: bool,
    number_sections: bool,
    resource_path: Path | None,
) -> subprocess.CompletedProcess[str]:
    command = build_pandoc_command(
        source=source,
        destination=destination,
        pdf_engine=pdf_engine,
        toc=toc,
        number_sections=number_sections,
        resource_path=resource_path,
    )
    return subprocess.run(command, capture_output=True, text=True)


def convert_with_libreoffice(source: Path, destination: Path, soffice_executable: str) -> subprocess.CompletedProcess[str]:
    with tempfile.TemporaryDirectory(prefix="pdf-convert-") as tmpdir:
        temp_output_dir = Path(tmpdir)
        command = [
            soffice_executable,
            "--headless",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "--outdir",
            str(temp_output_dir),
            str(source),
        ]
        result = subprocess.run(command, capture_output=True, text=True)

        if result.returncode != 0:
            return result

        generated_pdf = temp_output_dir / f"{source.stem}.pdf"
        if not generated_pdf.exists():
            return subprocess.CompletedProcess(
                args=command,
                returncode=1,
                stdout=result.stdout,
                stderr=(result.stderr + "\nLibreOffice did not produce the expected PDF file.").strip(),
            )

        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(generated_pdf), str(destination))
        return result


def convert_one(
    source: Path,
    destination: Path,
    prefer_pandoc: bool,
    pdf_engine: str,
    toc: bool,
    number_sections: bool,
    resource_path: Path | None,
    pandoc_available: bool,
    soffice_available: bool,
    soffice_executable: str | None,
) -> ConversionResult:
    strategies: list[tuple[str, Callable[[], subprocess.CompletedProcess[str]]]] = []

    if prefer_pandoc and is_pandoc_candidate(source) and pandoc_available:
        strategies.append(
            (
                "pandoc",
                lambda: convert_with_pandoc(
                    source=source,
                    destination=destination,
                    pdf_engine=pdf_engine,
                    toc=toc,
                    number_sections=number_sections,
                    resource_path=resource_path,
                ),
            )
        )

    if soffice_available and soffice_executable:
        strategies.append(
            (
                "libreoffice",
                lambda: convert_with_libreoffice(
                    source=source,
                    destination=destination,
                    soffice_executable=soffice_executable,
                ),
            )
        )

    if not prefer_pandoc and is_pandoc_candidate(source) and pandoc_available:
        strategies.append(
            (
                "pandoc",
                lambda: convert_with_pandoc(
                    source=source,
                    destination=destination,
                    pdf_engine=pdf_engine,
                    toc=toc,
                    number_sections=number_sections,
                    resource_path=resource_path,
                ),
            )
        )

    last_result: ConversionResult | None = None
    for engine_name, converter in strategies:
        run = converter()
        if run.returncode == 0 and destination.exists():
            return ConversionResult(
                source=str(source),
                target=str(destination),
                status="converted",
                engine=engine_name,
                returncode=0,
                stdout=run.stdout.strip(),
                stderr=run.stderr.strip(),
            )

        last_result = ConversionResult(
            source=str(source),
            target=str(destination),
            status="failed",
            engine=engine_name,
            returncode=run.returncode,
            stdout=run.stdout.strip(),
            stderr=run.stderr.strip(),
        )

    return last_result or ConversionResult(
        source=str(source),
        target=str(destination),
        status="failed",
        stderr="No conversion strategy was available.",
    )


def ensure_dependencies(prefer_pandoc: bool, pdf_engine: str) -> None:
    if prefer_pandoc:
        if not check_dependency("pandoc"):
            raise RuntimeError("Pandoc is not installed or not available in PATH")
        if not check_dependency(pdf_engine):
            raise RuntimeError(f"PDF engine '{pdf_engine}' is not installed or not available in PATH")


def resolve_pandoc_availability(files: list[Path], prefer_pandoc: bool, pdf_engine: str) -> bool:
    needs_pandoc = any(is_pandoc_candidate(path) for path in files)
    if not needs_pandoc:
        return False

    if not check_dependency("pandoc"):
        raise RuntimeError("Pandoc is required for .md or .epub conversion but is not available in PATH")

    if not check_dependency(pdf_engine):
        raise RuntimeError(f"PDF engine '{pdf_engine}' is not installed or not available in PATH")

    return True


def resolve_soffice_availability(files: list[Path]) -> tuple[bool, str | None]:
    needs_soffice = any(not is_pandoc_candidate(path) for path in files)
    if not needs_soffice:
        return False, None

    soffice_executable = resolve_soffice_executable()
    if not soffice_executable:
        raise RuntimeError(
            "LibreOffice is required for .docx, .wps, .txt, or .rtf conversion but no usable 'soffice' executable was found"
        )

    return True, soffice_executable


def convert_all(
    input_path: Path,
    output_dir: Path,
    recursive: bool,
    overwrite: bool,
    prefer_pandoc: bool,
    pdf_engine: str,
    toc: bool,
    number_sections: bool,
    resource_path: Path | None,
    report_name: str,
    progress_callback: ProgressCallback = None,
) -> dict:
    input_path = input_path.resolve()
    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    files = discover_input_files(input_path=input_path, recursive=recursive)
    ensure_dependencies(prefer_pandoc=prefer_pandoc, pdf_engine=pdf_engine)
    pandoc_available = resolve_pandoc_availability(files=files, prefer_pandoc=prefer_pandoc, pdf_engine=pdf_engine)
    soffice_available, soffice_executable = resolve_soffice_availability(files=files)
    results: list[ConversionResult] = []
    total = len(files)

    for index, source in enumerate(files, start=1):
        destination = resolve_output_path(source=source, input_path=input_path, output_dir=output_dir)
        destination.parent.mkdir(parents=True, exist_ok=True)

        if progress_callback:
            progress_callback(index - 1, total, f"Preparing {source.name}")

        if destination.exists() and not overwrite:
            results.append(
                ConversionResult(
                    source=str(source),
                    target=str(destination),
                    status="skipped",
                )
            )
            if progress_callback:
                progress_callback(index, total, f"Skipped {source.name}")
            continue

        result = convert_one(
            source=source,
            destination=destination,
            prefer_pandoc=prefer_pandoc,
            pdf_engine=pdf_engine,
            toc=toc,
            number_sections=number_sections,
            resource_path=resource_path,
            pandoc_available=pandoc_available,
            soffice_available=soffice_available,
            soffice_executable=soffice_executable,
        )
        results.append(result)
        if progress_callback:
            label = "Converted" if result.status == "converted" else "Failed"
            progress_callback(index, total, f"{label} {source.name}")

    report = {
        "started_at": utc_now(),
        "input_path": str(input_path),
        "output_dir": str(output_dir),
        "recursive": recursive,
        "overwrite": overwrite,
        "prefer_pandoc": prefer_pandoc,
        "pdf_engine": pdf_engine,
        "resource_path": str(resource_path.resolve()) if resource_path else None,
        "total_found": len(files),
        "results": [asdict(result) for result in results],
        "summary": {
            "converted": sum(1 for result in results if result.status == "converted"),
            "skipped": sum(1 for result in results if result.status == "skipped"),
            "failed": sum(1 for result in results if result.status == "failed"),
        },
        "finished_at": utc_now(),
    }

    report_path = output_dir / report_name
    report_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
    return report


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Convert docx, wps, epub, txt, rtf, and md files to PDF. "
            "Uses LibreOffice for office/text formats and can prefer Pandoc for epub/markdown."
        )
    )
    parser.add_argument("input_path", help="Input file or directory")
    parser.add_argument("output_dir", help="Destination directory for generated PDFs")
    parser.add_argument("--no-recursive", action="store_true", help="Do not scan subdirectories when input_path is a directory")
    parser.add_argument("--no-overwrite", action="store_true", help="Skip PDFs that already exist")
    parser.add_argument("--prefer-pandoc", action="store_true", help="Use Pandoc first for .md and .epub files")
    parser.add_argument("--pdf-engine", default="xelatex", help="Pandoc PDF engine when --prefer-pandoc is used")
    parser.add_argument("--toc", action="store_true", help="Add a table of contents for Pandoc conversions")
    parser.add_argument("--number-sections", action="store_true", help="Number sections for Pandoc conversions")
    parser.add_argument("--resource-path", type=Path, default=None, help="Pandoc resource path for markdown assets")
    parser.add_argument("--report-name", default="to-pdf-report.json", help="JSON report filename")
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    try:
        report = convert_all(
            input_path=Path(args.input_path),
            output_dir=Path(args.output_dir),
            recursive=not args.no_recursive,
            overwrite=not args.no_overwrite,
            prefer_pandoc=args.prefer_pandoc,
            pdf_engine=args.pdf_engine,
            toc=args.toc,
            number_sections=args.number_sections,
            resource_path=args.resource_path,
            report_name=args.report_name,
        )
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print(json.dumps(report["summary"], indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
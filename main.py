#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path


ROOT_DIR = Path(__file__).resolve().parent
SCRIPTS_DIR = ROOT_DIR / "scripts"


@dataclass(frozen=True)
class ScriptEntry:
	key: str
	filename: str
	path: Path


def slugify(name: str) -> str:
	slug = re.sub(r"[^a-zA-Z0-9]+", "_", name.strip().lower()).strip("_")
	return slug or "script"


def discover_scripts(directory: Path) -> dict[str, ScriptEntry]:
	if not directory.exists():
		return {}

	entries: dict[str, ScriptEntry] = {}
	for script_path in sorted(directory.glob("*.py")):
		key = slugify(script_path.stem)
		entries[key] = ScriptEntry(
			key=key,
			filename=script_path.name,
			path=script_path,
		)
	return entries


def parse_args() -> argparse.Namespace:
	parser = argparse.ArgumentParser(
		description="Launcher for personal scripts in ./scripts",
		epilog=(
			"Examples:\n"
			"  python main.py list\n"
			"  python main.py run md_to_pdf_batch -- --help"
		),
		formatter_class=argparse.RawTextHelpFormatter,
	)

	subparsers = parser.add_subparsers(dest="command", required=True)
	subparsers.add_parser("list", help="List available scripts")

	run_parser = subparsers.add_parser("run", help="Run one script from ./scripts")
	run_parser.add_argument(
		"script",
		help="Script key shown by 'list' or exact filename in scripts/",
	)
	run_parser.add_argument(
		"script_args",
		nargs=argparse.REMAINDER,
		help="Arguments forwarded to the target script (prefix with --)",
	)

	return parser.parse_args()


def print_scripts(scripts: dict[str, ScriptEntry]) -> None:
	if not scripts:
		print("No scripts found in ./scripts")
		return

	print("Available scripts:")
	for key, entry in scripts.items():
		print(f"  - {key:32} -> {entry.filename}")


def resolve_script(target: str, scripts: dict[str, ScriptEntry]) -> ScriptEntry | None:
	key = slugify(target)
	if key in scripts:
		return scripts[key]

	# Fallback: allow exact filename lookup.
	for entry in scripts.values():
		if entry.filename == target:
			return entry
	return None


def run_script(entry: ScriptEntry, script_args: list[str]) -> int:
	forwarded = script_args[1:] if script_args and script_args[0] == "--" else script_args
	cmd = [sys.executable, str(entry.path), *forwarded]
	completed = subprocess.run(cmd)
	return completed.returncode


def main() -> int:
	args = parse_args()
	scripts = discover_scripts(SCRIPTS_DIR)

	if args.command == "list":
		print_scripts(scripts)
		return 0

	if args.command == "run":
		entry = resolve_script(args.script, scripts)
		if entry is None:
			print(f"Script not found: {args.script}", file=sys.stderr)
			print_scripts(scripts)
			return 2
		return run_script(entry, args.script_args)

	print(f"Unknown command: {args.command}", file=sys.stderr)
	return 2


if __name__ == "__main__":
	raise SystemExit(main())

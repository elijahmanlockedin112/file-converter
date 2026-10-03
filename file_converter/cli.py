"""Command-line interface. Run ``python -m file_converter --help``."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .backends import backend_for, supported_exts
from .batch import Status, find_sources, plan_jobs, run_jobs


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python -m file_converter",
        description=(
            "Convert documents to PDF locally with Microsoft Word. "
            f"Supported types: {', '.join(supported_exts())}. "
            "Run with no arguments to open the app window instead."
        ),
    )
    parser.add_argument("paths", nargs="+", type=Path, help="files and/or folders to convert")
    parser.add_argument(
        "-o", "--out", type=Path,
        help="folder to save the PDFs in (default: next to each original)",
    )
    parser.add_argument(
        "-r", "--recursive", action="store_true",
        help="also convert files inside subfolders of any folder given",
    )
    parser.add_argument(
        "--overwrite", action="store_true",
        help="replace PDFs that already exist (default: skip them)",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    # Don't crash printing a file name the console's code page can't show.
    sys.stdout.reconfigure(errors="replace")

    for path in args.paths:
        if not path.exists():
            print(f"warning: {path} doesn't exist", file=sys.stderr)
        elif path.is_file() and backend_for(path) is None:
            print(f"warning: {path} isn't a supported file type", file=sys.stderr)

    sources = find_sources(args.paths, args.recursive)
    if not sources:
        print("Nothing to convert.", file=sys.stderr)
        return 1
    jobs = plan_jobs(sources, args.out)
    width = len(str(len(jobs)))

    def report(i, job, status, message):
        if status is Status.CONVERTING:
            return
        line = f"[{i + 1:>{width}}/{len(jobs)}] {status.value:<7} {job.src.name}"
        if status is Status.DONE:
            line += f" -> {job.dst}"
        if message:
            line += f"  ({message})"
        print(line, flush=True)

    try:
        summary = run_jobs(jobs, overwrite=args.overwrite, on_update=report)
    except KeyboardInterrupt:
        print("\nCancelled.")
        return 130
    print(f"\n{summary.done} converted, {summary.skipped} skipped, {summary.failed} failed.")
    return 1 if summary.failed else 0

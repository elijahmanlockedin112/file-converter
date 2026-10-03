"""Turning files and folders into conversion jobs, and running those jobs."""

from __future__ import annotations

import os
import threading
from contextlib import ExitStack
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Callable, Iterable

from .backends import ConversionError, backend_for


@dataclass(frozen=True)
class Source:
    """A file to convert.

    ``root`` is set when the file was found by scanning a folder; it lets the
    folder's subfolder layout be recreated inside an output folder.
    """

    path: Path
    root: Path | None = None


@dataclass(frozen=True)
class Job:
    src: Path
    dst: Path
    backend: type


class Status(Enum):
    CONVERTING = "Converting…"
    DONE = "Done"
    SKIPPED = "Skipped"
    FAILED = "Failed"


@dataclass
class Summary:
    done: int = 0
    skipped: int = 0
    failed: int = 0
    cancelled: bool = False


# Called as on_update(index, job, status, message) as each job progresses.
UpdateCallback = Callable[[int, Job, Status, str], None]


def path_key(path: Path) -> str:
    """A key that treats two spellings of the same Windows path as equal."""
    return os.path.normcase(os.path.abspath(path))


def is_convertible(path: Path) -> bool:
    # Word leaves "~$name.docx" lock files beside open documents; skip them.
    return (
        path.is_file()
        and not path.name.startswith("~$")
        and backend_for(path) is not None
    )


def find_sources(paths: Iterable[Path], recursive: bool = False) -> list[Source]:
    """Expand files and folders into a de-duplicated list of convertible files."""
    sources: list[Source] = []
    seen: set[str] = set()
    for path in paths:
        path = Path(path).resolve()
        if path.is_dir():
            candidates = sorted(path.glob("**/*" if recursive else "*"))
            root = path
        else:
            candidates = [path]
            root = None
        for candidate in candidates:
            key = path_key(candidate)
            if key not in seen and is_convertible(candidate):
                seen.add(key)
                sources.append(Source(candidate, root))
    return sources


def plan_jobs(sources: Iterable[Source], out_dir: Path | None = None) -> list[Job]:
    """Decide where each source's output file goes.

    With no ``out_dir``, each output goes next to its source. Otherwise outputs
    go into ``out_dir``, keeping the subfolder layout of scanned folders.
    Sources that would produce the same output (``report.doc`` and
    ``report.docx``) get numbered names rather than overwriting each other.
    """
    if out_dir is not None:
        out_dir = Path(out_dir).resolve()
    taken: set[str] = set()
    jobs = []
    for source in sources:
        backend = backend_for(source.path)
        if out_dir is None:
            folder = source.path.parent
        elif source.root is not None:
            folder = out_dir / source.path.parent.relative_to(source.root)
        else:
            folder = out_dir
        dst = _unused_name(folder, source.path.stem, backend.output_ext, taken)
        jobs.append(Job(source.path, dst, backend))
    return jobs


def _unused_name(folder: Path, stem: str, ext: str, taken: set[str]) -> Path:
    dst = folder / f"{stem}{ext}"
    n = 2
    while path_key(dst) in taken:
        dst = folder / f"{stem} ({n}){ext}"
        n += 1
    taken.add(path_key(dst))
    return dst


def run_jobs(
    jobs: list[Job],
    overwrite: bool = False,
    on_update: UpdateCallback | None = None,
    cancel: threading.Event | None = None,
) -> Summary:
    """Convert each job in order. Existing outputs are skipped unless
    ``overwrite`` is set. Setting ``cancel`` stops after the current file."""
    notify = on_update or (lambda *args: None)
    summary = Summary()
    with ExitStack() as stack:
        open_backends = {}
        for i, job in enumerate(jobs):
            if cancel is not None and cancel.is_set():
                summary.cancelled = True
                break
            if job.dst.exists() and not overwrite:
                summary.skipped += 1
                notify(i, job, Status.SKIPPED, f"{job.dst.name} already exists")
                continue

            notify(i, job, Status.CONVERTING, "")
            try:
                backend = open_backends.get(job.backend)
                if backend is None:
                    backend = stack.enter_context(job.backend())
                    open_backends[job.backend] = backend
                job.dst.parent.mkdir(parents=True, exist_ok=True)
                backend.convert(job.src, job.dst)
            except (ConversionError, OSError) as e:
                summary.failed += 1
                notify(i, job, Status.FAILED, str(e))
            except Exception as e:  # one bad file shouldn't sink the batch
                summary.failed += 1
                notify(i, job, Status.FAILED, f"{type(e).__name__}: {e}")
            else:
                summary.done += 1
                notify(i, job, Status.DONE, "")
    return summary

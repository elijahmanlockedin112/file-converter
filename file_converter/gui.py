"""Desktop app: add files, choose where the PDFs go, convert."""

from __future__ import annotations

import os
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, font, messagebox, ttk

from .backends import supported_exts
from .batch import Source, Status, find_sources, path_key, plan_jobs, run_jobs

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:  # drag and drop is optional
    TkinterDnD = None

POLL_MS = 100


class App:
    def __init__(self, root: tk.Tk, initial_paths=()) -> None:
        self.root = root
        self.sources: list[Source] = []
        self.rows: list[str] = []  # tree item id for each source, same order
        self.outputs: dict[str, Path] = {}  # tree item id -> converted file
        self.last_output_folder: Path | None = None

        self.events: queue.Queue = queue.Queue()
        self.cancel = threading.Event()
        self.worker: threading.Thread | None = None
        self.job_count = 0
        self.closing = False

        self.out_mode = tk.StringVar(value="beside")
        self.out_dir = tk.StringVar()
        self.recursive = tk.BooleanVar(value=True)
        self.overwrite = tk.BooleanVar(value=False)
        self.status = tk.StringVar()

        self._build()
        self.add_paths(initial_paths)
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.after(POLL_MS, self._poll)

    # ---- layout ---------------------------------------------------------

    def _build(self) -> None:
        root = self.root
        root.title("File Converter")
        root.geometry("820x520")
        root.minsize(620, 400)

        line = font.nametofont("TkDefaultFont").metrics("linespace")
        style = ttk.Style(root)
        style.configure("Treeview", rowheight=line + 8)
        style.configure("Accent.TButton", padding=(16, 4))

        outer = ttk.Frame(root, padding=12)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        # Toolbar
        bar = ttk.Frame(outer)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self.lockable = [
            ttk.Button(bar, text="Add files…", command=self.add_files),
            ttk.Button(bar, text="Add folder…", command=self.add_folder),
            ttk.Button(bar, text="Remove", command=self.remove_selected),
            ttk.Button(bar, text="Clear", command=self.clear),
        ]
        for i, button in enumerate(self.lockable):
            button.pack(side="left", padx=(0 if i == 0 else 6, 0))
        subfolders = ttk.Checkbutton(
            bar, text="Include subfolders when adding a folder", variable=self.recursive
        )
        subfolders.pack(side="right")
        self.lockable.append(subfolders)

        # File list
        table = ttk.Frame(outer)
        table.grid(row=1, column=0, sticky="nsew")
        table.columnconfigure(0, weight=1)
        table.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(
            table, columns=("name", "folder", "status"), show="headings", selectmode="extended"
        )
        for column, heading, width, stretch in (
            ("name", "File", 220, False),
            ("folder", "Folder", 300, True),
            ("status", "Status", 220, False),
        ):
            self.tree.heading(column, text=heading, anchor="w")
            self.tree.column(column, width=width, stretch=stretch, anchor="w")
        self.tree.tag_configure("done", foreground="#1a7f37")
        self.tree.tag_configure("failed", foreground="#cf222e")
        self.tree.tag_configure("skipped", foreground="#6e7781")
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())
        self.tree.bind("<Double-1>", self._open_row)

        hint = "Drag Word files or folders here" if TkinterDnD else "Use “Add files…” or “Add folder…”"
        self.placeholder = ttk.Label(
            self.tree,
            text=f"{hint}\n\nSupported: {', '.join(supported_exts())}",
            justify="center",
            foreground="#6e7781",
            background=style.lookup("Treeview", "fieldbackground") or "white",
        )

        if TkinterDnD:
            for widget in (self.tree, self.placeholder):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._on_drop)

        # Output options
        options = ttk.LabelFrame(outer, text="Save PDFs", padding=(10, 6))
        options.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        options.columnconfigure(1, weight=1)
        beside = ttk.Radiobutton(
            options, text="Next to each original file", value="beside", variable=self.out_mode
        )
        beside.grid(row=0, column=0, columnspan=3, sticky="w")
        into = ttk.Radiobutton(options, text="In this folder:", value="folder", variable=self.out_mode)
        into.grid(row=1, column=0, sticky="w", pady=4)
        entry = ttk.Entry(options, textvariable=self.out_dir)
        entry.grid(row=1, column=1, sticky="ew", padx=6)
        entry.bind("<FocusIn>", lambda e: self.out_mode.set("folder"))
        browse = ttk.Button(options, text="Browse…", command=self.choose_out_dir)
        browse.grid(row=1, column=2)
        overwrite = ttk.Checkbutton(
            options, text="Replace PDFs that already exist", variable=self.overwrite
        )
        overwrite.grid(row=2, column=0, columnspan=3, sticky="w")
        self.lockable += [beside, into, entry, browse, overwrite]

        # Progress and actions
        bottom = ttk.Frame(outer)
        bottom.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        bottom.columnconfigure(0, weight=1)
        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.open_button = ttk.Button(
            bottom, text="Open output folder", command=self.open_output_folder, state="disabled"
        )
        self.open_button.grid(row=0, column=1, padx=(0, 6))
        self.convert_button = ttk.Button(
            bottom, text="Convert to PDF", style="Accent.TButton", command=self.convert_or_stop
        )
        self.convert_button.grid(row=0, column=2)
        ttk.Label(bottom, textvariable=self.status).grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(6, 0)
        )

        self._refresh_placeholder()

    def _refresh_placeholder(self) -> None:
        if self.rows:
            self.placeholder.place_forget()
        else:
            self.placeholder.place(relx=0.5, rely=0.5, anchor="center")

    # ---- managing the file list -----------------------------------------

    def add_files(self) -> None:
        patterns = " ".join(f"*{ext}" for ext in supported_exts())
        paths = filedialog.askopenfilenames(
            title="Choose documents to convert",
            filetypes=[("Word documents", patterns), ("All files", "*.*")],
        )
        self.add_paths(paths)

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(title="Choose a folder of documents", mustexist=True)
        if folder:
            self.add_paths([folder])

    def _on_drop(self, event) -> None:
        self.add_paths(self.root.tk.splitlist(event.data))

    def add_paths(self, paths) -> None:
        if not paths or self.worker:
            return
        known = {path_key(s.path) for s in self.sources}
        found = find_sources([Path(p) for p in paths], self.recursive.get())
        added = [s for s in found if path_key(s.path) not in known]
        for source in added:
            row = self.tree.insert(
                "", "end", values=(source.path.name, _short_folder(source.path.parent), "")
            )
            self.sources.append(source)
            self.rows.append(row)
        if added:
            self.status.set(f"Added {_plural(len(added), 'file')} · {len(self.sources)} in the list")
        elif found:
            self.status.set("Those files are already in the list.")
        else:
            self.status.set("No Word documents found there.")
        self._refresh_placeholder()

    def remove_selected(self) -> None:
        if self.worker:
            return
        selected = set(self.tree.selection())
        if not selected:
            return
        kept = [(s, r) for s, r in zip(self.sources, self.rows) if r not in selected]
        self.sources = [s for s, _ in kept]
        self.rows = [r for _, r in kept]
        for row in selected:
            self.outputs.pop(row, None)
        self.tree.delete(*selected)
        self.status.set(f"{_plural(len(self.sources), 'file')} in the list")
        self._refresh_placeholder()

    def clear(self) -> None:
        if self.worker:
            return
        self.tree.delete(*self.rows)
        self.sources.clear()
        self.rows.clear()
        self.outputs.clear()
        self.progress["value"] = 0
        self.status.set("")
        self._refresh_placeholder()

    def choose_out_dir(self) -> None:
        folder = filedialog.askdirectory(title="Save PDFs in…")
        if folder:
            self.out_dir.set(os.path.normpath(folder))
            self.out_mode.set("folder")

    # ---- converting -----------------------------------------------------

    def convert_or_stop(self) -> None:
        if self.worker:
            self.cancel.set()
            self.convert_button.state(["disabled"])
            self.status.set("Stopping after the current file…")
            return

        if not self.sources:
            messagebox.showinfo("Nothing to convert", "Add some Word documents first.")
            return
        out_dir = None
        if self.out_mode.get() == "folder":
            text = self.out_dir.get().strip()
            if not text:
                messagebox.showwarning("Choose a folder", "Pick the folder to save the PDFs in.")
                return
            out_dir = Path(text)

        jobs = plan_jobs(self.sources, out_dir)
        self.job_count = len(jobs)
        for row in self.rows:
            self.tree.set(row, "status", "Waiting")
            self.tree.item(row, tags=())
        self.outputs.clear()
        self.last_output_folder = out_dir.resolve() if out_dir else None
        self.progress.configure(maximum=len(jobs), value=0)
        self.cancel.clear()
        self._set_running(True)
        self.status.set("Starting Word…")

        self.worker = threading.Thread(
            target=self._work, args=(jobs, self.overwrite.get()), daemon=True
        )
        self.worker.start()

    def _work(self, jobs, overwrite: bool) -> None:
        """Runs on the worker thread; only talks to the UI through the queue."""
        try:
            summary = run_jobs(
                jobs,
                overwrite=overwrite,
                on_update=lambda *update: self.events.put(("update", *update)),
                cancel=self.cancel,
            )
        except Exception as e:
            self.events.put(("crashed", e))
        else:
            self.events.put(("finished", summary))

    def _poll(self) -> None:
        try:
            while True:
                kind, *payload = self.events.get_nowait()
                if kind == "update":
                    self._on_update(*payload)
                else:
                    self._on_finished(kind, payload[0])
                    if self.closing:
                        self.root.destroy()
                        return
        except queue.Empty:
            pass
        self.root.after(POLL_MS, self._poll)

    def _on_update(self, i: int, job, status: Status, message: str) -> None:
        row = self.rows[i]
        self.tree.set(row, "status", f"{status.value}: {message}" if message else status.value)
        self.tree.item(row, tags=(status.name.lower(),))
        if status is Status.CONVERTING:
            self.tree.see(row)
            if not self.cancel.is_set():
                self.status.set(f"Converting {i + 1} of {self.job_count}: {job.src.name}")
            return
        self.progress["value"] = i + 1
        if status is Status.DONE:
            self.outputs[row] = job.dst
            if self.last_output_folder is None:
                self.last_output_folder = job.dst.parent

    def _on_finished(self, kind: str, result) -> None:
        self.worker = None
        self._set_running(False)
        if kind == "crashed":
            self.status.set(f"Stopped by an unexpected error: {result}")
            messagebox.showerror("Conversion stopped", str(result))
            return
        parts = [f"{result.done} converted"]
        if result.skipped:
            parts.append(f"{result.skipped} skipped")
        if result.failed:
            parts.append(f"{result.failed} failed")
        prefix = "Stopped. " if result.cancelled else "Finished. "
        self.status.set(prefix + " · ".join(parts))

    def _set_running(self, running: bool) -> None:
        for widget in self.lockable:
            widget.state(["disabled"] if running else ["!disabled"])
        self.convert_button.state(["!disabled"])
        self.convert_button.configure(text="Stop" if running else "Convert to PDF")
        has_output = not running and self.last_output_folder is not None and bool(self.outputs)
        self.open_button.state(["!disabled"] if has_output else ["disabled"])

    # ---- opening results ------------------------------------------------

    def open_output_folder(self) -> None:
        if self.last_output_folder and self.last_output_folder.exists():
            os.startfile(self.last_output_folder)

    def _open_row(self, event) -> None:
        row = self.tree.identify_row(event.y)
        if row in self.outputs and self.outputs[row].exists():
            os.startfile(self.outputs[row])

    def on_close(self) -> None:
        if self.worker:
            if not messagebox.askyesno(
                "Conversion in progress", "Stop after the current file and close?"
            ):
                return
            # _poll closes the window once Word has been shut down cleanly.
            self.closing = True
            self.cancel.set()
            self.status.set("Closing after the current file…")
            return
        self.root.destroy()


def _short_folder(folder: Path) -> str:
    """Keep the end of long paths, which is the part that tells files apart."""
    parts = folder.parts
    return str(folder) if len(parts) <= 4 else os.path.join("…", *parts[-3:])


def _plural(n: int, word: str) -> str:
    return f"{n} {word}{'' if n == 1 else 's'}"


def main(paths=None) -> None:
    if sys.platform == "win32":
        try:  # crisp text on high-DPI screens
            import ctypes

            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError, OSError):
            pass
    root = TkinterDnD.Tk() if TkinterDnD else tk.Tk()
    App(root, paths or [])
    root.mainloop()

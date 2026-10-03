"""Conversion backends, and the registry that maps file types to them.

A backend is a context manager with a ``convert(src, dst)`` method. The batch
runner opens each backend once and reuses it for every file it handles, so a
backend can keep an expensive resource (like a running copy of Word) alive for
the whole batch.

To support a new format, write a backend class and add it to ``BACKENDS``.
"""

from __future__ import annotations

import shutil
import tempfile
import threading
import uuid
from pathlib import Path

# Longest file path Word will open or save (it errors above 255; keep a margin).
WORD_MAX_PATH = 250

# Word constants (from the Word object model docs).
WD_ALERTS_NONE = 0
WD_DO_NOT_SAVE_CHANGES = 0
WD_EXPORT_FORMAT_PDF = 17
WD_EXPORT_OPTIMIZE_FOR_PRINT = 0
WD_EXPORT_CREATE_HEADING_BOOKMARKS = 1
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

# Passing a password Word will never match makes password-protected files fail
# with an error instead of popping up an invisible password prompt that hangs.
# (It must be a normal string: Word treats one starting with "\0" as empty.)
_NO_PASSWORD = "file-converter-no-password-9f3b2c7e"


class ConversionError(Exception):
    """A file could not be converted. The message is meant for the user."""


class WordToPdf:
    """Converts Word documents to PDF with a hidden copy of Microsoft Word.

    The output is exactly what Word's own "Save as PDF" produces. Word is
    started on first use and quit when the context manager exits. It runs as a
    separate instance, so any Word windows you have open are left alone.
    """

    name = "Word"
    input_exts = frozenset({".doc", ".docx"})
    output_ext = ".pdf"

    # If one file takes longer than this, Word is assumed to be stuck (e.g. on
    # a dialog nobody can see). It gets killed, and restarted for the next file.
    timeout_seconds = 300

    def __init__(self) -> None:
        self._word = None
        self._pid: int | None = None
        self._timed_out = False

    def __enter__(self) -> WordToPdf:
        import pythoncom

        # COM must be initialized on every thread that uses it (the GUI
        # converts on a background thread).
        pythoncom.CoInitialize()
        return self

    def __exit__(self, *exc_info) -> None:
        import pythoncom

        self._quit()
        pythoncom.CoUninitialize()

    def convert(self, src: Path, dst: Path) -> None:
        if len(str(src)) <= WORD_MAX_PATH and len(str(dst)) <= WORD_MAX_PATH:
            self._export(src, dst)
            return
        # Word rejects paths longer than 255 characters, so convert a copy at
        # a short temporary path and move the PDF into place afterwards.
        with tempfile.TemporaryDirectory(prefix="fc", ignore_cleanup_errors=True) as tmp:
            short_src = Path(tmp, f"in{src.suffix}")
            short_dst = Path(tmp, f"out{dst.suffix}")
            shutil.copyfile(src, short_src)
            self._export(short_src, short_dst)
            shutil.move(short_dst, dst)

    def _export(self, src: Path, dst: Path) -> None:
        word = self._ensure_word()
        self._timed_out = False
        watchdog = threading.Timer(self.timeout_seconds, self._kill_word)
        watchdog.daemon = True
        watchdog.start()
        try:
            self._open_and_export(word, src, dst)
        finally:
            watchdog.cancel()
            watchdog.join()  # if it fired, wait until Word has fully exited

    def _open_and_export(self, word, src: Path, dst: Path) -> None:
        import pywintypes

        try:
            doc = word.Documents.Open(
                FileName=str(src),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                PasswordDocument=_NO_PASSWORD,
                Visible=False,
                NoEncodingDialog=True,
            )
        except pywintypes.com_error as e:
            raise self._failure(e, "Word couldn't open the file") from e

        try:
            doc.ExportAsFixedFormat(
                OutputFileName=str(dst),
                ExportFormat=WD_EXPORT_FORMAT_PDF,
                OpenAfterExport=False,
                OptimizeFor=WD_EXPORT_OPTIMIZE_FOR_PRINT,
                CreateBookmarks=WD_EXPORT_CREATE_HEADING_BOOKMARKS,
                DocStructureTags=True,
            )
        except pywintypes.com_error as e:
            raise self._failure(e, "Word couldn't save the PDF") from e
        finally:
            try:
                doc.Close(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
            except pywintypes.com_error:
                pass  # Word died mid-export; _ensure_word restarts it next time.

    def _failure(self, error, fallback: str) -> ConversionError:
        if self._timed_out:
            secs = self.timeout_seconds
            limit = f"{secs // 60} min" if secs >= 60 else f"{secs} s"
            return ConversionError(f"Word got stuck on this file, gave up after {limit}")
        return ConversionError(_describe(error, fallback))

    def _kill_word(self) -> None:
        """Runs on the watchdog thread when a file takes too long."""
        import pywintypes
        import win32api
        import win32con
        import win32event

        self._timed_out = True
        if self._pid is None:
            return
        try:
            access = win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE
            handle = win32api.OpenProcess(access, False, self._pid)
            try:
                win32api.TerminateProcess(handle, 1)
                # Wait for it to really exit, so it lets go of its open files.
                win32event.WaitForSingleObject(handle, 10_000)
            finally:
                win32api.CloseHandle(handle)
        except pywintypes.error:
            pass  # already gone

    def _ensure_word(self):
        """Return a running Word instance, (re)starting it if needed."""
        import pywintypes
        import win32com.client

        if self._word is not None:
            try:
                self._word.Visible  # cheap call to check Word is still alive
                return self._word
            except pywintypes.com_error:
                self._word = None

        try:
            word = win32com.client.DispatchEx("Word.Application")
        except pywintypes.com_error as e:
            raise ConversionError(
                "Couldn't start Microsoft Word. Is it installed and activated?"
            ) from e
        word.Visible = False
        word.DisplayAlerts = WD_ALERTS_NONE
        # Never run macros from the documents being converted.
        word.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
        self._word = word
        self._pid = _find_pid(word)
        return word

    def _quit(self) -> None:
        if self._word is None:
            return
        try:
            self._word.Quit(SaveChanges=WD_DO_NOT_SAVE_CHANGES)
        except Exception:
            pass
        self._word = None
        self._pid = None


def _find_pid(word) -> int | None:
    """Find the process ID of a Word instance we started, so the watchdog can
    kill exactly that one. Word has no PID property, so tag its (hidden) main
    window with a unique title and look the window up."""
    import win32gui
    import win32process

    marker = f"file-converter-{uuid.uuid4().hex}"
    try:
        word.Caption = marker
        hwnd = win32gui.FindWindow("OpusApp", marker)  # Word's window class
        return win32process.GetWindowThreadProcessId(hwnd)[1] if hwnd else None
    except Exception:
        return None  # no watchdog for this instance; conversion still works


def _describe(error, fallback: str) -> str:
    """Pull Word's own error message out of a COM error, if it has one."""
    excepinfo = getattr(error, "excepinfo", None)
    message = excepinfo[2] if excepinfo and excepinfo[2] else None
    if message and "password" in message.lower():
        return "File is password-protected"
    return " ".join(message.split()) if message else fallback


BACKENDS = [WordToPdf]


def backend_for(path: Path):
    """Return the backend class that handles this file, or None."""
    ext = path.suffix.lower()
    for backend in BACKENDS:
        if ext in backend.input_exts:
            return backend
    return None


def supported_exts() -> list[str]:
    return sorted(ext for backend in BACKENDS for ext in backend.input_exts)

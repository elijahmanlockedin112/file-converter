# File Converter

Convert Word documents to PDF **on your own computer**, as many as you like at once.
No uploads, no daily limits, no waiting in a queue.

It uses your installed copy of Microsoft Word to do the converting, so each PDF looks
exactly like it would if you opened the document and chose *Save as PDF* yourself.

Right now it converts `.docx` and `.doc` to `.pdf`. It's built so more formats can be
added later (see [Adding a format](#adding-a-format)).

## Requirements

- Windows
- Microsoft Word (desktop version, installed and activated)
- Python 3.10 or newer

## Setup

```bash
pip install -r requirements.txt
```

## Using the app

Double-click **`File Converter.cmd`**.

1. Drag Word files or whole folders into the window, or use **Add files…** / **Add folder…**.
2. Choose where the PDFs go: next to each original, or all into one folder.
3. Click **Convert to PDF**.

Each file shows its result as it finishes. Double-click a converted row to open its PDF.
You can also drop files straight onto the `File Converter.cmd` icon to open the app with
them already added.

## Using the command line

```bash
python -m file_converter "C:\path\to\essay.docx"
python -m file_converter "C:\path\to\folder" --recursive --out "C:\path\to\pdfs"
```

| Option | What it does |
| --- | --- |
| `-o`, `--out FOLDER` | Save PDFs in this folder (default: next to each original) |
| `-r`, `--recursive` | Also convert files in subfolders of any folder you give |
| `--overwrite` | Replace PDFs that already exist (default: skip them) |

Run `python -m file_converter` with no arguments to open the app window instead.

## How it behaves

- **Existing PDFs are skipped** unless you choose to replace them, so you can rerun a
  batch after fixing a few files and only the missing ones get converted.
- **Subfolders are kept.** When you convert a folder into an output folder, its
  subfolder layout is recreated there.
- **Name clashes are handled.** If `report.doc` and `report.docx` would both become
  `report.pdf`, the second one is saved as `report (2).pdf`.
- **One bad file doesn't stop the batch.** Corrupt and password-protected files are
  reported, and the rest carry on.
- **Word can't get stuck.** If Word spends more than 5 minutes on one file (say, on a
  hidden dialog), that file is marked as failed and Word restarts for the next one.
- **Long paths work.** Word can't open paths over 255 characters, so those files are
  converted through a short temporary copy.
- **It's safe to run with Word open.** Conversion uses its own hidden copy of Word, and
  macros in the documents are never run.

## Adding a format

Converters live in [`file_converter/backends.py`](file_converter/backends.py). A backend
is a class with `input_exts`, `output_ext`, and a `convert(src, dst)` method, used as a
context manager so it can keep something expensive (like Word) open for the whole batch.
Add the class to `BACKENDS` and both the app and the command line will pick it up.

## Running the tests

```bash
python -m unittest
```

The tests use a fake converter, so they don't need Word.

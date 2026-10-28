# pdf2word

Convert PDFs into editable documents with Microsoft Word's PDF reflow engine.

`pdf2word.py` is a single-file Python library and CLI. It keeps a pool of Word
instances warm in long-lived PowerShell workers, so a batch converts several
documents in parallel instead of paying Word's cold-start cost for every file.

The original `Convert-PdfToWord.ps1` command-line experience is preserved: it is now
a thin wrapper over the same shared implementation.

## Requirements

- Windows 10/11
- Microsoft Word (desktop, not the web version)
- Windows PowerShell 5.1+ (PowerShell 7 also works)
- Python 3.9+ (standard library only; no `pywin32`, no pip installs)

Everything lives next to the scripts. `pdf2word.py` expects `PdfToWord.Core.ps1`
and `PdfToWordWorker.ps1` in the same folder.

## Quick start

```powershell
python pdf2word.py doctor                     # check Word, PowerShell, scripts
python pdf2word.py convert report.pdf         # -> report.docx
python pdf2word.py convert report.pdf --format html
python pdf2word.py convert .\Input -r --out .\out --format docx
python pdf2word.py convert .\inbox -r --out .\out --format pdf --force
python pdf2word.py bench .\sample.pdf --copies 48 --workers 2 4
```

Inputs can be files and/or folders. A folder's **contents** keep their relative
tree, without repeating the input folder name: `Input\blockchain\NP 2025-32\NP 2025-32.pdf`
becomes `out\blockchain\NP 2025-32\NP 2025-32.docx`. Use `--base-dir` on the parent
to keep the `Input` level in the output. Folders are searched non-recursively unless
`-r/--recursive` is given. Use `--json` for one machine-readable object per result,
`--quiet` to print only failures, and `--progress` for an in-place progress line in
a terminal.

### CLI reference

```
pdf2word convert INPUT... [options]

  -o, --out PATH          output folder, or a file for a single-file input
  --format FORMAT         docx (default), doc, rtf, txt, odt, html, htmlfull, pdf
  --workers N             parallel Word instances (default: min(4, cpu_count - 1))
  --image-format          png (default) or original, for HTML images
  --images-folder NAME    HTML image folder name (default: source_folder)
  --timeout SECONDS       per-document watchdog, default 120
  --retries N             retries after a failure, default 1
  -f, --force             overwrite existing output
  --layout MODE           auto (default), mirror, subfolders, or flat
  --base-dir PATH         root for mirror paths; parent of input keeps its name
  --flat                  alias for --layout flat
  --fail-fast             stop taking new work after the first failure
  -r, --recursive         expand input folders recursively
  --json                  one JSON object per result
  -q, --quiet             only report failures
  --progress              in-place progress line (TTY only; not with --json/--quiet)
  -v, --verbose           debug logging
```

For a single **file** input, `--out` is treated as a directory unless it already
ends with the format's extension (`--out .\out` -> `out\report.docx`,
`--out .\out\report.docx` -> that exact file). For a **folder** input, `--out` must
be a directory.

### CLI examples

```powershell
# Mirror the contents of the input folder
python pdf2word.py convert "C:\Input" -r --out "C:\Converted" --format docx

# Keep the "Input" level too (base_dir is the parent)
python pdf2word.py convert "C:\Input" -r --out "C:\Converted" `
    --base-dir "C:\"

# One file (flat output)
python pdf2word.py convert "C:\Input\blockchain\NP 2025-32\NP 2025-32.pdf" --out "C:\Converted"

# Folder of files, flat, with 3 retries and a 5-minute watchdog
python pdf2word.py convert "C:\Inbox" -r --out "C:\Converted" `
    --layout flat --retries 3 --timeout 300

# HTML: NP 2025-32.html + source_folder in blockchain\NP 2025-32
python pdf2word.py convert "C:\Input" -r --out "C:\Converted" --format html

# Folders are recursive only with -r; a plain folder uses its direct PDFs
python pdf2word.py convert "C:\Inbox" --out "C:\Converted"

# In-place progress line with percentage, rate and ETA (terminal only)
python pdf2word.py convert "C:\Input" -r --out "C:\Converted" --progress

# One JSON object per document (including validation failures), no summary
python pdf2word.py convert "C:\Input" -r --out "C:\Converted" --json
```

## Library

```python
from pdf2word import PdfConverter

with PdfConverter(format="docx", workers=4, timeout=120, retries=1) as conv:
    one = conv.convert("report.pdf", out="out/report.docx", force=True)
    print(one.output, one.pages, one.size_kb)

    batch = conv.convert_many("Input", out_dir="out", recursive=True, force=True)
    print(f"{len(batch.ok)} converted, {len(batch.failed)} failed")
    for failure in batch.failed:
        print(failure.source, failure.error)
```

`convert_many` accepts a single path or any iterable of paths. Items may be PDF
files, folders, or a mix. A folder's contents are mirrored under `out_dir`; use
`recursive=True` to include nested folders. Pass `base_dir` if you want the input
folder's own name included in the output.

### `PdfConverter(...)` options

| Option | Default | Meaning |
| --- | --- | --- |
| `format` | `"docx"` | default output format |
| `workers` | `min(4, cpu_count - 1)` | maximum parallel Word instances |
| `images_folder` | `"source_folder"` | image folder for HTML output |
| `image_format` | `"png"` | `"png"` or `"original"` for HTML images |
| `timeout` | `120` | per-document watchdog, in seconds (0 disables it) |
| `retries` | `1` | retries after the first failure; total attempts = `retries + 1` |
| `recycle_after` | `50` | jobs per Word instance before a clean restart |
| `force` | `False` | default overwrite policy |
| `powershell` | auto-detect | path to `powershell.exe` or `pwsh.exe` |
| `script_dir` | next to `pdf2word.py` | folder containing the `.ps1` files |
| `startup_timeout` | `180.0` | seconds allowed for a worker's Word to start |

All options can be used without a `with` block, in which case call `close()` when
finished. Workers are also shut down by a finalizer and by process exit.

### `convert(...)`

```python
conv.convert(
    source,              # PDF file
    out=None,            # file or folder; None = next to the source
    *,
    format=None,         # overrides PdfConverter(format=...)
    force=None,          # overrides PdfConverter(force=...)
    timeout=None,        # per-document watchdog
    retries=None,        # retries after the first failure
    on_event=None,       # callback(ProgressEvent): started/completed/finished
)
```

Returns a `ConversionResult`. Output stays flat (no hierarchy) for a single file.
If `out` is a folder, the file name is `<source stem><format extension>`.

### `convert_many(...)`

```python
conv.convert_many(
    sources,             # a path, or an iterable of files/folders
    out_dir=None,        # output root; None = beside each source
    *,
    format=None,
    force=None,
    layout="auto",       # auto | mirror | subfolders | flat
    recursive=False,     # expand folders recursively
    base_dir=None,       # root used to compute mirror relative paths
    on_progress=None,    # callback(ConversionResult | ConversionFailure)
    on_event=None,       # callback(ProgressEvent) with done/total and ETA
    fail_fast=False,     # stop taking new work after the first failure
    timeout=None,
    retries=None,
)
```

Returns a `BatchResult`.

| Option | Meaning |
| --- | --- |
| `sources` | PDF files and/or folders; a single path is accepted |
| `out_dir` | output root; created on demand; omit to write beside sources |
| `layout` | see [Output layouts](#output-layouts) |
| `recursive` | when a source is a folder, include nested folders |
| `base_dir` | root for relative paths; useful with explicit file lists |
| `on_progress` | called for each success and each failure (including validation failures) |
| `on_event` | called with a `ProgressEvent` for every phase, including start and finish |
| `fail_fast` | stop starting new jobs after the first failure |
| `timeout` | per-document watchdog, overrides the constructor |
| `retries` | retries per document, overrides the constructor |

Both callbacks run from worker threads and are serialised: printing from them is
safe. Results are returned in input order; callbacks arrive in completion order.
`on_progress` is the historical one-argument callback and remains supported.

### Output layouts

**`auto`** (default) mirrors the input contents when a folder or `base_dir` is
involved. For a plain list of files it falls back to `subfolders`, preserving the
previous behaviour.

**`mirror`** preserves the hierarchy *below* the input root; the input folder itself
is not repeated:

```
Input/blockchain/NP 2025-32/NP 2025-32.pdf

out/
  blockchain/
    NP 2025-32/
      NP 2025-32.docx
```

With `base_dir`, paths are relative to that root. Set it to the parent to include the
input folder name:

```python
# run from the folder that contains Input
conv.convert_many("Input", out_dir="out", recursive=True, base_dir=".")
# out/Input/blockchain/NP 2025-32/NP 2025-32.docx
```

For an explicit file list, pass `base_dir` to get a mirror:

```python
conv.convert_many(
    ["Input/blockchain/NP 2025-32/NP 2025-32.pdf"],
    out_dir="out",
    base_dir="Input",
)
# out/blockchain/NP 2025-32/NP 2025-32.docx
```

HTML is written directly in the mirrored folder, next to `source_folder`:

```
out/blockchain/NP 2025-32/NP 2025-32.html
out/blockchain/NP 2025-32/source_folder/image001.png
```

If one folder contains several HTML documents, each gets its own subfolder
(`.../NP 2025-32/NP 2025-32.html` style) so the fixed `source_folder` names do not
collide. Word's localized companion folder (`_files`, `_fichiers`, `_Dateien`, ...)
is detected and renamed to `source_folder` automatically.

**`subfolders`** gives every document its own folder, ignoring the input tree:

```
out/NFXXX/NFXXX.docx
out/NFYYY/NFYYY.docx
```

**`flat`** writes everything directly into `out_dir`:

```
out/NFXXX.docx
out/NFYYY.docx
```

Duplicate names are refused. A flat HTML batch with more than one document is also
refused, because all documents would share one `source_folder`.

### Results

`BatchResult` contains:

| Member | Meaning |
| --- | --- |
| `.ok` | `list[ConversionResult]` |
| `.failed` | `list[ConversionFailure]` |
| `.total` | successes + failures |
| `.success` | `True` when there are no failures |
| `len(batch)` | number of successes |
| `list(batch)` | iterates over successes |
| `.raise_for_failures()` | raises `PdfToWordError` with a summary |

`ConversionResult` fields: `source`, `output`, `format`, `pages`, `size_kb`,
`seconds`, `attempts`, `companion_files`, `pdf_warning_set`.

`ConversionFailure` fields: `source`, `error`, `error_type`, `attempts`.

`ProgressEvent` fields:

| Field | Meaning |
| --- | --- |
| `phase` | `started`, `completed`, `failed`, or `finished` |
| `index` | 1-based position in the expanded batch (0 on `finished`) |
| `done` | documents finished so far (including failures) |
| `total` | documents in the batch, including pre-flight failures |
| `source` | current PDF ( `None` on `finished` ) |
| `output` | destination, when known |
| `result` | `ConversionResult` on `completed` |
| `failure` | `ConversionFailure` on `failed` |
| `elapsed` | seconds since the batch started |
| `eta` | estimated seconds remaining, or `None` |
| `workers` | worker count used for the batch |
| `cancelled` | jobs never started because of `fail_fast` (on `finished`) |
| `percent` | property: `100 * done / total` |

Example:

```python
def on_event(event):
    if event.phase == "started":
        print(f"[{event.done}/{event.total}] start {event.source.name}")
    elif event.phase == "completed":
        print(f"[{event.done}/{event.total}] ok    {event.result.output.name}")
    elif event.phase == "failed":
        print(f"[{event.done}/{event.total}] ERR   {event.failure.error}")
    elif event.phase == "finished":
        eta = "cancelled" if event.cancelled else "done"
        print(f"{event.percent:.0f}% {eta} in {event.elapsed:.1f}s")

with PdfConverter() as conv:
    conv.convert_many("Input", out_dir="out", recursive=True, on_event=on_event)
```

In a batch, a bad document (missing file, non-PDF, existing output without
`force`) becomes an entry in `.failed`; the other documents still convert. These
pre-flight failures are also delivered to `on_progress` and `on_event` (phase
`failed`), and the CLI prints them (or emits them as JSON). Invalid call arguments
(unknown format, bad layout, flat HTML batch) raise `ValidationError` immediately.
Word failures raise `WordError`, watchdog expirations raise
`ConversionTimeoutError`, and worker startup/crash problems raise `WorkerError`. All
inherit `PdfToWordError`.

### Retries, timeouts and RPC failures

`timeout` is enforced by a watchdog around each document. `retries=N` means up to
`N + 1` attempts; set it at the constructor or per call:

```python
with PdfConverter(timeout=300, retries=3) as conv:
    batch = conv.convert_many("Input", out_dir="out", recursive=True)

    # or just for this batch
    batch = conv.convert_many("Other", out_dir="out", timeout=600, retries=5)
```

```powershell
python pdf2word.py convert "C:\Input" -r --out "C:\Converted" `
    --timeout 300 --retries 3
```

On `0x800706BE`, `0x800706BA`, `0x80010108` and similar COM/RPC errors, the library
kills the affected worker, starts a clean PowerShell process after a 10-second
pause, and retries the document. After any other Word failure the worker replaces
its Word instance before the next job. A worker also checks that its Word process
is still alive between jobs, so a Word crash does not poison the rest of the batch.

### Logging

Progress callbacks are covered above (`on_event`, `ProgressEvent`); this section
describes the diagnostic logging channel and the CLI output modes.

The library uses the standard `logging` module and never prints on its own:

- logger name: `pdf2word`;
- no handler by default (`NullHandler`);
- `INFO`: batch start, per-document start/end, batch summary;
- `DEBUG`: worker startup, worker stderr (Word ready/recycle), long-path staging;
- `WARNING`: failed documents, pre-flight failures, worker restarts, callback errors.

Configure it in the host application:

```python
import logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
```

To write to a file without touching the root logger:

```python
import logging

logger = logging.getLogger("pdf2word")
logger.setLevel(logging.DEBUG)
handler = logging.FileHandler("pdf2word.log", encoding="utf-8")
handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
logger.addHandler(handler)
```

`convert_many` logs `[i/N] starting ...`, `[i/N] converted ...` and a final
`batch finished: ...` line. `convert` logs one `converting`/`converted` pair.

CLI output modes:

- **default**: one line per result, prefixed `[done/total]` for batches;
- **`--progress`**: updates one line in place with percentage, files/min and ETA
  (TTY only; ignored with `--json`/`--quiet`);
- **`--json`**: one JSON object per completed or failed document, including
  validation failures; no summary and no progress bar;
- **`--quiet`**: only failures;
- **`-v`**: DEBUG logs on stderr.

Results go to stdout; warnings, errors and progress go to stderr.

## Long paths

Windows paths longer than the classic 260-character limit are handled
automatically by the Python library and CLI:

- when the source, the destination, or an HTML image path is at least
  250 characters, the document is copied to a short `%TEMP%\pdf2word-stage-*`
  folder, converted there, then moved to the final destination using the
  `\\?\` extended-length path prefix;
- Word therefore never sees a long path, and the final file (and HTML
  `source_folder`) lands where you asked;
- staging folders are removed at the end of each job;
- a single file or folder **name** is still limited to 255 characters by Windows,
  and that is reported as a validation error before conversion.

This applies to `pdf2word.py` and `python pdf2word.py convert`. The legacy
`Convert-PdfToWord.ps1` wrapper does not stage; prefer the Python CLI for very
long paths.

## How the pool works

- Each worker is a long-lived PowerShell process that creates **its own** Word
  instance through COM activation and keeps it open.
- A named mutex serialises just the COM activation and process-id detection, so
  several workers can start at once without claiming each other's Word.
- A watchdog per document terminates only the worker's own Word if it stalls.
  Python then retries the document with a fresh worker (default: once).
- A worker recycles its Word every 50 jobs, and after any failure. After a
  force-kill it waits at least 10 seconds, because repeated force-kills leave the
  COM server unable to activate cleanly.
- The library never attaches to, quits, or kills a Word instance it did not create.
- The source PDF is only read; the library never saves over it. Existing output is
  refused unless `force=True` / `--force`.

The CLI keeps its original flags (`-KeepOpen`, `-ShowWord`, `-Attempts`,
`-TimeoutSeconds`, etc.) and its behaviour of attaching to a running Word when one
is available.

### Worker count

The default is `min(4, cpu_count - 1)`; this machine reports 4. For batches the
library starts fewer workers unless you passed `workers=` explicitly:

```
workers = min(default, max(1, ceil(jobs / 12)))
```

So 16 files use 2 workers, and 37+ files use 4. The reason is Word's 5-15 s
cold-start per instance: on small batches, starting four Words costs more than it
saves. An explicit `workers=N` is always honoured. If your machine is unstable
with four Words (antivirus, OneDrive, low RAM), use `workers=2`.

## Benchmarks

One-page PDF with a table, converted to `docx`, on an i7-1365U (2 performance +
8 efficiency cores, 12 logical, 15.6 GB RAM, NVMe). Times include worker startup.

| Files | Workers | Time | Files/min | Peak Word RAM |
| ---: | ---: | ---: | ---: | ---: |
| 48 | 1 | 69.6 s | 41.4 | 404 MB |
| 48 | 2 | 53.7 s | 53.7 | 696 MB |
| 48 | 4 | 46.1 s | 62.5 | 1263 MB |
| 48 | 6 | 54.9 s | 52.4 | 1890 MB |
| 96 | 2 | 96.4 s | 59.8 | 795 MB |
| 96 | 4 | **67.6 s** | **85.2** | 1371 MB |

For 16 files the adaptive default (2 workers) finished in 22.0 s; forcing 1, 2, or
4 workers took 25.8-30.5 s. The crossover where a fourth worker starts to pay off
is around 40 files.

More workers stop helping once Word instances compete for the two fast cores:
six workers were slower than four, and each job's average time rose from 1.3 s
with one worker to 3.5 s with six. Reproduce on your own documents with:

```powershell
python pdf2word.py bench .\sample.pdf --copies 48 --workers 1 2 4 --gap 10
```

## Limitations

- Word does no OCR. Scanned PDFs become documents full of pictures with no
  editable text.
- Layout is reflowed, not reproduced. Columns, floats and complex tables move.
- Encrypted PDFs fail.
- HTML uses one fixed image folder name (`source_folder`). Mirror and subfolder
  layouts give each document its own folder when needed; a flat HTML batch with
  several documents is refused.
- Word automation occasionally stalls; the watchdog and retry exist because of it.
- The first Word start after a reboot, Office update, or antivirus scan can be slow;
  `--timeout` raises the per-document limit.

## Troubleshooting

```
python pdf2word.py doctor
```

This prints the platform, CPU/RAM, script paths, PowerShell version, running Word
instances, then starts one worker and reports the Word version and PIDs.

- **`0x800706BE` (RPC call failed) / `0x800706BA` (server unavailable)**: Word
  crashed or the COM channel died. The library restarts the worker and retries;
  raise `retries` (for example `--retries 3`) and, if it keeps happening, lower
  `--workers` to 2. Long paths are already staged automatically, so they no longer
  cause this. Check antivirus/OneDrive too.
- **Word prompt on open**: the scripts set
  `HKCU\Software\Microsoft\Office\<ver>\Word\Options\DisableConvertPdfWarning = 1`
  so the modal "convert this PDF?" dialog cannot stall an invisible instance.
- **Execution policy**: the library launches PowerShell with
  `-ExecutionPolicy Bypass`, so no machine-wide change is needed.
- **No Word left behind**: every test ends with zero `WINWORD.EXE`. If you see one,
  check `python pdf2word.py doctor` to tell a user-launched Word from a worker's.
- **A removed Word PID between jobs**: the worker now detects it and starts a fresh
  Word before opening the next document.

## Files

| File | Role |
| --- | --- |
| `pdf2word.py` | library + CLI (this is the deliverable) |
| `PdfToWord.Core.ps1` | shared conversion, validation, HTML post-processing |
| `PdfToWordWorker.ps1` | long-lived JSON-lines worker owning one Word |
| `Convert-PdfToWord.ps1` | original CLI, thin wrapper over the core |

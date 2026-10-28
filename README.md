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
python pdf2word.py convert .\inbox --out .\out --format docx
python pdf2word.py convert .\inbox -r --out .\out --format pdf --force
python pdf2word.py bench .\sample.pdf --copies 48 --workers 2 4
```

Inputs can be files and/or folders; folders are searched non-recursively unless
`-r/--recursive` is given. Use `--json` for one machine-readable object per result
and `--quiet` to print only failures.

### CLI reference

```
pdf2word convert INPUT... [options]

  -o, --out PATH          output folder, or a file for a single input
  --format FORMAT         docx (default), doc, rtf, txt, odt, html, htmlfull, pdf
  --workers N             parallel Word instances (default: min(4, cpu_count - 1))
  --image-format          png (default) or original, for HTML images
  --images-folder NAME    HTML image folder name (default: source_folder)
  --timeout SECONDS       per-document watchdog, default 120
  --retries N             retries after a failure, default 1
  -f, --force             overwrite existing output
  --flat                  batch: one folder instead of one subfolder per document
  --fail-fast             stop taking new work after the first failure
  -r, --recursive         recurse into input folders
  --json                  one JSON object per result
  -q, --quiet             only report failures
  -v, --verbose           debug logging
```

For a single input, `--out` is treated as a directory unless it already ends with
the format's extension (`--out .\out` -> `out\report.docx`,
`--out .\out\report.docx` -> that exact file).

## Library

```python
from pdf2word import PdfConverter

with PdfConverter(format="html", workers=4) as conv:
    one = conv.convert("report.pdf", out="out/report.html", force=True)
    print(one.output, one.pages, one.size_kb)

    batch = conv.convert_many("inbox", out_dir="out", force=True)
    print(f"{len(batch.ok)} converted, {len(batch.failed)} failed")
    for failure in batch.failed:
        print(failure.source, failure.error)
```

`convert_many` accepts any iterable of paths. Progress reporting and early stop:

```python
batch = conv.convert_many(
    pdfs,
    out_dir="out",
    on_progress=lambda item: print(item),
    fail_fast=False,
)
batch.raise_for_failures()   # raises PdfToWordError with a summary
```

Useful members:

| Member | Purpose |
| --- | --- |
| `convert(source, out=None, format=..., force=...)` | one PDF, flat output |
| `convert_many(sources, out_dir=..., layout="subfolders")` | parallel batch |
| `ping()` | start the first worker and return diagnostics |
| `word_pids()` | Word processes owned by this converter |
| `close()` | quit workers and their Word instances |
| `BatchResult.ok / failed / total / success` | batch outcome |

Without a `with` block, call `close()` when finished. Workers are also shut down by
a finalizer and by process exit; the worker exits on stdin EOF if Python is killed.

### Batch output layout

With `out_dir`, every document gets its own subfolder:

```
out/
  report-a/
    report-a.html
    source_folder/
      image001.png
  report-b/
    report-b.html
    source_folder/
```

This exists because the HTML image folder name is fixed (`source_folder`). A flat
HTML batch would make every document write into the same image folder and overwrite
the previous document's pictures, so that combination is refused. `layout="flat"`
is fine for all non-HTML formats. Batches converted without `out_dir` write beside
their sources, with the same one-folder-per-HTML-document rule.

## Formats

| `--format` | Extension | Notes |
| --- | --- | --- |
| `docx` | `.docx` | default |
| `doc` | `.doc` | legacy Word 97-2003 |
| `rtf` | `.rtf` | |
| `txt` | `.txt` | UTF-8 (Word writes a BOM) |
| `odt` | `.odt` | |
| `html` | `.html` | filtered HTML, UTF-8, images in `source_folder` |
| `htmlfull` | `.html` | full HTML, UTF-8, images in `source_folder` |
| `pdf` | `.pdf` | re-saved PDF; never overwrites the source |

HTML images are PNG by default (`--image-format original` keeps Word's choice).
Word's `AllowPNG` handles most pictures; a post-pass converts anything left over
(JPEG, GIF, BMP, TIFF, EMF/WMF) and repairs the links. Spaces in output names are
handled even though Word URL-encodes the image folder reference.

## How the pool works

- Each worker is a long-lived PowerShell process that creates **its own** Word
  instance through COM activation and keeps it open.
- A named mutex serialises just the COM activation and process-id detection, so
  several workers can start at once without claiming each other's Word.
- A watchdog per document terminates only the worker's own Word if it stalls.
  Python then retries the document once (default) with a fresh Word.
- A worker recycles its Word every 50 jobs, and after any failure. After a
  force-kill it waits at least 10 seconds, because repeated force-kills leave the
  COM server unable to activate cleanly.
- The library never attaches to, quits, or kills a Word instance it did not create.
- The source PDF is opened read-only and is never modified or overwritten.
- Existing output is refused unless `force=True` / `--force`.

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
saves. An explicit `workers=N` is always honoured.

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
- Word automation occasionally stalls; the watchdog and retry exist because of it.
- The first Word start after a reboot, Office update, or antivirus scan can be slow;
  `--timeout` raises the per-document limit.

## Troubleshooting

```
python pdf2word.py doctor
```

This prints the platform, CPU/RAM, script paths, PowerShell version, running Word
instances, then starts one worker and reports the Word version and PIDs.

- **Word prompt on open**: the scripts set
  `HKCU\Software\Microsoft\Office\<ver>\Word\Options\DisableConvertPdfWarning = 1`
  so the modal "convert this PDF?" dialog cannot stall an invisible instance.
- **Execution policy**: the library launches PowerShell with
  `-ExecutionPolicy Bypass`, so no machine-wide change is needed.
- **No Word left behind**: every test ends with zero `WINWORD.EXE`. If you see one,
  check `python pdf2word.py doctor` to tell a user-launched Word from a worker's.

## Files

| File | Role |
| --- | --- |
| `pdf2word.py` | library + CLI (this is the deliverable) |
| `PdfToWord.Core.ps1` | shared conversion, validation, HTML post-processing |
| `PdfToWordWorker.ps1` | long-lived JSON-lines worker owning one Word |
| `Convert-PdfToWord.ps1` | original CLI, thin wrapper over the core |

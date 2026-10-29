# pdf2word tests

No Python dependencies: the suite uses `unittest` from the standard library.
The PowerShell post-processing tests use Pester (3.4+ is already available on this
machine). Microsoft Word is only needed for the opt-in integration tests.

## Quick commands

```powershell
# Fast unit tests: no Word, no worker, no PowerShell
python -m unittest discover -s tests/unit -t . -v

# Everything that does not need Word (integration classes are skipped)
python -m unittest discover -s tests -t . -v

# Real Word tests
$env:PDF2WORD_RUN_WORD_TESTS = "1"
python -m unittest discover -s tests -t . -v

# Add the slow stress test (48 files, 4 workers)
$env:PDF2WORD_SLOW_TESTS = "1"
python -m unittest tests.integration.test_stress -v

# Heavy PDF (the file must be a PDF)
$env:PDF2WORD_HEAVY_PDF = "C:\path\to\heavy.pdf"
# or copy it to tests/fixtures/heavy.pdf
python -m unittest tests.integration.test_heavy -v

# PowerShell post-processing tests
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Invoke-Pester -Script .\tests\ps\Complete-PdfToWordOutput.Tests.ps1
```

## What runs where

| Folder | Needs Word | Purpose |
| --- | --- | --- |
| `tests/unit/` | no | paths, planning, layouts, long-path publishing, retries, progress, CLI, logging, JSON |
| `tests/integration/` | yes | real conversion, hierarchy, HTML post-processing, long paths, failures, stress, heavy |
| `tests/ps/` | no | localized HTML companion folders, image conversion, `Resolve-PdfToWordJob` |

Integration tests are skipped unless `PDF2WORD_RUN_WORD_TESTS=1`. They run
serially: Word and the output files are process-wide resources.

## Fixtures and safety

- Integration tests copy `sample-with-table.pdf` into a temporary directory; the
  repository's sample is never modified.
- The heavy test reads `PDF2WORD_HEAVY_PDF` or `tests/fixtures/heavy.pdf` and is
  skipped when neither exists.
- Long-path tests create trees longer than 260 characters and clean them with a
  dedicated `remove_long_tree()` helper.
- Tests never kill a Word instance they did not start: only PIDs reported by
  `PdfConverter.word_pids()` are touched.
- No test changes the library or PowerShell implementation.

## Notes

The integration suite currently reports `ResourceWarning: unclosed file` from the
worker shutdown path. It does not fail the tests; it is a library cleanup detail
left unchanged in this test-only task.

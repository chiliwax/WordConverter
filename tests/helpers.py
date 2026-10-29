"""Shared helpers for the pdf2word test suite.

Standard library only. Nothing here starts Microsoft Word; the integration tests
do that explicitly through pdf2word.PdfConverter.
"""

from __future__ import annotations

import contextlib
import io
import os
import tempfile
import threading
from pathlib import Path
from typing import Callable, Iterable

import pdf2word
from pdf2word import (
    BatchResult,
    ConversionFailure,
    ConversionResult,
    ProgressEvent,
    ValidationError,
    WordError,
    WorkerError,
)

REPO_ROOT = Path(__file__).resolve().parents[1]
SAMPLE_PDF = REPO_ROOT / "sample-with-table.pdf"
PDF2WORD_SCRIPT = REPO_ROOT / "pdf2word.py"


# ---------------------------------------------------------------------------
# environment / opt-in switches
# ---------------------------------------------------------------------------

def word_tests_enabled() -> bool:
    return os.environ.get("PDF2WORD_RUN_WORD_TESTS", "").strip() == "1"


def slow_tests_enabled() -> bool:
    return os.environ.get("PDF2WORD_SLOW_TESTS", "").strip() == "1"


def heavy_pdf_path() -> Path | None:
    """PDF supplied by the user for heavy tests, or tests/fixtures/heavy.pdf."""
    candidates: list[Path] = []
    raw = os.environ.get("PDF2WORD_HEAVY_PDF")
    if raw:
        candidates.append(Path(raw).expanduser())
    candidates.append(REPO_ROOT / "tests" / "fixtures" / "heavy.pdf")
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return None


# ---------------------------------------------------------------------------
# jobs / fake workers / fake slots
# ---------------------------------------------------------------------------

def make_job(
    source: str | os.PathLike[str],
    output: str | os.PathLike[str],
    *,
    fmt: str = "docx",
    force: bool = False,
    timeout: int = 120,
    retries: int = 0,
    images_folder: str = "source_folder",
    image_format: str = "png",
    index: int = 1,
) -> pdf2word._Job:
    return pdf2word._Job(
        index=index,
        source=Path(source),
        output=Path(output),
        format=fmt,
        images_folder=images_folder,
        image_format=image_format,
        force=force,
        timeout=timeout,
        retries=retries,
    )


def success_payload(job: pdf2word._Job) -> dict:
    return {
        "ok": True,
        "result": {
            "source": str(job.source),
            "output": str(job.output),
            "format": job.format,
            "pages": 1,
            "size_kb": 0.0,
            "seconds": 0.0,
            "attempts": 1,
            "companion_files": [],
            "pdf_warning_set": False,
        },
    }


def error_payload(error_type: str, message: str) -> dict:
    return {"ok": False, "error": {"type": error_type, "message": message}}


class FakeWorker:
    """Minimal _Worker replacement for _Slot retry tests."""

    def __init__(self, responses: Iterable[dict | BaseException]):
        self.responses = list(responses)
        self.requests: list[dict] = []
        self.killed = False
        self.closed = False

    def request(self, payload: dict, timeout: float) -> dict:
        self.requests.append(payload)
        if not self.responses:
            raise RuntimeError("FakeWorker has no responses left")
        response = self.responses.pop(0)
        if isinstance(response, BaseException):
            raise response
        return response

    def kill(self) -> None:
        self.killed = True

    def close(self) -> None:
        self.closed = True


class FakeSlot:
    """Minimal _Slot replacement for convert()/convert_many() tests.

    It writes a tiny placeholder output (and a source_folder for HTML) so the
    library's planning, publishing and progress logic can be exercised without Word.
    """

    def __init__(
        self,
        index: int = 0,
        behavior: Callable[[pdf2word._Job], BaseException | None] | None = None,
    ) -> None:
        self.index = index
        self.worker = object()
        self._behavior = behavior
        self.calls: list[pdf2word._Job] = []
        self._lock = threading.Lock()

    def ensure(self):
        return self.worker

    def convert(self, job: pdf2word._Job) -> ConversionResult:
        with self._lock:
            self.calls.append(job)

        if self._behavior is not None:
            outcome = self._behavior(job)
            if isinstance(outcome, BaseException):
                raise outcome

        os.makedirs(pdf2word._fs_path(job.output.parent), exist_ok=True)
        with open(pdf2word._fs_path(job.output), "wb") as handle:
            handle.write(b"fake-output")

        companions: tuple[str, ...] = ()
        if job.format in ("html", "htmlfull"):
            folder = job.output.parent / job.images_folder
            os.makedirs(pdf2word._fs_path(folder), exist_ok=True)
            with open(pdf2word._fs_path(folder / "image001.png"), "wb") as handle:
                handle.write(b"fake-png")
            companions = (job.images_folder,)

        return ConversionResult(
            source=job.source,
            output=job.output,
            format=job.format,
            pages=1,
            size_kb=0.001,
            seconds=0.0,
            attempts=1,
            companion_files=companions,
        )


@contextlib.contextmanager
def patched_slots(converter: pdf2word.PdfConverter, slots: list[FakeSlot]):
    """Replace converter._slot() with a fake, without touching the library."""
    original = converter._slot
    converter._slot = lambda index: slots[index % len(slots)]
    try:
        yield
    finally:
        converter._slot = original


# ---------------------------------------------------------------------------
# CLI fakes
# ---------------------------------------------------------------------------

def fake_converter_class(
    results: Iterable[ConversionResult] = (),
    failures: Iterable[ConversionFailure] = (),
    *,
    raise_single: BaseException | None = None,
):
    """Build a PdfConverter-compatible fake for cmd_convert/main tests."""
    results = list(results)
    failures = list(failures)

    class FakeConverter:
        instances: list["FakeConverter"] = []

        def __init__(self, **kwargs):
            self.kwargs = kwargs
            self.instances.append(self)

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb):
            return False

        @staticmethod
        def _emit(on_event, event):
            if on_event is not None:
                on_event(event)

        def convert(self, source, out=None, force=None, on_event=None):
            if raise_single is not None:
                raise raise_single
            result = results[0] if results else ConversionResult(
                source=Path(source),
                output=Path(out) if out else Path(source),
                format="docx",
                pages=1,
                size_kb=0.001,
                seconds=0.0,
                attempts=1,
            )
            self._emit(on_event, ProgressEvent(
                phase="started", index=1, done=0, total=1,
                source=result.source, output=result.output, workers=1,
            ))
            self._emit(on_event, ProgressEvent(
                phase="completed", index=1, done=1, total=1,
                source=result.source, output=result.output, result=result, workers=1,
            ))
            self._emit(on_event, ProgressEvent(
                phase="finished", done=1, total=1, workers=1,
            ))
            return result

        def convert_many(self, sources, out_dir=None, on_event=None, **kwargs):
            total = len(results) + len(failures)
            done = 0
            for failure in failures:
                done += 1
                self._emit(on_event, ProgressEvent(
                    phase="failed", index=done, done=done, total=total,
                    source=failure.source, failure=failure,
                ))
            for result in results:
                done += 1
                self._emit(on_event, ProgressEvent(
                    phase="started", index=done, done=done - 1, total=total,
                    source=result.source, output=result.output,
                ))
                self._emit(on_event, ProgressEvent(
                    phase="completed", index=done, done=done, total=total,
                    source=result.source, output=result.output, result=result,
                ))
            self._emit(on_event, ProgressEvent(
                phase="finished", done=done, total=total,
            ))
            return BatchResult(ok=list(results), failed=list(failures))

    return FakeConverter


class FakeTTY(io.StringIO):
    def isatty(self) -> bool:
        return True


def run_cli(argv: list[str]):
    stdout = io.StringIO()
    stderr = io.StringIO()
    with contextlib.redirect_stdout(stdout), contextlib.redirect_stderr(stderr):
        code = pdf2word.main(argv)
    return code, stdout.getvalue(), stderr.getvalue()


# ---------------------------------------------------------------------------
# files / paths
# ---------------------------------------------------------------------------

def write_pdf(path: Path, data: bytes = b"%PDF-1.4\n") -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    return path


def make_result(
    source: str | os.PathLike[str],
    output: str | os.PathLike[str],
    fmt: str = "docx",
    **kwargs,
) -> ConversionResult:
    return ConversionResult(
        source=Path(source),
        output=Path(output),
        format=fmt,
        pages=1,
        size_kb=0.1,
        seconds=0.1,
        attempts=1,
        **kwargs,
    )


def make_failure(
    source: str | os.PathLike[str],
    error: str = "fake failure",
    error_type: str = "word",
) -> ConversionFailure:
    return ConversionFailure(source=Path(source), error=error, error_type=error_type, attempts=1)


@contextlib.contextmanager
def long_temp_dir(prefix: str = "pdf2word-test-"):
    root = Path(tempfile.mkdtemp(prefix=prefix))
    try:
        yield root
    finally:
        remove_long_tree(root)


def remove_long_tree(path: str | os.PathLike[str]) -> None:
    """Remove a tree even when its children are longer than MAX_PATH.

    shutil.rmtree() only prefixes the root; children can still exceed 260.
    """
    target = pdf2word._win_long(path)
    try:
        if os.path.isdir(target) and not os.path.islink(target):
            with os.scandir(target) as entries:
                for entry in entries:
                    remove_long_tree(entry.path)
            os.rmdir(target)
        else:
            os.remove(target)
    except FileNotFoundError:
        pass


def make_long_dir(root: Path, components: int = 5, size: int = 45) -> Path:
    path = root
    for _ in range(components):
        path = path / ("x" * size)
    os.makedirs(pdf2word._fs_path(path), exist_ok=True)
    return path


def staging_dirs() -> list[Path]:
    return list(Path(tempfile.gettempdir()).glob("pdf2word-stage-*"))

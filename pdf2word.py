"""pdf2word - convert PDFs to editable documents with Microsoft Word.

A single-file library and CLI. It drives one or more long-lived PowerShell workers,
each owning its own Microsoft Word instance, so Word stays warm between conversions
and several documents can be converted in parallel.

Requires Windows, Microsoft Word, and PowerShell 5.1+. Python 3.9+, standard library
only.

    from pdf2word import PdfConverter

    with PdfConverter(format="html") as conv:
        one = conv.convert("report.pdf")
        batch = conv.convert_many(["a.pdf", "b.pdf"], out_dir="out")
        print(len(batch.ok), len(batch.failed))

Word performs no OCR. Scanned PDFs become documents full of pictures with no editable
text. Layout is reflowed, not reproduced: columns, floats and complex tables shift.
"""

from __future__ import annotations

import argparse
import collections
import ctypes
import json
import logging
import math
import os
import platform
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import uuid
import weakref
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Sequence

__version__ = "1.0.0"

__all__ = [
    "PdfConverter",
    "ConversionResult",
    "ConversionFailure",
    "BatchResult",
    "PdfToWordError",
    "ValidationError",
    "ConversionTimeoutError",
    "WordError",
    "WorkerError",
    "FORMATS",
    "DEFAULT_IMAGES_FOLDER",
    "default_workers",
    "main",
]

# ---------------------------------------------------------------------------
# constants
# ---------------------------------------------------------------------------

#: Supported output formats and their file extensions.
FORMATS = {
    "docx": ".docx",
    "doc": ".doc",
    "rtf": ".rtf",
    "txt": ".txt",
    "odt": ".odt",
    "html": ".html",
    "htmlfull": ".html",
    "pdf": ".pdf",
}

DEFAULT_FORMAT = "docx"
DEFAULT_IMAGES_FOLDER = "source_folder"   # deliberately fixed by the user's spec
DEFAULT_IMAGE_FORMAT = "png"
DEFAULT_TIMEOUT = 120                     # seconds, per document, enforced by Word's watchdog
DEFAULT_RECYCLE_AFTER = 50                # jobs per Word instance before a clean restart
DEFAULT_RETRIES = 1                       # retries after the first failure
DEFAULT_STARTUP_TIMEOUT = 180.0           # cold Word start, four workers at once
IMAGES_FOLDER_SENTINEL = "source_folder"

_HERE = Path(__file__).resolve().parent
_WORKER_SCRIPT = _HERE / "PdfToWordWorker.ps1"
_CORE_SCRIPT = _HERE / "PdfToWord.Core.ps1"
_CLI_SCRIPT = _HERE / "Convert-PdfToWord.ps1"

_LOG = logging.getLogger("pdf2word")
_LOG.addHandler(logging.NullHandler())


def default_workers() -> int:
    """min(4, cpu_count - 1), never below 1."""
    return max(1, min(4, (os.cpu_count() or 1) - 1))


# ---------------------------------------------------------------------------
# errors
# ---------------------------------------------------------------------------

class PdfToWordError(Exception):
    """Base class for every failure raised by this library."""


class ValidationError(PdfToWordError):
    """The request itself is wrong: missing file, wrong extension, existing output."""


class ConversionTimeoutError(PdfToWordError):
    """Word did not finish before the watchdog fired."""


class WordError(PdfToWordError):
    """Word failed while converting."""


class WorkerError(PdfToWordError):
    """A worker process died or stopped responding."""


# ---------------------------------------------------------------------------
# results
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class ConversionResult:
    """One successfully converted document."""

    source: Path
    output: Path
    format: str
    pages: int
    size_kb: float
    seconds: float
    attempts: int
    companion_files: tuple[str, ...] = ()
    pdf_warning_set: bool = False

    @classmethod
    def _from_payload(cls, payload: dict) -> "ConversionResult":
        return cls(
            source=Path(payload["source"]),
            output=Path(payload["output"]),
            format=str(payload.get("format", "")),
            pages=int(payload.get("pages") or 0),
            size_kb=float(payload.get("size_kb") or 0.0),
            seconds=float(payload.get("seconds") or 0.0),
            attempts=int(payload.get("attempts") or 1),
            companion_files=tuple(payload.get("companion_files") or ()),
            pdf_warning_set=bool(payload.get("pdf_warning_set")),
        )

    def __str__(self) -> str:
        extra = f", {len(self.companion_files)} companion(s)" if self.companion_files else ""
        return (f"{self.source.name} -> {self.output} "
                f"({self.format}, {self.size_kb:.1f} KB, {self.seconds:.1f}s{extra})")


@dataclass(frozen=True)
class ConversionFailure:
    """One document that could not be converted."""

    source: Path
    error: str
    error_type: str = "word"
    attempts: int = 0

    def __str__(self) -> str:
        return f"{self.source.name}: {self.error}"


@dataclass
class BatchResult:
    """Successes and failures from convert_many()."""

    ok: list[ConversionResult] = field(default_factory=list)
    failed: list[ConversionFailure] = field(default_factory=list)

    @property
    def total(self) -> int:
        return len(self.ok) + len(self.failed)

    @property
    def success(self) -> bool:
        return not self.failed

    def __len__(self) -> int:
        return len(self.ok)

    def __iter__(self):
        return iter(self.ok)

    def raise_for_failures(self) -> None:
        if not self.failed:
            return
        details = "\n".join(f"  - {failure}" for failure in self.failed)
        raise PdfToWordError(
            f"{len(self.failed)} of {self.total} conversion(s) failed:\n{details}"
        )


# ---------------------------------------------------------------------------
# Windows process helpers (ctypes only, no pywin32)
# ---------------------------------------------------------------------------

if os.name == "nt":
    _kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    _SYNCHRONIZE = 0x00100000
    _PROCESS_TERMINATE = 0x0001
    _PROCESS_QUERY_INFORMATION = 0x0400
    _WAIT_TIMEOUT = 0x00000102
else:  # pragma: no cover - the library is Windows-only, but import must not crash
    _kernel32 = None
    _SYNCHRONIZE = _PROCESS_TERMINATE = _PROCESS_QUERY_INFORMATION = 0
    _WAIT_TIMEOUT = 0


def _process_alive(pid: int | None) -> bool:
    if os.name != "nt" or not pid:
        return False
    handle = _kernel32.OpenProcess(_SYNCHRONIZE, False, int(pid))
    if not handle:
        return False
    try:
        return _kernel32.WaitForSingleObject(handle, 0) == _WAIT_TIMEOUT
    finally:
        _kernel32.CloseHandle(handle)


def _kill_pid(pid: int | None) -> bool:
    if os.name != "nt" or not pid:
        return False
    handle = _kernel32.OpenProcess(_PROCESS_TERMINATE, False, int(pid))
    if not handle:
        return False
    try:
        return bool(_kernel32.TerminateProcess(handle, 1))
    finally:
        _kernel32.CloseHandle(handle)


class _ProcessMemoryCounters(ctypes.Structure):
    _fields_ = [
        ("cb", ctypes.c_ulong),
        ("PageFaultCount", ctypes.c_ulong),
        ("PeakWorkingSetSize", ctypes.c_size_t),
        ("WorkingSetSize", ctypes.c_size_t),
        ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPagedPoolUsage", ctypes.c_size_t),
        ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
        ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
        ("PagefileUsage", ctypes.c_size_t),
        ("PeakPagefileUsage", ctypes.c_size_t),
    ]


def _process_memory_mb(pid: int | None) -> float:
    """Working set of one process in MiB, or 0.0 when it cannot be read."""
    if os.name != "nt" or not pid:
        return 0.0
    handle = _kernel32.OpenProcess(_PROCESS_QUERY_INFORMATION, False, int(pid))
    if not handle:
        return 0.0
    try:
        counters = _ProcessMemoryCounters()
        counters.cb = ctypes.sizeof(counters)
        if not ctypes.windll.psapi.GetProcessMemoryInfo(
                handle, ctypes.byref(counters), counters.cb):
            return 0.0
        return counters.WorkingSetSize / (1024 * 1024)
    finally:
        _kernel32.CloseHandle(handle)


# ---------------------------------------------------------------------------
# worker process
# ---------------------------------------------------------------------------

class _WorkerCrashed(WorkerError):
    pass


class _WorkerTimeout(WorkerError):
    pass


@dataclass(frozen=True)
class _Job:
    index: int
    source: Path
    output: Path
    format: str
    images_folder: str
    image_format: str
    force: bool
    timeout: int
    retries: int

    def payload(self) -> dict:
        return {
            "cmd": "convert",
            "path": str(self.source),
            "out": str(self.output),
            "format": self.format,
            "images_folder": self.images_folder,
            "image_format": self.image_format,
            "force": self.force,
            "timeout": self.timeout,
        }


class _Worker:
    """One PowerShell worker process, owning one Word instance."""

    def __init__(
        self,
        *,
        index: int,
        powershell: str,
        recycle_after: int,
        default_timeout: int,
        startup_timeout: float,
        cwd: Path | None = None,
    ) -> None:
        self.index = index
        self.word_pid: int | None = None
        self.word_version: str | None = None
        self._lines: "queue.Queue[dict | None]" = queue.Queue()
        self._lock = threading.Lock()
        self._stderr: "collections.deque[str]" = collections.deque(maxlen=300)
        self._closed = False

        cmd = [
            powershell,
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", str(_WORKER_SCRIPT),
            "-RecycleAfter", str(int(recycle_after)),
            "-DefaultTimeoutSeconds", str(int(default_timeout)),
        ]
        _LOG.debug("starting worker %d: %s", index, " ".join(cmd))
        self.proc = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            cwd=str(cwd) if cwd else None,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        threading.Thread(
            target=self._pump_stdout, name=f"pdf2word-out-{index}", daemon=True
        ).start()
        threading.Thread(
            target=self._pump_stderr, name=f"pdf2word-err-{index}", daemon=True
        ).start()
        try:
            self._await_ready(startup_timeout)
        except Exception:
            # Never leave a half-started worker (or the Word it may already have
            # created) behind when startup fails.
            self.kill()
            raise

    # -- pipes -------------------------------------------------------------

    def _pump_stdout(self) -> None:
        try:
            for raw in self.proc.stdout:  # type: ignore[union-attr]
                raw = raw.strip()
                if not raw:
                    continue
                try:
                    message = json.loads(raw)
                except json.JSONDecodeError:
                    _LOG.debug("worker %d wrote non-JSON stdout: %s", self.index, raw[:300])
                    continue
                if not isinstance(message, dict):
                    continue
                # Unsolicited lifecycle events update state even while no request is
                # pending, so a recycled Word is never mistaken for a dead one.
                if message.get("type") == "word":
                    # Explicit nulls matter: after a failure the worker reports that its
                    # previous Word is gone, and a stale pid must not survive that.
                    if "word_pid" in message:
                        self.word_pid = message.get("word_pid")
                    if "word_version" in message:
                        self.word_version = message.get("word_version")
                    continue
                self._lines.put(message)
        except Exception as exc:  # pragma: no cover - only on a broken pipe
            _LOG.debug("worker %d stdout reader stopped: %s", self.index, exc)
        finally:
            self._lines.put(None)

    def _pump_stderr(self) -> None:
        try:
            for raw in self.proc.stderr:  # type: ignore[union-attr]
                line = raw.rstrip()
                if line:
                    self._stderr.append(line)
                    _LOG.debug("worker %d: %s", self.index, line)
        except Exception:  # pragma: no cover
            pass

    def _tail(self, count: int = 12) -> str:
        lines = list(self._stderr)
        return " | ".join(lines[-count:]) if lines else "no stderr output"

    def _startup_error(self, reason: str) -> str:
        return f"worker {self.index} {reason}; stderr: {self._tail()}"

    def _await_ready(self, timeout: float) -> None:
        deadline = time.monotonic() + timeout
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise WorkerError(self._startup_error(
                    f"did not become ready within {timeout:.0f}s"))
            try:
                message = self._lines.get(timeout=min(remaining, 1.0))
            except queue.Empty:
                if self.proc.poll() is not None:
                    raise WorkerError(self._startup_error(
                        f"exited during startup (code {self.proc.returncode})"))
                continue
            if message is None:
                raise WorkerError(self._startup_error(
                    "closed stdout during startup"))
            if message.get("type") == "ready":
                self.word_pid = message.get("word_pid") or self.word_pid
                self.word_version = message.get("word_version") or self.word_version
                _LOG.debug(
                    "worker %d ready (pid %s, Word %s pid %s)",
                    self.index, message.get("pid"), self.word_version, self.word_pid,
                )
                return
            if message.get("type") == "error":
                raise WorkerError(
                    f"worker {self.index} could not start: "
                    f"{message.get('message', 'unknown error')}"
                )

    # -- requests ----------------------------------------------------------

    def request(self, payload: dict, timeout: float) -> dict:
        if self._closed:
            raise _WorkerCrashed(f"worker {self.index} is closed")
        request_id = payload.get("id") or uuid.uuid4().hex
        payload = dict(payload, id=request_id)

        with self._lock:
            if self.proc.poll() is not None:
                raise _WorkerCrashed(
                    f"worker {self.index} exited (code {self.proc.returncode}); "
                    f"stderr: {self._tail()}"
                )
            try:
                self.proc.stdin.write(json.dumps(payload, ensure_ascii=False) + "\n")  # type: ignore[union-attr]
                self.proc.stdin.flush()  # type: ignore[union-attr]
            except (BrokenPipeError, OSError) as exc:
                raise _WorkerCrashed(
                    f"worker {self.index} stopped accepting requests; "
                    f"stderr: {self._tail()}"
                ) from exc

            deadline = time.monotonic() + timeout
            while True:
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    raise _WorkerTimeout(
                        f"worker {self.index} did not answer within {timeout:.0f}s"
                    )
                try:
                    message = self._lines.get(timeout=min(remaining, 0.5))
                except queue.Empty:
                    if self.proc.poll() is not None:
                        raise _WorkerCrashed(
                            f"worker {self.index} exited while handling a request "
                            f"(code {self.proc.returncode}); stderr: {self._tail()}"
                        )
                    continue
                if message is None:
                    raise _WorkerCrashed(
                        f"worker {self.index} closed stdout while handling a request; "
                        f"stderr: {self._tail()}"
                    )
                if message.get("id") != request_id:
                    _LOG.debug("worker %d: ignoring stale response %r",
                               self.index, message.get("id"))
                    continue
                return message

    # -- shutdown ----------------------------------------------------------

    def close(self) -> None:
        """Ask the worker to quit Word and exit. Never raises."""
        if self._closed:
            return
        try:
            self.request({"cmd": "shutdown"}, timeout=30.0)
        except Exception:
            pass
        self._closed = True
        try:
            self.proc.wait(timeout=20)
        except Exception:
            pass
        if self.proc.poll() is None:
            self.kill()
        elif _process_alive(self.word_pid):
            # The worker exited but its Word survived; this pid came from the worker,
            # so terminating it cannot touch anyone else's Word.
            _kill_pid(self.word_pid)
            self.word_pid = None
        else:
            self.word_pid = None

    def kill(self) -> None:
        """Force-kill the worker and the Word instance it reported starting."""
        self._closed = True
        if self.proc.poll() is None:
            try:
                self.proc.kill()
            except Exception:
                pass
        try:
            self.proc.wait(timeout=10)
        except Exception:
            pass
        if _process_alive(self.word_pid):
            _kill_pid(self.word_pid)
        self.word_pid = None


class _Slot:
    """A worker plus its restart policy, used by exactly one thread at a time."""

    def __init__(self, index: int, factory: Callable[[int], _Worker]) -> None:
        self.index = index
        self._factory = factory
        self.worker: _Worker | None = None

    def ensure(self) -> _Worker:
        if self.worker is None:
            self.worker = self._factory(self.index)
        return self.worker

    def reset(self, *, kill: bool = False) -> None:
        worker, self.worker = self.worker, None
        if worker is None:
            return
        if kill:
            worker.kill()
        else:
            worker.close()

    def convert(self, job: _Job) -> ConversionResult:
        last_error: Exception | None = None
        attempts = job.retries + 1

        for attempt in range(1, attempts + 1):
            try:
                worker = self.ensure()
                response = worker.request(job.payload(), timeout=job.timeout + 120.0)
            except _WorkerTimeout as exc:
                last_error = exc
                self.reset(kill=True)
            except _WorkerCrashed as exc:
                last_error = exc
                self.reset(kill=True)
            except Exception as exc:  # pragma: no cover - defensive
                last_error = exc
                self.reset(kill=True)
            else:
                if response.get("ok"):
                    return ConversionResult._from_payload(response["result"])

                error = response.get("error") or {}
                error_type = str(error.get("type") or "word")
                message = str(error.get("message") or "unknown error")

                if error_type == "validation":
                    raise ValidationError(message)
                if error_type == "timeout":
                    last_error = ConversionTimeoutError(message)
                elif error_type == "worker":
                    last_error = WorkerError(message)
                else:
                    last_error = WordError(message)

            if attempt < attempts:
                _LOG.warning(
                    "attempt %d/%d failed for %s: %s - retrying",
                    attempt, attempts, job.source.name, last_error,
                )
                time.sleep(min(2.0 * attempt, 5.0))

        assert last_error is not None
        raise last_error


def _error_type(exc: BaseException) -> str:
    if isinstance(exc, ValidationError):
        return "validation"
    if isinstance(exc, ConversionTimeoutError):
        return "timeout"
    if isinstance(exc, WorkerError):
        return "worker"
    return "word"


# ---------------------------------------------------------------------------
# public converter
# ---------------------------------------------------------------------------

def _shutdown_slots(slots: list["_Slot"]) -> None:
    for slot in list(slots):
        try:
            slot.reset(kill=False)
        except Exception:  # pragma: no cover - finalizer must never raise
            pass


class PdfConverter:
    """A pool of warm Word instances, driven by long-lived PowerShell workers.

    Each worker owns exactly one Word process, created by COM activation. The library
    never attaches to, quits, or kills a Word instance it did not create.

    Use as a context manager so the workers (and their Word instances) are shut down::

        with PdfConverter(workers=4, format="html") as conv:
            conv.convert_many(pdfs, out_dir="out")

    Without a ``with`` block the pool stays alive until ``close()`` or process exit.
    """

    def __init__(
        self,
        *,
        format: str = DEFAULT_FORMAT,
        workers: int | None = None,
        images_folder: str = DEFAULT_IMAGES_FOLDER,
        image_format: str = DEFAULT_IMAGE_FORMAT,
        timeout: int = DEFAULT_TIMEOUT,
        retries: int = DEFAULT_RETRIES,
        recycle_after: int = DEFAULT_RECYCLE_AFTER,
        force: bool = False,
        powershell: str | None = None,
        script_dir: str | os.PathLike[str] | None = None,
        startup_timeout: float = DEFAULT_STARTUP_TIMEOUT,
    ) -> None:
        if os.name != "nt":
            raise PdfToWordError("pdf2word requires Windows and Microsoft Word.")
        if format not in FORMATS:
            raise ValidationError(
                f"Unknown format {format!r}. Choose one of: {', '.join(FORMATS)}.")
        if image_format not in ("png", "original"):
            raise ValidationError("image_format must be 'png' or 'original'.")
        if not images_folder or (
            "/" in images_folder or "\\" in images_folder or images_folder in (".", "..")
        ):
            raise ValidationError(f"Invalid images_folder {images_folder!r}.")
        if workers is None:
            workers = default_workers()
            workers_explicit = False
        else:
            workers_explicit = True
        if int(workers) < 1:
            raise ValidationError("workers must be at least 1.")
        if int(timeout) < 0:
            raise ValidationError("timeout must be 0 or greater.")
        if int(retries) < 0:
            raise ValidationError("retries must be 0 or greater.")

        self.format = format
        self.workers = int(workers)
        self.workers_explicit = workers_explicit
        self.images_folder = images_folder
        self.image_format = image_format
        self.timeout = int(timeout)
        self.retries = int(retries)
        self.recycle_after = max(0, int(recycle_after))
        self.force = bool(force)
        self.startup_timeout = float(startup_timeout)

        base = Path(script_dir).expanduser().resolve() if script_dir else _HERE
        self.script_dir = base
        self.worker_script = base / "PdfToWordWorker.ps1"
        self.core_script = base / "PdfToWord.Core.ps1"
        for script in (self.worker_script, self.core_script):
            if not script.is_file():
                raise WorkerError(f"Required PowerShell script not found: {script}")

        self.powershell = powershell or self._find_powershell()

        self._slots: list[_Slot] = []
        self._slots_lock = threading.Lock()
        self._closed = False
        self._finalizer = weakref.finalize(self, _shutdown_slots, self._slots)

    @staticmethod
    def _find_powershell() -> str:
        for candidate in ("powershell.exe", "powershell", "pwsh.exe", "pwsh"):
            found = shutil.which(candidate)
            if found:
                return found
        raise WorkerError(
            "PowerShell was not found on PATH. Windows PowerShell 5.1 or PowerShell 7 "
            "is required.")

    # -- lifecycle ---------------------------------------------------------

    def __enter__(self) -> "PdfConverter":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    def close(self) -> None:
        """Shut down every worker and its Word instance. Safe to call twice."""
        with self._slots_lock:
            slots, self._slots = self._slots, []
            self._closed = True
        for slot in slots:
            slot.reset(kill=False)
        self._finalizer.detach()

    def __del__(self) -> None:  # pragma: no cover - interpreter shutdown path
        try:
            self.close()
        except Exception:
            pass

    # -- slot management ---------------------------------------------------

    def _make_worker(self, index: int) -> _Worker:
        return _Worker(
            index=index,
            powershell=self.powershell,
            recycle_after=self.recycle_after,
            default_timeout=self.timeout,
            startup_timeout=self.startup_timeout,
            cwd=self.script_dir,
        )

    def _slot(self, index: int) -> _Slot:
        with self._slots_lock:
            if self._closed:
                raise PdfToWordError("This converter has been closed.")
            while len(self._slots) <= index:
                idx = len(self._slots)
                self._slots.append(_Slot(idx, self._make_worker))
            return self._slots[index]

    # -- planning ----------------------------------------------------------

    @staticmethod
    def _normalise_source(source: str | os.PathLike[str]) -> Path:
        path = Path(source).expanduser()
        if not path.is_absolute():
            path = Path.cwd() / path
        return path

    def _validate_format(self, fmt: str | None) -> str:
        fmt = self.format if fmt is None else fmt
        if fmt not in FORMATS:
            raise ValidationError(
                f"Unknown format {fmt!r}. Choose one of: {', '.join(FORMATS)}.")
        return fmt

    def _plan_single(
        self,
        source: str | os.PathLike[str],
        out: str | os.PathLike[str] | None,
        *,
        format: str | None,
        force: bool | None,
        timeout: int | None,
        retries: int | None,
    ) -> _Job:
        fmt = self._validate_format(format)
        src = self._normalise_source(source)
        force = self.force if force is None else bool(force)

        if not src.is_file():
            raise ValidationError(f"Source file not found: {src}")
        if src.suffix.lower() != ".pdf":
            raise ValidationError(f"{src} is not a PDF. Only PDF input is supported.")

        extension = FORMATS[fmt]
        if out is None:
            output = src.with_suffix(extension)
        else:
            candidate = Path(out).expanduser()
            is_dir = str(out).endswith(("/", "\\")) or candidate.is_dir()
            if is_dir:
                candidate = candidate / (src.stem + extension)
            elif not candidate.suffix:
                candidate = candidate.with_suffix(extension)
            output = candidate if candidate.is_absolute() else Path.cwd() / candidate
            output = output.resolve()

        if output.suffix.lower() != extension:
            raise ValidationError(
                f"Output extension {output.suffix!r} does not match format {fmt!r} "
                f"({extension}).")
        if output == src:
            raise ValidationError(
                "Refusing to overwrite the source PDF. Choose a different output path.")
        if output.exists() and not force:
            raise ValidationError(
                f"Output already exists: {output} - pass force=True to overwrite.")

        return _Job(
            index=0,
            source=src.resolve(),
            output=output,
            format=fmt,
            images_folder=self.images_folder,
            image_format=self.image_format,
            force=force,
            timeout=self.timeout if timeout is None else int(timeout),
            retries=self.retries if retries is None else int(retries),
        )

    # -- single conversion -------------------------------------------------

    def convert(
        self,
        source: str | os.PathLike[str],
        out: str | os.PathLike[str] | None = None,
        *,
        format: str | None = None,
        force: bool | None = None,
        timeout: int | None = None,
        retries: int | None = None,
    ) -> ConversionResult:
        """Convert one PDF. Output stays flat next to ``out`` (a file or a folder)."""
        job = self._plan_single(
            source, out, format=format, force=force, timeout=timeout, retries=retries
        )
        return self._slot(0).convert(job)

    def ping(self) -> dict:
        """Start (if needed) the first worker and return its diagnostics."""
        worker = self._slot(0).ensure()
        response = worker.request({"cmd": "ping"}, timeout=60.0)
        if not response.get("ok"):
            raise WorkerError(str((response.get("error") or {}).get("message", "ping failed")))
        return response.get("result") or {}

    def word_pids(self) -> list[int]:
        """Process ids of the Word instances currently owned by this converter."""
        with self._slots_lock:
            return [slot.worker.word_pid
                    for slot in self._slots
                    if slot.worker is not None and slot.worker.word_pid]

    # -- batch conversion --------------------------------------------------

    def convert_many(
        self,
        sources: Iterable[str | os.PathLike[str]] | str | os.PathLike[str],
        out_dir: str | os.PathLike[str] | None = None,
        *,
        format: str | None = None,
        force: bool | None = None,
        layout: str = "subfolders",
        on_progress: Callable[[ConversionResult | ConversionFailure], None] | None = None,
        fail_fast: bool = False,
        timeout: int | None = None,
        retries: int | None = None,
    ) -> BatchResult:
        """Convert many PDFs in parallel.

        With ``out_dir`` the default layout gives every document its own subfolder::

            out/report-a/report-a.html
            out/report-a/source_folder/image001.png
            out/report-b/report-b.html
            out/report-b/source_folder/

        That keeps the fixed ``source_folder`` name collision-free. ``layout="flat"``
        writes ``out/<stem>.<ext>`` instead and refuses duplicate names. HTML batches
        also refuse any set of outputs that share a parent folder, since each export
        writes one ``source_folder``.
        """
        fmt = self._validate_format(format)
        force = self.force if force is None else bool(force)
        if layout not in ("subfolders", "flat"):
            raise ValidationError("layout must be 'subfolders' or 'flat'.")
        if isinstance(sources, (str, os.PathLike)):
            sources = [sources]

        base = Path(out_dir).expanduser().resolve() if out_dir is not None else None
        if base is not None and base.exists() and not base.is_dir():
            raise ValidationError(f"out_dir must be a folder, not a file: {base}")
        extension = FORMATS[fmt]

        jobs: list[_Job] = []
        failures: list[ConversionFailure] = []
        used_outputs: dict[str, Path] = {}
        used_folders: set[str] = set()
        seen_sources: set[str] = set()

        for source in sources:
            src = self._normalise_source(source)
            try:
                if not src.is_file():
                    raise ValidationError(f"Source file not found: {src}")
                if src.suffix.lower() != ".pdf":
                    raise ValidationError(f"{src} is not a PDF. Only PDF input is supported.")

                resolved = src.resolve()
                source_key = str(resolved).lower()
                if source_key in seen_sources:
                    raise ValidationError(f"Duplicate input in this batch: {resolved}")
                seen_sources.add(source_key)

                if base is None:
                    output = src.with_suffix(extension)
                    if output == src:
                        raise ValidationError(
                            "Refusing to overwrite the source PDF. Choose an out_dir.")
                elif layout == "flat":
                    output = base / (src.stem + extension)
                    key = str(output).lower()
                    if key in used_outputs:
                        raise ValidationError(
                            f"Two inputs would both write {output} "
                            f"({used_outputs[key]} and {src}); use layout='subfolders'.")
                    used_outputs[key] = src
                else:
                    folder = base / src.stem
                    counter = 2
                    while str(folder).lower() in used_folders:
                        folder = base / f"{src.stem}-{counter}"
                        counter += 1
                    used_folders.add(str(folder).lower())
                    output = folder / (folder.name + extension)

                if output.exists() and not force:
                    raise ValidationError(
                        f"Output already exists: {output} - pass force=True to overwrite.")

                jobs.append(_Job(
                    index=len(jobs) + len(failures),
                    source=src.resolve(),
                    output=output,
                    format=fmt,
                    images_folder=self.images_folder,
                    image_format=self.image_format,
                    force=force,
                    timeout=self.timeout if timeout is None else int(timeout),
                    retries=self.retries if retries is None else int(retries),
                ))
            except ValidationError as exc:
                failures.append(ConversionFailure(
                    source=src, error=str(exc), error_type="validation", attempts=0
                ))

        result = BatchResult(failed=list(failures))
        if not jobs:
            return result

        if fmt in ("html", "htmlfull") and len(jobs) > 1:
            parents = {str(job.output.parent).lower() for job in jobs}
            if len(parents) < len(jobs):
                raise ValidationError(
                    f"HTML batches need one output folder per document: every export "
                    f"writes to a single '{self.images_folder}' image folder, so "
                    f"{len(jobs)} documents in one folder would overwrite each other's "
                    "images. Pass out_dir=... (default layout='subfolders') or convert "
                    "them one at a time."
                )

        if self.workers_explicit:
            count = min(self.workers, len(jobs))
        else:
            # Word cold-starts cost 5-15 s per instance, and on a hybrid CPU more
            # instances than fast cores mostly buys startup contention. Small batches
            # therefore use fewer workers. An explicit workers=N is always honoured.
            count = min(self.workers, max(1, math.ceil(len(jobs) / 12)))
        slots = [self._slot(i) for i in range(count)]
        _LOG.debug("batch: %d job(s), using %d worker(s)", len(jobs), count)

        job_queue: "queue.Queue[_Job]" = queue.Queue()
        for job in jobs:
            job_queue.put(job)

        results: list[ConversionResult] = []
        all_failures: list[ConversionFailure] = list(failures)
        result_lock = threading.Lock()
        callback_lock = threading.Lock()
        stop = threading.Event()
        started = threading.Event()
        start_errors: list[tuple[int, Exception]] = []
        start_lock = threading.Lock()

        def report(item: ConversionResult | ConversionFailure) -> None:
            if on_progress is None:
                return
            with callback_lock:
                try:
                    on_progress(item)
                except Exception:  # a broken callback must not kill a worker
                    _LOG.exception("on_progress callback raised")

        def run(slot: _Slot) -> None:
            # Do not wait for the whole pool: the first worker to come up starts
            # converting while the others are still cold-starting.
            try:
                slot.ensure()
            except Exception as exc:  # noqa: BLE001 - collected and reported below
                with start_lock:
                    start_errors.append((slot.index, exc))
                return
            started.set()

            while not stop.is_set():
                try:
                    job = job_queue.get_nowait()
                except queue.Empty:
                    return
                try:
                    converted = slot.convert(job)
                except Exception as exc:  # noqa: BLE001 - recorded per document
                    failure = ConversionFailure(
                        source=job.source,
                        error=str(exc),
                        error_type=_error_type(exc),
                        attempts=1 if isinstance(exc, ValidationError) else job.retries + 1,
                    )
                    with result_lock:
                        all_failures.append(failure)
                    report(failure)
                    if fail_fast:
                        stop.set()
                else:
                    with result_lock:
                        results.append(converted)
                    report(converted)

        runners = [threading.Thread(target=run, args=(slot,), daemon=True, name=f"pdf2word-job-{slot.index}")
                   for slot in slots]
        for thread in runners:
            thread.start()
        for thread in runners:
            thread.join()

        if not started.is_set():
            details = "; ".join(f"worker {index}: {exc}" for index, exc in start_errors)
            raise WorkerError(f"No conversion worker could start. {details}")
        for index, exc in start_errors:
            _LOG.warning("worker %d failed to start (%s); using %d worker(s)",
                         index, exc, count - len(start_errors))

        order = {job.source: job.index for job in jobs}
        result.ok = sorted(results, key=lambda item: order.get(item.source, 0))
        result.failed = sorted(all_failures, key=lambda item: order.get(item.source, 0))
        return result


# ---------------------------------------------------------------------------
# command line
# ---------------------------------------------------------------------------

class _MemoryStatusEx(ctypes.Structure):
    _fields_ = [
        ("dwLength", ctypes.c_ulong),
        ("dwMemoryLoad", ctypes.c_ulong),
        ("ullTotalPhys", ctypes.c_ulonglong),
        ("ullAvailPhys", ctypes.c_ulonglong),
        ("ullTotalPageFile", ctypes.c_ulonglong),
        ("ullAvailPageFile", ctypes.c_ulonglong),
        ("ullTotalVirtual", ctypes.c_ulonglong),
        ("ullAvailVirtual", ctypes.c_ulonglong),
        ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
    ]


def _total_ram_gb() -> float:
    if os.name != "nt":
        return 0.0
    status = _MemoryStatusEx()
    status.dwLength = ctypes.sizeof(status)
    if not ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(status)):
        return 0.0
    return status.ullTotalPhys / (1024 ** 3)


def _running_word_pids() -> list[int]:
    if os.name != "nt":
        return []
    try:
        output = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=15,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        ).stdout
    except Exception:
        return []
    pids: list[int] = []
    for line in output.splitlines():
        parts = [part.strip('"') for part in line.split('","')]
        if len(parts) >= 2 and parts[0].upper() == "WINWORD.EXE":
            try:
                pids.append(int(parts[1]))
            except ValueError:
                pass
    return pids


def _expand_inputs(inputs: Sequence[str], recursive: bool) -> list[Path]:
    results: list[Path] = []
    for raw in inputs:
        path = Path(raw).expanduser()
        if path.is_dir():
            pattern = "**/*.pdf" if recursive else "*.pdf"
            results.extend(sorted(path.glob(pattern)))
        else:
            results.append(path)
    return results


def _configure_logging(verbose: bool) -> None:
    logger = logging.getLogger("pdf2word")
    if verbose:
        logger.setLevel(logging.DEBUG)
        if not any(isinstance(h, logging.StreamHandler) for h in logger.handlers):
            handler = logging.StreamHandler(sys.stderr)
            handler.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
            logger.addHandler(handler)
    else:
        logger.setLevel(logging.WARNING)


def _result_json(result: ConversionResult) -> str:
    return json.dumps({
        "ok": True,
        "source": str(result.source),
        "output": str(result.output),
        "format": result.format,
        "pages": result.pages,
        "size_kb": result.size_kb,
        "seconds": result.seconds,
        "attempts": result.attempts,
        "companion_files": list(result.companion_files),
        "pdf_warning_set": result.pdf_warning_set,
    }, ensure_ascii=False)


def _failure_json(failure: ConversionFailure) -> str:
    return json.dumps({
        "ok": False,
        "source": str(failure.source),
        "error": failure.error,
        "error_type": failure.error_type,
        "attempts": failure.attempts,
    }, ensure_ascii=False)


def cmd_convert(args: argparse.Namespace) -> int:
    fmt = args.format
    inputs = _expand_inputs(args.inputs, args.recursive)
    if not inputs:
        print("No PDF files matched.", file=sys.stderr)
        return 2

    converter_kwargs = dict(
        format=fmt,
        workers=args.workers,
        images_folder=args.images_folder,
        image_format=args.image_format,
        timeout=args.timeout,
        retries=args.retries,
        force=args.force,
    )

    def emit(result: ConversionResult | ConversionFailure) -> None:
        if args.json:
            line = _result_json(result) if isinstance(result, ConversionResult) else _failure_json(result)
            print(line, flush=True)
            return
        if isinstance(result, ConversionResult):
            if not args.quiet:
                print(f"  ok  {result}", flush=True)
        else:
            print(f"  ERR {failure_line(result)}", file=sys.stderr, flush=True)

    started = time.monotonic()
    try:
        with PdfConverter(**converter_kwargs) as converter:
            if len(inputs) == 1:
                source = inputs[0]
                out = None
                if args.out:
                    out_path = Path(args.out).expanduser()
                    if out_path.suffix.lower() == FORMATS[fmt]:
                        out = out_path
                    else:
                        out = out_path / (source.stem + FORMATS[fmt])
                result = converter.convert(source, out=out, force=args.force)
                emit(result)
                ok, failed = 1, []
            else:
                batch = converter.convert_many(
                    inputs,
                    out_dir=args.out,
                    layout="flat" if args.flat else "subfolders",
                    on_progress=emit,
                    fail_fast=args.fail_fast,
                )
                ok, failed = len(batch.ok), batch.failed
    except (PdfToWordError, OSError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    elapsed = time.monotonic() - started
    if not args.json:
        summary = f"{ok} converted, {len(failed)} failed in {elapsed:.1f}s"
        print(summary, file=sys.stderr if failed else sys.stdout)
    return 0 if not failed else 1


def failure_line(failure: ConversionFailure) -> str:
    return f"{failure.source} - {failure.error}"


# ---------------------------------------------------------------------------
# doctor and benchmark
# ---------------------------------------------------------------------------

def cmd_doctor(args: argparse.Namespace) -> int:
    print(f"pdf2word {__version__}")
    print(f"platform       : {platform.platform()}")
    print(f"python         : {sys.version.split()[0]}")
    print(f"cpu cores      : {os.cpu_count()}")
    print(f"memory         : {_total_ram_gb():.1f} GB")
    print(f"default workers: {default_workers()}")

    scripts = {
        "core": _CORE_SCRIPT,
        "worker": _WORKER_SCRIPT,
        "cli": _CLI_SCRIPT,
    }
    scripts_ok = True
    for label, path in scripts.items():
        exists = path.is_file()
        scripts_ok = scripts_ok and exists
        print(f"script {label:<7}: {'ok' if exists else 'MISSING'}  {path}")

    powershell = shutil.which("powershell.exe") or shutil.which("pwsh.exe") or shutil.which("powershell")
    if powershell:
        try:
            version = subprocess.run(
                [powershell, "-NoProfile", "-Command", "$PSVersionTable.PSVersion.ToString()"],
                capture_output=True, text=True, timeout=30,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            ).stdout.strip() or "unknown"
        except Exception as exc:
            version = f"error: {exc}"
        print(f"powershell     : {version}  {powershell}")
    else:
        print("powershell     : NOT FOUND")
        scripts_ok = False

    running = _running_word_pids()
    if running:
        print(f"running Word   : {len(running)} instance(s) (left untouched) "
              f"[{', '.join(map(str, running))}]")
    else:
        print("running Word   : none")

    if not scripts_ok:
        print("\nRESULT: FAILED - required files are missing.")
        return 1

    print("\nstarting one worker and checking Word COM ...")
    try:
        with PdfConverter(workers=1) as converter:
            info = converter.ping()
    except Exception as exc:
        print(f"RESULT: FAILED - {exc}")
        return 1

    print(f"Word version   : {info.get('word_version')}")
    print(f"worker pid     : {info.get('pid')}")
    print(f"Word pid       : {info.get('word_pid')}")
    print("\nRESULT: OK")
    return 0


def cmd_bench(args: argparse.Namespace) -> int:
    source = Path(args.input).expanduser().resolve()
    if not source.is_file():
        print(f"error: source file not found: {source}", file=sys.stderr)
        return 2
    if args.copies < 1:
        print("error: --copies must be at least 1", file=sys.stderr)
        return 2

    cores = max(1, os.cpu_count() or 1)
    counts = args.workers or [n for n in (1, 2, 4, 6) if n <= cores]
    counts = sorted({max(1, int(n)) for n in counts})

    workspace = Path(tempfile.mkdtemp(prefix="pdf2word-bench-"))
    copies_dir = workspace / "in"
    copies_dir.mkdir(parents=True, exist_ok=True)
    pdfs: list[Path] = []
    for index in range(args.copies):
        destination = copies_dir / f"case-{index + 1:02d}.pdf"
        shutil.copy2(source, destination)
        pdfs.append(destination)

    print(f"{args.copies} copies of {source.name}; format={args.format}; timeout={args.timeout}s")
    print(f"{'workers':>7}  {'files':>5}  {'ok':>4}  {'failed':>6}  {'seconds':>8}  "
          f"{'startup':>7}  {'avg job':>7}  {'files/min':>9}  {'peak Word RAM':>13}")
    print("-" * 84)

    try:
        for count in counts:
            out_dir = workspace / f"out-{count}"
            peak = {"mb": 0.0}
            stop = threading.Event()
            started = time.monotonic()
            first = {"at": None}
            failures = 0
            converted = 0

            def progress(_item, first=first) -> None:
                if first["at"] is None:
                    first["at"] = time.monotonic()

            try:
                with PdfConverter(
                    format=args.format, workers=count, timeout=args.timeout, force=True
                ) as converter:
                    def sample() -> None:
                        while not stop.wait(0.5):
                            pids = converter.word_pids()
                            total = sum(_process_memory_mb(pid) for pid in pids)
                            peak["mb"] = max(peak["mb"], total)

                    sampler = threading.Thread(target=sample, daemon=True)
                    sampler.start()
                    try:
                        batch = converter.convert_many(
                            pdfs, out_dir=out_dir, on_progress=progress
                        )
                        converted, failures = len(batch.ok), len(batch.failed)
                        job_seconds = [item.seconds for item in batch.ok]
                    finally:
                        stop.set()
                        sampler.join(timeout=3)
            except Exception as exc:
                print(f"{count:>7}  FAILED: {exc}", file=sys.stderr)
                continue

            elapsed = time.monotonic() - started
            rate = (converted / elapsed * 60.0) if elapsed > 0 else 0.0
            startup = (first["at"] - started) if first["at"] else 0.0
            average = (sum(job_seconds) / len(job_seconds)) if job_seconds else 0.0
            print(f"{count:>7}  {len(pdfs):>5}  {converted:>4}  {failures:>6}  "
                  f"{elapsed:>8.1f}  {startup:>7.1f}  {average:>7.2f}  {rate:>9.1f}  "
                  f"{peak['mb']:>10.0f} MB")

            # Give Word a moment to fully exit before the next round.
            time.sleep(args.gap)
    finally:
        if args.keep:
            print(f"\nkept: {workspace}")
        else:
            shutil.rmtree(workspace, ignore_errors=True)

    return 0


# ---------------------------------------------------------------------------
# entry point
# ---------------------------------------------------------------------------

def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pdf2word",
        description="Convert PDFs to editable documents with Microsoft Word.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "examples:\n"
            "  pdf2word convert report.pdf --format html\n"
            "  pdf2word convert in\\*.pdf --out out --workers 4\n"
            "  pdf2word convert in -r --out out --format docx --force\n"
            "  pdf2word doctor\n"
            "  pdf2word bench sample.pdf --copies 16 --workers 1 2 4 6\n"
        ),
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    convert = sub.add_parser("convert", help="convert one or more PDFs")
    convert.add_argument("inputs", nargs="+", help="PDF files and/or folders")
    convert.add_argument("-o", "--out", help="output folder (or file for a single input)")
    convert.add_argument("--format", choices=list(FORMATS), default=DEFAULT_FORMAT)
    convert.add_argument("--workers", type=int, default=None,
                         help=f"parallel Word instances (default {default_workers()})")
    convert.add_argument("--image-format", choices=("png", "original"), default=DEFAULT_IMAGE_FORMAT,
                         help="html/htmlfull image format (default png)")
    convert.add_argument("--images-folder", default=DEFAULT_IMAGES_FOLDER,
                         help="html/htmlfull image folder name (default source_folder)")
    convert.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT,
                         help=f"per-document watchdog in seconds (default {DEFAULT_TIMEOUT})")
    convert.add_argument("--retries", type=int, default=DEFAULT_RETRIES,
                         help=f"retries after a failure (default {DEFAULT_RETRIES})")
    convert.add_argument("-f", "--force", action="store_true", help="overwrite existing output")
    convert.add_argument("--flat", action="store_true",
                         help="batch: write all documents into --out instead of one subfolder each")
    convert.add_argument("--fail-fast", action="store_true",
                         help="stop starting new conversions after the first failure")
    convert.add_argument("-r", "--recursive", action="store_true",
                         help="when an input is a folder, search it recursively")
    convert.add_argument("--json", action="store_true", help="one JSON object per result")
    convert.add_argument("-q", "--quiet", action="store_true", help="only print failures")
    convert.add_argument("-v", "--verbose", action="store_true", help="debug logging")
    convert.set_defaults(func=cmd_convert)

    doctor = sub.add_parser("doctor", help="check Windows, PowerShell and Word")
    doctor.add_argument("-v", "--verbose", action="store_true")
    doctor.set_defaults(func=cmd_doctor)

    bench = sub.add_parser("bench", help="measure throughput at several worker counts")
    bench.add_argument("input", help="PDF to copy for the benchmark")
    bench.add_argument("--copies", type=int, default=16)
    bench.add_argument("--workers", type=int, nargs="+", default=None,
                       help="worker counts to test (default 1 2 4 6, capped by cores)")
    bench.add_argument("--format", choices=list(FORMATS), default=DEFAULT_FORMAT)
    bench.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT)
    bench.add_argument("--gap", type=float, default=10.0,
                       help="seconds to wait between rounds (default 10)")
    bench.add_argument("--keep", action="store_true", help="keep the temporary output")
    bench.add_argument("-v", "--verbose", action="store_true")
    bench.set_defaults(func=cmd_bench)

    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    _configure_logging(getattr(args, "verbose", False))
    try:
        return int(args.func(args))
    except KeyboardInterrupt:
        print("interrupted", file=sys.stderr)
        return 130


if __name__ == "__main__":
    raise SystemExit(main())

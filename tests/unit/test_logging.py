from __future__ import annotations

import contextlib
import io
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import pdf2word
from tests.helpers import (
    FakeSlot,
    FakeWorker,
    error_payload,
    make_job,
    patched_slots,
    success_payload,
    write_pdf,
)


class LoggingTests(unittest.TestCase):
    def setUp(self):
        self._temp = tempfile.TemporaryDirectory(prefix="pdf2word-log-")
        self.root = Path(self._temp.name)
        self.addCleanup(self._temp.cleanup)

    def converter(self) -> pdf2word.PdfConverter:
        converter = pdf2word.PdfConverter(workers=1)
        self.addCleanup(converter.close)
        return converter

    def test_info_logs_for_batch(self):
        source = write_pdf(self.root / "a.pdf")
        converter = self.converter()
        slot = FakeSlot()
        with patched_slots(converter, [slot]):
            with self.assertLogs("pdf2word", level="INFO") as captured:
                converter.convert_many(
                    [source], out_dir=self.root / "out", force=True,
                )
        text = "\n".join(captured.output)
        self.assertIn("batch:", text)
        self.assertIn("starting a.pdf", text)
        self.assertIn("converted a.pdf", text)
        self.assertIn("batch finished:", text)

    def test_default_library_use_is_silent(self):
        source = write_pdf(self.root / "a.pdf")
        converter = self.converter()
        slot = FakeSlot()
        stderr = io.StringIO()
        with patched_slots(converter, [slot]):
            with contextlib.redirect_stderr(stderr):
                converter.convert(source, out=self.root / "a.docx", force=True)
        self.assertEqual(stderr.getvalue(), "")

    def test_callback_exception_is_logged(self):
        source = write_pdf(self.root / "a.pdf")
        converter = self.converter()
        slot = FakeSlot()

        def broken(_event):
            raise RuntimeError("boom")

        with patched_slots(converter, [slot]):
            with self.assertLogs("pdf2word", level="ERROR") as captured:
                batch = converter.convert_many(
                    [source], out_dir=self.root / "out", force=True,
                    on_event=broken,
                )
        self.assertEqual(len(batch.ok), 1)
        self.assertTrue(any("on_event callback raised" in line for line in captured.output))

    def test_retry_is_logged_as_warning(self):
        job = make_job(self.root / "in.pdf", self.root / "out.docx", retries=1)
        worker = FakeWorker([
            error_payload("word", "temporary failure"),
            success_payload(job),
        ])
        slot = pdf2word._Slot(0, lambda _index: worker)
        with mock.patch("pdf2word.time.sleep"):
            with self.assertLogs("pdf2word", level="WARNING") as captured:
                slot.convert(job)
        self.assertTrue(any("retrying" in line for line in captured.output))


if __name__ == "__main__":
    unittest.main()

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import pdf2word
from tests.helpers import FakeSlot, patched_slots, write_pdf


class ProgressTestCase(unittest.TestCase):
    def setUp(self):
        self._temp = tempfile.TemporaryDirectory(prefix="pdf2word-progress-")
        self.root = Path(self._temp.name)
        self.addCleanup(self._temp.cleanup)

    def run_batch(self, sources, *, workers=1, behavior=None, **kwargs):
        converter = pdf2word.PdfConverter(workers=workers)
        self.addCleanup(converter.close)
        slots = [FakeSlot(i, behavior=behavior) for i in range(workers)]
        with patched_slots(converter, slots):
            batch = converter.convert_many(sources, **kwargs)
        return batch, slots

    def test_batch_progress_phases_and_counts(self):
        first = write_pdf(self.root / "a.pdf")
        second = write_pdf(self.root / "b.pdf")
        missing = self.root / "missing.pdf"
        events = []
        progress = []
        batch, _ = self.run_batch(
            [first, missing, second], out_dir=self.root / "out", force=True,
            on_event=events.append, on_progress=progress.append,
        )

        self.assertEqual(len(batch.ok), 2)
        self.assertEqual(len(batch.failed), 1)
        self.assertEqual(len(progress), 3)

        phases = [event.phase for event in events]
        self.assertEqual(phases.count("failed"), 1)
        self.assertEqual(phases.count("started"), 2)
        self.assertEqual(phases.count("completed"), 2)
        self.assertEqual(phases.count("finished"), 1)

        finished = events[-1]
        self.assertEqual(finished.phase, "finished")
        self.assertEqual(finished.done, 3)
        self.assertEqual(finished.total, 3)
        self.assertEqual(finished.cancelled, 0)
        self.assertEqual(finished.percent, 100.0)

        done_values = [event.done for event in events if event.phase in ("completed", "failed")]
        self.assertEqual(done_values, sorted(done_values))

    def test_batch_progress_preflight_failure_is_reported(self):
        source = write_pdf(self.root / "a.pdf")
        events = []
        batch, _ = self.run_batch(
            [source, self.root / "missing.pdf"], out_dir=self.root / "out",
            force=True, on_event=events.append,
        )
        failed = [event for event in events if event.phase == "failed"]
        self.assertEqual(len(failed), 1)
        self.assertEqual(failed[0].source.name, "missing.pdf")
        self.assertEqual(failed[0].done, 1)
        self.assertEqual(batch.total, 2)
        self.assertEqual(batch.failed[0].error_type, "validation")

    def test_batch_progress_fail_fast_cancelled(self):
        folder = self.root / "Input"
        write_pdf(folder / "a.pdf")
        write_pdf(folder / "b.pdf")

        def fail_first(job):
            return pdf2word.WordError("fake failure")

        events = []
        batch, slots = self.run_batch(
            folder, workers=1, out_dir=self.root / "out", force=True,
            fail_fast=True, on_event=events.append, behavior=fail_first,
        )
        self.assertEqual(len(batch.failed), 1)
        self.assertEqual(len(slots[0].calls), 1)

        finished = events[-1]
        self.assertEqual(finished.phase, "finished")
        self.assertEqual(finished.total, 2)
        self.assertEqual(finished.done, 1)
        self.assertEqual(finished.cancelled, 1)
        self.assertLess(finished.percent, 100.0)

    def test_batch_progress_callback_exception_does_not_stop_batch(self):
        source = write_pdf(self.root / "a.pdf")

        def broken(_event):
            raise RuntimeError("broken callback")

        with self.assertLogs("pdf2word", level="ERROR"):
            batch, _ = self.run_batch(
                [source], out_dir=self.root / "out", force=True, on_event=broken,
            )
        self.assertEqual(len(batch.ok), 1)

    def test_batch_on_progress_legacy_callback(self):
        source = write_pdf(self.root / "a.pdf")
        items = []
        batch, _ = self.run_batch(
            [source, self.root / "missing.pdf"], out_dir=self.root / "out",
            force=True, on_progress=items.append,
        )
        self.assertEqual(len(items), 2)
        self.assertTrue(any(isinstance(item, pdf2word.ConversionResult) for item in items))
        self.assertTrue(any(isinstance(item, pdf2word.ConversionFailure) for item in items))
        self.assertEqual(len(batch.ok) + len(batch.failed), 2)

    def test_convert_progress_success(self):
        source = write_pdf(self.root / "a.pdf")
        converter = pdf2word.PdfConverter(workers=1)
        self.addCleanup(converter.close)
        slot = FakeSlot()
        events = []
        with patched_slots(converter, [slot]):
            result = converter.convert(
                source, out=self.root / "a.docx", force=True,
                on_event=events.append,
            )
        self.assertEqual(result.output.name, "a.docx")
        self.assertEqual([event.phase for event in events],
                         ["started", "completed", "finished"])
        self.assertEqual(events[-1].done, 1)
        self.assertEqual(events[-1].total, 1)

    def test_convert_progress_failure(self):
        source = write_pdf(self.root / "a.pdf")

        def fail(_job):
            return pdf2word.WordError("fake failure")

        converter = pdf2word.PdfConverter(workers=1)
        self.addCleanup(converter.close)
        slot = FakeSlot(behavior=fail)
        events = []
        with patched_slots(converter, [slot]):
            with self.assertRaises(pdf2word.WordError):
                converter.convert(
                    source, out=self.root / "a.docx", force=True,
                    on_event=events.append,
                )
        self.assertEqual([event.phase for event in events],
                         ["started", "failed", "finished"])
        self.assertEqual(events[-1].done, 1)
        self.assertEqual(events[-1].total, 1)


if __name__ == "__main__":
    unittest.main()

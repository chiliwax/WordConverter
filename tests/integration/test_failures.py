from __future__ import annotations

import subprocess
import time
import unittest

from pdf2word import ConversionTimeoutError, PdfConverter
from tests.integration.base import WordTestCase


class FailureRecoveryTests(WordTestCase):
    def test_timeout_raises_and_leaves_no_partial_output(self):
        source = self.copy_sample()
        output = self.root / "out" / "timeout.docx"
        converter = PdfConverter(workers=1, timeout=1, retries=0)
        with converter:
            with self.assertRaises(ConversionTimeoutError):
                converter.convert(source, out=output, force=True)
        self.assertFalse(output.exists())
        self.assert_no_orphan_word(converter)

    def test_worker_kill_recovery(self):
        source = self.copy_sample()
        converter = PdfConverter(workers=1, retries=1, timeout=180)
        with converter:
            converter.convert(source, out=self.root / "out" / "first.docx", force=True)
            worker = converter._slot(0).worker
            self.assertIsNotNone(worker)
            worker.proc.kill()
            worker.proc.wait(timeout=10)
            result = converter.convert(
                source, out=self.root / "out" / "second.docx", force=True,
            )
            self.assertTrue(result.output.is_file())
        self.assert_no_orphan_word(converter)

    def test_word_kill_recovery(self):
        source = self.copy_sample()
        converter = PdfConverter(workers=1, retries=1, timeout=180)
        with converter:
            converter.convert(source, out=self.root / "out" / "first.docx", force=True)
            word_pids = converter.word_pids()
            self.assertEqual(len(word_pids), 1)
            subprocess.run(
                ["taskkill", "/PID", str(word_pids[0]), "/F"],
                capture_output=True, check=False,
            )
            time.sleep(1)
            result = converter.convert(
                source, out=self.root / "out" / "second.docx", force=True,
            )
            self.assertTrue(result.output.is_file())
        self.assert_no_orphan_word(converter)

    def test_no_orphan_word_after_close(self):
        source = self.copy_sample()
        converter = PdfConverter(workers=2)
        with converter:
            converter.convert(source, out=self.root / "out" / "one.docx", force=True)
            self.assertTrue(converter.word_pids())
        self.assertEqual(converter.word_pids(), [])

    def test_no_orphan_word_after_timeout(self):
        source = self.copy_sample()
        converter = PdfConverter(workers=1, timeout=1, retries=0)
        with converter:
            try:
                converter.convert(source, out=self.root / "out" / "x.docx", force=True)
            except ConversionTimeoutError:
                pass
        self.assertEqual(converter.word_pids(), [])


if __name__ == "__main__":
    unittest.main()

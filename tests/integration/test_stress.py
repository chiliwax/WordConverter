from __future__ import annotations

import unittest

from pdf2word import PdfConverter
from tests.integration.base import WordTestCase, slow


class StressTests(WordTestCase):
    def make_files(self, count: int):
        for index in range(1, count + 1):
            self.copy_sample(
                f"Input/Cloud/NF {index:03d}/NF{index:03d}.pdf"
            )

    def test_16_files_2_workers(self):
        self.make_files(16)
        out = self.root / "Output"
        converter = PdfConverter(workers=2, retries=2, timeout=180)
        with converter:
            batch = converter.convert_many(
                self.root / "Input", out_dir=out, recursive=True, force=True,
            )
        self.assertEqual(len(batch.ok), 16, batch.failed)
        self.assertEqual(len(batch.failed), 0)
        self.assertEqual(len(list(out.rglob("*.docx"))), 16)
        self.assert_no_orphan_word(converter)

    @slow
    def test_48_files_4_workers(self):
        self.make_files(48)
        out = self.root / "Output"
        converter = PdfConverter(workers=4, retries=2, timeout=180)
        with converter:
            batch = converter.convert_many(
                self.root / "Input", out_dir=out, recursive=True, force=True,
            )
        self.assertEqual(len(batch.ok), 48, batch.failed)
        self.assertEqual(len(batch.failed), 0)
        self.assertEqual(len(list(out.rglob("*.docx"))), 48)
        self.assert_no_orphan_word(converter)


if __name__ == "__main__":
    unittest.main()

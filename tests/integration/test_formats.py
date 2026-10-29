from __future__ import annotations

import unittest

from pdf2word import FORMATS, PdfConverter
from tests.integration.base import WordTestCase


class FormatMatrixTests(WordTestCase):
    def test_all_formats(self):
        source = self.copy_sample()
        converter = PdfConverter(workers=1, retries=0)
        with converter:
            for fmt, extension in FORMATS.items():
                output = self.root / fmt / f"sample{extension}"
                result = converter.convert(
                    source, out=output, format=fmt, force=True,
                )
                self.assertTrue(result.output.is_file(), fmt)
                self.assertGreater(result.output.stat().st_size, 0, fmt)

                data = result.output.read_bytes()
                if fmt in ("docx", "odt"):
                    self.assertTrue(data.startswith(b"PK"), fmt)
                elif fmt == "rtf":
                    self.assertTrue(data.startswith(b"{\\rtf"), fmt)
                elif fmt == "txt":
                    self.assertEqual(data[:3], b"\xef\xbb\xbf", fmt)
                elif fmt == "pdf":
                    self.assertEqual(data[:5], b"%PDF-", fmt)
                elif fmt in ("html", "htmlfull"):
                    html = result.output.read_text(encoding="utf-8")
                    self.assertIn("charset=utf-8", html.lower(), fmt)
                    self.assertIn("source_folder/", html, fmt)
                    images = result.output.parent / "source_folder"
                    self.assertTrue(images.is_dir(), fmt)
                    self.assertTrue(any(images.glob("*.png")), fmt)
                    self.assertFalse(
                        any(result.output.parent.glob("*_files")), fmt
                    )
                    self.assertFalse(
                        any(result.output.parent.glob("*_fichiers")), fmt
                    )
        self.assert_no_orphan_word(converter)

    def test_single_file_output(self):
        source = self.copy_sample("Input/blockchain/NP 2025-32/NP 2025-32.pdf")
        converter = PdfConverter(workers=1, retries=0)
        with converter:
            result = converter.convert(source, force=True)
            self.assertEqual(result.output, source.with_suffix(".docx"))
            self.assertTrue(result.output.is_file())
        self.assert_no_orphan_word(converter)


if __name__ == "__main__":
    unittest.main()

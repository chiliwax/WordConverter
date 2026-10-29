from __future__ import annotations

import unittest

from pdf2word import ConversionTimeoutError, PdfConverter
from tests.helpers import heavy_pdf_path
from tests.integration.base import WordTestCase

HEAVY_PDF = heavy_pdf_path()


@unittest.skipUnless(
    HEAVY_PDF is not None,
    "set PDF2WORD_HEAVY_PDF to a heavy PDF (or place one at tests/fixtures/heavy.pdf)",
)
class HeavyDocumentTests(WordTestCase):
    def test_heavy_pdf_to_docx(self):
        output = self.root / "out" / "heavy.docx"
        converter = PdfConverter(workers=1, retries=1, timeout=600)
        with converter:
            result = converter.convert(HEAVY_PDF, out=output, force=True)
            self.assertTrue(result.output.is_file())
            self.assertGreater(result.output.stat().st_size, 0)
            self.assertGreaterEqual(result.pages, 1)
        self.assert_no_orphan_word(converter)

    def test_heavy_pdf_to_html(self):
        output = self.root / "out" / "heavy.html"
        converter = PdfConverter(workers=1, format="html", retries=1, timeout=600)
        with converter:
            result = converter.convert(HEAVY_PDF, out=output, force=True)
            self.assertTrue(result.output.is_file())
            html = result.output.read_text(encoding="utf-8")
            self.assertIn("charset=utf-8", html.lower())
            self.assertFalse(any(result.output.parent.glob("*_fichiers")))
            self.assertFalse(any(result.output.parent.glob("*_files")))
            images = result.output.parent / "source_folder"
            if images.is_dir():
                self.assertTrue(any(images.glob("*.png")))
        self.assert_no_orphan_word(converter)

    def test_heavy_pdf_low_timeout(self):
        output = self.root / "out" / "timeout.docx"
        converter = PdfConverter(workers=1, timeout=1, retries=0)
        with converter:
            with self.assertRaises(ConversionTimeoutError):
                converter.convert(HEAVY_PDF, out=output, force=True)
        self.assertFalse(output.exists())
        self.assert_no_orphan_word(converter)


if __name__ == "__main__":
    unittest.main()

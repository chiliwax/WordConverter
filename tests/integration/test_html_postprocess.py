from __future__ import annotations

import unittest

from pdf2word import PdfConverter, ValidationError
from tests.integration.base import WordTestCase


class HtmlPostprocessingTests(WordTestCase):
    def test_source_folder_and_links(self):
        source = self.copy_sample("doc.pdf")
        out = self.root / "out" / "doc.html"
        converter = PdfConverter(workers=1, format="html", retries=0)
        with converter:
            result = converter.convert(source, out=out, force=True)
        html = result.output.read_text(encoding="utf-8")
        self.assertIn("charset=utf-8", html.lower())
        self.assertIn("source_folder/", html)
        self.assertTrue((result.output.parent / "source_folder" / "image001.png").is_file())
        self.assertFalse(any(result.output.parent.glob("*_files")))
        self.assertFalse(any(result.output.parent.glob("*_fichiers")))
        self.assert_no_orphan_word(converter)

    def test_force_replaces_existing_source_folder(self):
        source = self.copy_sample("doc.pdf")
        out = self.root / "out" / "doc.html"
        converter = PdfConverter(workers=1, format="html", retries=0)
        with converter:
            first = converter.convert(source, out=out, force=True)
            marker = first.output.parent / "source_folder" / "old.txt"
            marker.write_text("old", encoding="utf-8")
            second = converter.convert(source, out=out, force=True)
        self.assertTrue(second.output.is_file())
        self.assertFalse(marker.exists())
        self.assertTrue((second.output.parent / "source_folder" / "image001.png").is_file())
        self.assert_no_orphan_word(converter)

    def test_existing_source_folder_without_force(self):
        source = self.copy_sample("doc.pdf")
        out = self.root / "out" / "doc.html"
        (out.parent / "source_folder").mkdir(parents=True)
        converter = PdfConverter(workers=1, format="html", retries=0)
        with converter:
            with self.assertRaises(ValidationError):
                converter.convert(source, out=out, force=False)
        self.assertFalse(out.exists())
        self.assert_no_orphan_word(converter)


if __name__ == "__main__":
    unittest.main()

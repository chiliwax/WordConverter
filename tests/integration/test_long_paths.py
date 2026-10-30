from __future__ import annotations

import os
import shutil
import unittest

import pdf2word
from pdf2word import PdfConverter, ValidationError
from tests.helpers import SAMPLE_PDF, make_long_dir, staging_dirs
from tests.integration.base import WordTestCase


class LongPathTests(WordTestCase):
    def test_long_source_path(self):
        long_dir = make_long_dir(self.root, components=5, size=45)
        source = long_dir / "input.pdf"
        shutil.copy2(pdf2word._win_long(SAMPLE_PDF), pdf2word._win_long(source))
        self.assertGreater(len(str(source)), 260)

        output = self.root / "out" / "short.docx"
        converter = PdfConverter(workers=1, retries=0)
        with converter:
            result = converter.convert(source, out=output, force=True)
        self.assertTrue(result.output.is_file())
        self.assert_no_orphan_word(converter)

    def test_long_output_path(self):
        source = self.copy_sample()
        long_dir = make_long_dir(self.root, components=5, size=45)
        output = long_dir / "long.docx"
        self.assertGreater(len(str(output)), 260)

        converter = PdfConverter(workers=1, retries=0)
        with converter:
            result = converter.convert(source, out=output, force=True)
        self.assertTrue(pdf2word._exists(result.output))
        self.assert_no_orphan_word(converter)

    def test_long_html_image_path(self):
        source = self.copy_sample()
        long_dir = make_long_dir(self.root, components=5, size=45)
        output = long_dir / "long.html"
        self.assertGreater(len(str(output)), 260)

        converter = PdfConverter(workers=1, format="html", retries=0)
        with converter:
            result = converter.convert(source, out=output, force=True)
        self.assertTrue(pdf2word._exists(result.output))
        self.assertTrue(pdf2word._exists(
            result.output.parent / "source_folder" / "image001.png"
        ))
        html = PathRead(result.output)
        self.assertIn("source_folder/", html)
        self.assert_no_orphan_word(converter)

    def test_long_batch_mirror(self):
        long_root = make_long_dir(self.root, components=5, size=45)
        input_dir = long_root / "Input"
        source = input_dir / "blockchain" / "NP 2025-32" / "NP 2025-32.pdf"
        os.makedirs(pdf2word._win_long(source.parent), exist_ok=True)
        shutil.copy2(pdf2word._win_long(SAMPLE_PDF), pdf2word._win_long(source))
        out = long_root / "Output"

        converter = PdfConverter(workers=1, retries=0)
        with converter:
            batch = converter.convert_many(input_dir, out_dir=out, recursive=True, force=True)
        self.assertEqual(len(batch.ok), 1, batch.failed)
        expected = out / "blockchain" / "NP 2025-32" / "NP 2025-32.docx"
        self.assertTrue(pdf2word._exists(expected))
        self.assert_no_orphan_word(converter)

    def test_staging_cleanup(self):
        source = self.copy_sample()
        long_dir = make_long_dir(self.root, components=5, size=45)
        output = long_dir / "long.docx"
        before = set(staging_dirs())

        converter = PdfConverter(workers=1, retries=0)
        with converter:
            converter.convert(source, out=output, force=True)
        after = set(staging_dirs())
        self.assertEqual(after - before, set())
        self.assert_no_orphan_word(converter)

    def test_component_too_long(self):
        source = self.copy_sample()
        output = self.root / (("a" * 256) + ".docx")
        converter = PdfConverter(workers=1, retries=0)
        with converter:
            with self.assertRaises(ValidationError):
                converter.convert(source, out=output, force=True)
        self.assert_no_orphan_word(converter)

    def test_long_filename_docx(self):
        stem = "LJ 2025-106 " + ("N" * 170)
        source = self.root / "in" / (stem + ".pdf")
        source.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(pdf2word._win_long(SAMPLE_PDF), pdf2word._win_long(source))
        output = self.root / "out" / (stem + ".docx")
        self.assertGreater(len(source.name), 150)

        converter = PdfConverter(workers=1, retries=0)
        with converter:
            result = converter.convert(source, out=output, force=True)
        self.assertTrue(pdf2word._exists(result.output))
        self.assertEqual(result.output.name, stem + ".docx")
        self.assert_no_orphan_word(converter)

    def test_long_filename_html(self):
        stem = "LJ 2025-106 " + ("H" * 170)
        source = self.root / "in" / (stem + ".pdf")
        source.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(pdf2word._win_long(SAMPLE_PDF), pdf2word._win_long(source))
        output = self.root / "out" / (stem + ".html")
        self.assertGreater(len(source.name), 150)

        converter = PdfConverter(workers=1, format="html", retries=0)
        with converter:
            result = converter.convert(source, out=output, force=True)
        self.assertTrue(pdf2word._exists(result.output))
        self.assertEqual(result.output.name, stem + ".html")
        self.assertTrue(pdf2word._exists(
            result.output.parent / "source_folder" / "image001.png"
        ))
        self.assert_no_orphan_word(converter)


def PathRead(path):
    with open(pdf2word._fs_path(path), "r", encoding="utf-8") as handle:
        return handle.read()


if __name__ == "__main__":
    unittest.main()

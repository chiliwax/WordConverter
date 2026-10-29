from __future__ import annotations

import json
import unittest
from pathlib import Path

import pdf2word


class JsonSerializationTests(unittest.TestCase):
    def test_result_json(self):
        result = pdf2word.ConversionResult(
            source=Path(r"C:\Data\entrée.pdf"),
            output=Path(r"C:\Data\sortie.docx"),
            format="docx",
            pages=3,
            size_kb=12.5,
            seconds=1.2,
            attempts=2,
            companion_files=("source_folder",),
            pdf_warning_set=True,
        )
        payload = json.loads(pdf2word._result_json(result))
        self.assertTrue(payload["ok"])
        self.assertEqual(payload["source"], r"C:\Data\entrée.pdf")
        self.assertEqual(payload["output"], r"C:\Data\sortie.docx")
        self.assertEqual(payload["format"], "docx")
        self.assertEqual(payload["pages"], 3)
        self.assertEqual(payload["companion_files"], ["source_folder"])
        self.assertTrue(payload["pdf_warning_set"])

    def test_failure_json(self):
        failure = pdf2word.ConversionFailure(
            source=Path(r"C:\Data\missing.pdf"),
            error="Source file not found",
            error_type="validation",
            attempts=0,
        )
        payload = json.loads(pdf2word._failure_json(failure))
        self.assertFalse(payload["ok"])
        self.assertEqual(payload["source"], r"C:\Data\missing.pdf")
        self.assertEqual(payload["error"], "Source file not found")
        self.assertEqual(payload["error_type"], "validation")
        self.assertEqual(payload["attempts"], 0)


if __name__ == "__main__":
    unittest.main()

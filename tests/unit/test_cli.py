from __future__ import annotations

import contextlib
import io
import json
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import pdf2word
from tests.helpers import (
    FakeTTY,
    fake_converter_class,
    make_failure,
    make_result,
    run_cli,
    write_pdf,
)


class CliTests(unittest.TestCase):
    def setUp(self):
        self._temp = tempfile.TemporaryDirectory(prefix="pdf2word-cli-")
        self.root = Path(self._temp.name)
        self.addCleanup(self._temp.cleanup)
        self.source = write_pdf(self.root / "report.pdf")
        self.folder = self.root / "inbox"
        self.folder.mkdir()
        write_pdf(self.folder / "a.pdf")
        write_pdf(self.folder / "b.pdf")

    def test_single_file(self):
        result = make_result(self.source, self.root / "out" / "report.docx")
        fake_cls = fake_converter_class([result])
        with mock.patch.object(pdf2word, "PdfConverter", fake_cls):
            code, stdout, stderr = run_cli([
                "convert", str(self.source), "--out", str(self.root / "out"),
            ])
        self.assertEqual(code, 0, stderr)
        self.assertIn("1 converted, 0 failed", stdout)
        self.assertIn("ok", stdout)

    def test_folder_batch(self):
        results = [
            make_result(self.folder / "a.pdf", self.root / "out" / "a.docx"),
            make_result(self.folder / "b.pdf", self.root / "out" / "b.docx"),
        ]
        fake_cls = fake_converter_class(results)
        with mock.patch.object(pdf2word, "PdfConverter", fake_cls):
            code, stdout, stderr = run_cli([
                "convert", str(self.folder), "--out", str(self.root / "out"),
            ])
        self.assertEqual(code, 0, stderr)
        self.assertIn("[1/2]", stdout)
        self.assertIn("[2/2]", stdout)
        self.assertIn("2 converted, 0 failed", stdout)

    def test_json_includes_failures(self):
        result = make_result(self.folder / "a.pdf", self.root / "out" / "a.docx")
        failure = make_failure(self.root / "missing.pdf", "Source file not found",
                               error_type="validation")
        fake_cls = fake_converter_class([result], [failure])
        with mock.patch.object(pdf2word, "PdfConverter", fake_cls):
            code, stdout, stderr = run_cli([
                "convert", str(self.folder), "--out", str(self.root / "out"),
                "--json",
            ])
        self.assertEqual(code, 1)
        lines = [line for line in stdout.splitlines() if line.strip()]
        self.assertEqual(len(lines), 2)
        payloads = [json.loads(line) for line in lines]
        self.assertEqual({payload["ok"] for payload in payloads}, {True, False})
        self.assertNotIn("converted,", stdout)

    def test_quiet_only_reports_failures(self):
        result = make_result(self.folder / "a.pdf", self.root / "out" / "a.docx")
        failure = make_failure(self.root / "missing.pdf")
        fake_cls = fake_converter_class([result], [failure])
        with mock.patch.object(pdf2word, "PdfConverter", fake_cls):
            code, stdout, stderr = run_cli([
                "convert", str(self.folder), "--out", str(self.root / "out"),
                "--quiet",
            ])
        self.assertEqual(code, 1)
        self.assertNotIn("  ok  ", stdout)
        self.assertIn("ERR", stderr)

    def test_progress_non_tty_falls_back_to_lines(self):
        results = [
            make_result(self.folder / "a.pdf", self.root / "out" / "a.docx"),
            make_result(self.folder / "b.pdf", self.root / "out" / "b.docx"),
        ]
        fake_cls = fake_converter_class(results)
        with mock.patch.object(pdf2word, "PdfConverter", fake_cls):
            code, stdout, stderr = run_cli([
                "convert", str(self.folder), "--out", str(self.root / "out"),
                "--progress",
            ])
        self.assertEqual(code, 0)
        self.assertIn("[1/2]", stdout)
        self.assertIn("[2/2]", stdout)
        self.assertNotIn("\r", stderr)

    def test_progress_tty_uses_bar(self):
        results = [
            make_result(self.folder / "a.pdf", self.root / "out" / "a.docx"),
            make_result(self.folder / "b.pdf", self.root / "out" / "b.docx"),
        ]
        fake_cls = fake_converter_class(results)
        args = pdf2word._build_parser().parse_args([
            "convert", str(self.folder), "--out", str(self.root / "out"),
            "--progress",
        ])
        fake_stderr = FakeTTY()
        stdout = io.StringIO()
        with mock.patch.object(pdf2word, "PdfConverter", fake_cls):
            with contextlib.redirect_stdout(stdout):
                with mock.patch("sys.stderr", fake_stderr):
                    code = pdf2word.cmd_convert(args)
        text = fake_stderr.getvalue()
        self.assertEqual(code, 0)
        self.assertIn("\r", text)
        self.assertIn("%", text)
        self.assertIn("ETA", text)
        self.assertNotIn("  ok  ", text)

    def test_flat_conflict(self):
        code, stdout, stderr = run_cli([
            "convert", str(self.folder), "--layout", "subfolders", "--flat",
        ])
        self.assertEqual(code, 2)
        self.assertIn("--flat", stderr)

    def test_exit_code_on_failure(self):
        failure = make_failure(self.root / "missing.pdf")
        fake_cls = fake_converter_class([], [failure])
        with mock.patch.object(pdf2word, "PdfConverter", fake_cls):
            code, stdout, stderr = run_cli([
                "convert", str(self.folder), "--out", str(self.root / "out"),
            ])
        self.assertEqual(code, 1)

    def test_invalid_format_exits_2(self):
        with self.assertRaises(SystemExit) as captured:
            pdf2word.main(["convert", str(self.source), "--format", "invalid"])
        self.assertEqual(captured.exception.code, 2)

    def test_workers_option_is_passed(self):
        result = make_result(self.source, self.root / "out" / "report.docx")
        fake_cls = fake_converter_class([result])
        with mock.patch.object(pdf2word, "PdfConverter", fake_cls):
            code, _, _ = run_cli([
                "convert", str(self.source), "--out", str(self.root / "out"),
                "--workers", "3",
            ])
        self.assertEqual(code, 0)
        self.assertEqual(fake_cls.instances[-1].kwargs["workers"], 3)

    def test_single_validation_failure_as_json(self):
        fake_cls = fake_converter_class(
            raise_single=pdf2word.ValidationError("Source file not found")
        )
        with mock.patch.object(pdf2word, "PdfConverter", fake_cls):
            code, stdout, stderr = run_cli([
                "convert", str(self.source), "--json",
            ])
        self.assertEqual(code, 2)
        lines = [line for line in stdout.splitlines() if line.strip()]
        self.assertEqual(len(lines), 1)
        payload = json.loads(lines[0])
        self.assertFalse(payload["ok"])
        self.assertEqual(payload["error_type"], "validation")


if __name__ == "__main__":
    unittest.main()

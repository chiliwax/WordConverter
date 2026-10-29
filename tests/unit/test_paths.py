from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path

import pdf2word
from tests.helpers import make_job


@unittest.skipUnless(os.name == "nt", "Windows path helpers")
class PathHelperTests(unittest.TestCase):
    def test_absolute_relative(self):
        result = pdf2word._absolute("some/dir/file.pdf")
        self.assertTrue(result.is_absolute())
        self.assertEqual(result, Path.cwd() / "some/dir/file.pdf")

    def test_absolute_keeps_absolute_path(self):
        path = Path(r"C:\Data\file.pdf")
        self.assertEqual(pdf2word._absolute(path), path)

    def test_path_key_is_case_insensitive(self):
        self.assertEqual(
            pdf2word._path_key(r"C:\Data\File.PDF"),
            pdf2word._path_key(r"c:\data\file.pdf"),
        )

    def test_win_long_drive_path(self):
        self.assertEqual(
            pdf2word._win_long(r"C:\Data\file.pdf"),
            r"\\?\C:\Data\file.pdf",
        )

    def test_win_long_unc_path(self):
        self.assertEqual(
            pdf2word._win_long(r"\\server\share\file.pdf"),
            r"\\?\UNC\server\share\file.pdf",
        )

    def test_win_long_already_prefixed(self):
        self.assertEqual(
            pdf2word._win_long(r"\\?\C:\Data\file.pdf"),
            r"\\?\C:\Data\file.pdf",
        )

    def test_strip_long_prefix_drive(self):
        self.assertEqual(
            pdf2word._strip_long_prefix(r"\\?\C:\Data\file.pdf"),
            r"C:\Data\file.pdf",
        )

    def test_strip_long_prefix_unc(self):
        self.assertEqual(
            pdf2word._strip_long_prefix(r"\\?\UNC\server\share\file.pdf"),
            r"\\server\share\file.pdf",
        )

    def test_fs_path_short_unchanged(self):
        path = r"C:\Data\file.pdf"
        self.assertEqual(pdf2word._fs_path(path), path)

    def test_fs_path_long_gets_prefix(self):
        path = "C:\\" + ("a" * (pdf2word.LONG_PATH_THRESHOLD + 10)) + ".pdf"
        self.assertTrue(pdf2word._fs_path(path).startswith("\\\\?\\"))

    def test_relative_to_parent(self):
        self.assertEqual(
            pdf2word._relative_to(Path(r"C:\Data\a\b.pdf"), Path(r"C:\Data")),
            Path(r"a\b.pdf"),
        )

    def test_relative_to_same_path(self):
        self.assertEqual(
            pdf2word._relative_to(Path(r"C:\Data"), Path(r"C:\Data")),
            Path("."),
        )

    def test_relative_to_outside_base(self):
        self.assertIsNone(
            pdf2word._relative_to(Path(r"D:\Other\file.pdf"), Path(r"C:\Data"))
        )


@unittest.skipUnless(os.name == "nt", "Windows path helpers")
class IterPdfTests(unittest.TestCase):
    def setUp(self):
        self._temp = tempfile.TemporaryDirectory(prefix="pdf2word-iter-")
        self.root = Path(self._temp.name)
        (self.root / "sub").mkdir()

    def tearDown(self):
        self._temp.cleanup()

    def test_non_recursive(self):
        (self.root / "a.pdf").write_bytes(b"%PDF-")
        (self.root / "b.PDF").write_bytes(b"%PDF-")
        (self.root / "note.txt").write_text("x")
        (self.root / "sub" / "c.pdf").write_bytes(b"%PDF-")

        found = [path.name for path in pdf2word._iter_pdfs(self.root, False)]
        self.assertEqual(found, ["a.pdf", "b.PDF"])

    def test_recursive(self):
        (self.root / "a.pdf").write_bytes(b"%PDF-")
        (self.root / "sub" / "c.pdf").write_bytes(b"%PDF-")

        found = [path.name for path in pdf2word._iter_pdfs(self.root, True)]
        self.assertEqual(found, ["a.pdf", "c.pdf"])

    def test_empty_folder(self):
        self.assertEqual(pdf2word._iter_pdfs(self.root, True), [])

    def test_missing_folder(self):
        self.assertEqual(pdf2word._iter_pdfs(self.root / "missing", True), [])


class ComponentLengthTests(unittest.TestCase):
    def test_long_component_is_rejected(self):
        with self.assertRaises(pdf2word.ValidationError):
            pdf2word._check_component_length(
                Path("C:/") / ("a" * 256) / "file.pdf", "test path"
            )

    def test_normal_component_is_accepted(self):
        pdf2word._check_component_length(
            Path("C:/") / ("a" * 255) / "file.pdf", "test path"
        )


class StagingDecisionTests(unittest.TestCase):
    def test_short_paths_do_not_stage(self):
        job = make_job(r"C:\Data\in.pdf", r"C:\Data\out.docx")
        self.assertFalse(pdf2word._needs_staging(job))

    def test_long_source_stages(self):
        long = "C:\\" + ("a" * 260) + "\\in.pdf"
        job = make_job(long, r"C:\Data\out.docx")
        self.assertTrue(pdf2word._needs_staging(job))

    def test_long_output_stages(self):
        long = "C:\\" + ("a" * 260) + "\\out.docx"
        job = make_job(r"C:\Data\in.pdf", long)
        self.assertTrue(pdf2word._needs_staging(job))

    def test_long_html_image_path_stages(self):
        parent = "C:\\" + ("a" * 230)
        job = make_job(
            r"C:\Data\in.pdf",
            parent + "\\out.html",
            fmt="html",
        )
        self.assertTrue(pdf2word._needs_staging(job))


class SmallLogicTests(unittest.TestCase):
    def test_estimate_eta_without_data(self):
        self.assertIsNone(pdf2word._estimate_eta(0, 10, 5.0))

    def test_estimate_eta_midway(self):
        self.assertAlmostEqual(pdf2word._estimate_eta(5, 10, 10.0), 10.0)

    def test_estimate_eta_finished(self):
        self.assertEqual(pdf2word._estimate_eta(10, 10, 10.0), 0.0)

    def test_progress_percent(self):
        event = pdf2word.ProgressEvent(phase="completed", done=1, total=4)
        self.assertAlmostEqual(event.percent, 25.0)

    def test_progress_percent_empty(self):
        event = pdf2word.ProgressEvent(phase="finished", done=0, total=0)
        self.assertEqual(event.percent, 100.0)

    def test_rpc_failure_codes(self):
        for code in ("0x800706BE", "0x800706ba", "0x80010108", "0x800401FD"):
            self.assertTrue(pdf2word._is_rpc_failure(f"HRESULT: {code}"), code)

    def test_rpc_failure_other_error(self):
        self.assertFalse(pdf2word._is_rpc_failure("Source file not found"))


if __name__ == "__main__":
    unittest.main()

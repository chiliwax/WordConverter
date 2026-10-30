from __future__ import annotations

import unittest
from pathlib import Path

import pdf2word
from tests.helpers import (
    FakeSlot,
    long_temp_dir,
    make_job,
    make_long_dir,
    make_result,
)


class PublishStagedTests(unittest.TestCase):
    def setUp(self):
        self._temp = long_temp_dir("pdf2word-publish-")
        self.root = self._temp.__enter__()
        self.addCleanup(self._temp.__exit__, None, None, None)
        self.staging = self.root / "staging"
        self.staging.mkdir()

    def test_file_is_moved(self):
        staged = self.staging / "out.docx"
        staged.write_bytes(b"new")
        final = self.root / "final" / "out.docx"
        job = make_job(self.root / "in.pdf", final)
        result = make_result(job.source, staged)

        pdf2word._publish_staged(job, result)

        self.assertTrue(final.is_file())
        self.assertEqual(final.read_bytes(), b"new")
        self.assertFalse(staged.exists())

    def test_html_folder_is_moved_first(self):
        staged = self.staging / "out.html"
        staged.write_text("<html></html>", encoding="utf-8")
        images = self.staging / "source_folder"
        images.mkdir()
        (images / "image001.png").write_bytes(b"png")
        final = self.root / "final" / "out.html"
        job = make_job(self.root / "in.pdf", final, fmt="html")
        result = make_result(job.source, staged, fmt="html", companion_files=("source_folder",))

        pdf2word._publish_staged(job, result)

        self.assertTrue(final.is_file())
        self.assertTrue((final.parent / "source_folder" / "image001.png").is_file())

    def test_force_replaces_existing_file(self):
        staged = self.staging / "out.docx"
        staged.write_bytes(b"new")
        final = self.root / "final" / "out.docx"
        final.parent.mkdir()
        final.write_bytes(b"old")
        job = make_job(self.root / "in.pdf", final, force=True)
        result = make_result(job.source, staged)

        pdf2word._publish_staged(job, result)

        self.assertEqual(final.read_bytes(), b"new")

    def test_companion_conflict_without_force(self):
        staged = self.staging / "out.html"
        staged.write_text("<html></html>", encoding="utf-8")
        images = self.staging / "source_folder"
        images.mkdir()
        (images / "image001.png").write_bytes(b"png")
        final = self.root / "final" / "out.html"
        (final.parent / "source_folder").mkdir(parents=True)
        job = make_job(self.root / "in.pdf", final, fmt="html")
        result = make_result(job.source, staged, fmt="html", companion_files=("source_folder",))

        with self.assertRaises(pdf2word.ValidationError):
            pdf2word._publish_staged(job, result)

        self.assertFalse(final.exists())

    def test_rollback_removes_moved_companions(self):
        staged = self.staging / "out.html"
        images = self.staging / "source_folder"
        images.mkdir()
        (images / "image001.png").write_bytes(b"png")
        final = self.root / "final" / "out.html"
        job = make_job(self.root / "in.pdf", final, fmt="html")
        result = make_result(job.source, staged, fmt="html", companion_files=("source_folder",))

        with self.assertRaises(FileNotFoundError):
            pdf2word._publish_staged(job, result)

        self.assertFalse((final.parent / "source_folder").exists())

    def test_long_path_file(self):
        staged = self.staging / "out.docx"
        staged.write_bytes(b"long")
        long_parent = make_long_dir(self.root, components=5, size=45)
        final = long_parent / "out.docx"
        self.assertGreater(len(str(final)), 260)
        job = make_job(self.root / "in.pdf", final)
        result = make_result(job.source, staged)

        pdf2word._publish_staged(job, result)

        self.assertTrue(pdf2word._exists(final))
        with open(pdf2word._fs_path(final), "rb") as handle:
            self.assertEqual(handle.read(), b"long")

    def test_long_path_html_folder(self):
        staged = self.staging / "out.html"
        staged.write_text("<html></html>", encoding="utf-8")
        images = self.staging / "source_folder"
        images.mkdir()
        (images / "image001.png").write_bytes(b"png")
        long_parent = make_long_dir(self.root, components=5, size=45)
        final = long_parent / "out.html"
        job = make_job(self.root / "in.pdf", final, fmt="html")
        result = make_result(job.source, staged, fmt="html", companion_files=("source_folder",))

        pdf2word._publish_staged(job, result)

        self.assertTrue(pdf2word._exists(final))
        self.assertTrue(pdf2word._exists(final.parent / "source_folder" / "image001.png"))

    def test_run_job_uses_short_staged_names(self):
        source = self.root / "source.pdf"
        source.write_bytes(b"%PDF-1.4")
        long_parent = make_long_dir(self.root, components=5, size=45)
        final = long_parent / "final.docx"
        job = make_job(source, final)
        slot = FakeSlot()

        result = pdf2word._run_job(slot, job)

        self.assertEqual(len(slot.calls), 1)
        self.assertEqual(slot.calls[0].source.name, "input.pdf")
        self.assertEqual(slot.calls[0].output.name, "output.docx")
        self.assertEqual(result.output, final)
        self.assertTrue(pdf2word._exists(final))


if __name__ == "__main__":
    unittest.main()

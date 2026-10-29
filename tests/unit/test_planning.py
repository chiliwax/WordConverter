from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path

import pdf2word
from tests.helpers import FakeSlot, patched_slots, write_pdf


class PlanningTestCase(unittest.TestCase):
    def setUp(self):
        self._temp = tempfile.TemporaryDirectory(prefix="pdf2word-plan-")
        self.root = Path(self._temp.name)
        self.addCleanup(self._temp.cleanup)

    def converter(self, *, workers: int = 1, **kwargs) -> pdf2word.PdfConverter:
        converter = pdf2word.PdfConverter(workers=workers, **kwargs)
        self.addCleanup(converter.close)
        return converter

    def run_batch(self, sources, *, workers: int = 2, **kwargs):
        converter = self.converter(workers=workers)
        slots = [FakeSlot(i) for i in range(workers)]
        with patched_slots(converter, slots):
            result = converter.convert_many(sources, **kwargs)
        return result, slots

    def make_tree(self, relative: str, name: str = "file.pdf") -> Path:
        folder = self.root / relative
        folder.mkdir(parents=True, exist_ok=True)
        return write_pdf(folder / name)


class PlanSingleTests(PlanningTestCase):
    def test_default_output(self):
        source = write_pdf(self.root / "report.pdf")
        job = self.converter()._plan_single(
            source, None, format="docx", force=None, timeout=None, retries=None,
        )
        self.assertEqual(job.output, source.with_suffix(".docx"))
        self.assertEqual(job.source, source)

    def test_out_directory(self):
        source = write_pdf(self.root / "report.pdf")
        (self.root / "out").mkdir()
        job = self.converter()._plan_single(
            source, self.root / "out", format="html", force=None,
            timeout=None, retries=None,
        )
        self.assertEqual(job.output, self.root / "out" / "report.html")

    def test_out_file(self):
        source = write_pdf(self.root / "report.pdf")
        target = self.root / "custom.html"
        job = self.converter()._plan_single(
            source, target, format="html", force=None, timeout=None, retries=None,
        )
        self.assertEqual(job.output, target)

    def test_wrong_extension(self):
        source = write_pdf(self.root / "report.pdf")
        with self.assertRaises(pdf2word.ValidationError):
            self.converter()._plan_single(
                source, self.root / "report.txt", format="docx",
                force=None, timeout=None, retries=None,
            )

    def test_existing_output_without_force(self):
        source = write_pdf(self.root / "report.pdf")
        output = write_pdf(self.root / "report.docx")
        with self.assertRaises(pdf2word.ValidationError):
            self.converter()._plan_single(
                source, output, format="docx", force=False,
                timeout=None, retries=None,
            )

    def test_existing_output_with_force(self):
        source = write_pdf(self.root / "report.pdf")
        output = write_pdf(self.root / "report.docx")
        job = self.converter()._plan_single(
            source, output, format="docx", force=True,
            timeout=None, retries=None,
        )
        self.assertTrue(job.force)

    def test_refuses_source_overwrite(self):
        source = write_pdf(self.root / "report.pdf")
        with self.assertRaises(pdf2word.ValidationError):
            self.converter()._plan_single(
                source, source, format="pdf", force=True,
                timeout=None, retries=None,
            )

    def test_missing_source(self):
        with self.assertRaises(pdf2word.ValidationError):
            self.converter()._plan_single(
                self.root / "missing.pdf", None, format="docx",
                force=None, timeout=None, retries=None,
            )

    def test_non_pdf_source(self):
        source = self.root / "report.txt"
        source.write_text("x")
        with self.assertRaises(pdf2word.ValidationError):
            self.converter()._plan_single(
                source, None, format="docx", force=None,
                timeout=None, retries=None,
            )

    def test_component_too_long(self):
        source = write_pdf(self.root / "report.pdf")
        output = self.root / (("a" * 256) + ".docx")
        with self.assertRaises(pdf2word.ValidationError):
            self.converter()._plan_single(
                source, output, format="docx", force=None,
                timeout=None, retries=None,
            )

    def test_negative_timeout(self):
        source = write_pdf(self.root / "report.pdf")
        with self.assertRaises(pdf2word.ValidationError):
            self.converter()._plan_single(
                source, None, format="docx", force=None,
                timeout=-1, retries=None,
            )

    def test_negative_retries(self):
        source = write_pdf(self.root / "report.pdf")
        with self.assertRaises(pdf2word.ValidationError):
            self.converter()._plan_single(
                source, None, format="docx", force=None,
                timeout=None, retries=-1,
            )


class BatchLayoutTests(PlanningTestCase):
    def test_auto_mirror_for_folder(self):
        source = self.make_tree("Input/blockchain/NP 2025-32", "NP 2025-32.pdf")
        out = self.root / "Output"
        batch, _ = self.run_batch(
            self.root / "Input", out_dir=out, recursive=True, force=True,
        )
        self.assertEqual(len(batch.ok), 1, batch.failed)
        relative = batch.ok[0].output.relative_to(out)
        self.assertEqual(relative, Path("blockchain/NP 2025-32/NP 2025-32.docx"))
        self.assertFalse((out / "Input").exists())

    def test_auto_subfolders_for_file_list(self):
        first = write_pdf(self.root / "a" / "a.pdf")
        second = write_pdf(self.root / "b" / "b.pdf")
        out = self.root / "out"
        batch, _ = self.run_batch([first, second], out_dir=out, force=True)
        self.assertEqual(len(batch.ok), 2, batch.failed)
        self.assertEqual(
            [item.output.relative_to(out).as_posix() for item in batch.ok],
            ["a/a.docx", "b/b.docx"],
        )

    def test_mirror_excludes_input_root(self):
        self.make_tree("Input/blockchain/NP 2025-32", "NP 2025-32.pdf")
        out = self.root / "Output"
        batch, _ = self.run_batch(
            self.root / "Input", out_dir=out, recursive=True, force=True,
            layout="mirror",
        )
        self.assertEqual(
            batch.ok[0].output.relative_to(out).as_posix(),
            "blockchain/NP 2025-32/NP 2025-32.docx",
        )

    def test_base_dir_includes_input_root(self):
        self.make_tree("Input/blockchain/NP 2025-32", "NP 2025-32.pdf")
        out = self.root / "Output"
        batch, _ = self.run_batch(
            self.root / "Input", out_dir=out, recursive=True, force=True,
            base_dir=self.root,
        )
        self.assertEqual(
            batch.ok[0].output.relative_to(out).as_posix(),
            "Input/blockchain/NP 2025-32/NP 2025-32.docx",
        )

    def test_mirror_recursive(self):
        self.make_tree("Input/a/b", "nested.pdf")
        out = self.root / "out"
        batch, _ = self.run_batch(
            self.root / "Input", out_dir=out, recursive=True, force=True,
        )
        self.assertEqual(len(batch.ok), 1, batch.failed)
        self.assertEqual(
            batch.ok[0].output.relative_to(out).as_posix(),
            "a/b/nested.docx",
        )

    def test_subfolders_duplicate_stems(self):
        first = write_pdf(self.root / "one" / "same.pdf")
        second = write_pdf(self.root / "two" / "same.pdf")
        out = self.root / "out"
        batch, _ = self.run_batch(
            [first, second], out_dir=out, layout="subfolders", force=True,
        )
        self.assertEqual(
            [item.output.relative_to(out).as_posix() for item in batch.ok],
            ["same/same.docx", "same-2/same-2.docx"],
        )

    def test_flat_duplicate_names(self):
        first = write_pdf(self.root / "one" / "same.pdf")
        second = write_pdf(self.root / "two" / "same.pdf")
        out = self.root / "out"
        batch, _ = self.run_batch(
            [first, second], out_dir=out, layout="flat", force=True,
        )
        self.assertEqual(len(batch.ok), 1)
        self.assertEqual(len(batch.failed), 1)
        self.assertEqual(batch.failed[0].error_type, "validation")

    def test_flat_html_multiple_refused(self):
        self.make_tree("Input/dir", "a.pdf")
        write_pdf(self.root / "Input" / "dir" / "b.pdf")
        with self.assertRaises(pdf2word.ValidationError):
            self.run_batch(
                self.root / "Input", out_dir=self.root / "out",
                recursive=True, format="html", layout="flat", force=True,
            )

    def test_html_inline(self):
        source = self.make_tree("Input/blockchain/NP 2025-32", "NP 2025-32.pdf")
        out = self.root / "Output"
        batch, _ = self.run_batch(
            self.root / "Input", out_dir=out, recursive=True,
            format="html", force=True,
        )
        self.assertEqual(len(batch.ok), 1, batch.failed)
        output = batch.ok[0].output
        self.assertEqual(
            output.relative_to(out).as_posix(),
            "blockchain/NP 2025-32/NP 2025-32.html",
        )
        self.assertTrue((output.parent / "source_folder" / "image001.png").is_file())
        self.assertFalse((output.parent / source.stem).exists())

    def test_html_multiple_same_folder_subfolders(self):
        folder = self.root / "Input" / "dir"
        folder.mkdir(parents=True)
        write_pdf(folder / "a.pdf")
        write_pdf(folder / "b.pdf")
        out = self.root / "out"
        batch, _ = self.run_batch(
            self.root / "Input", out_dir=out, recursive=True,
            format="html", force=True,
        )
        self.assertEqual(len(batch.ok), 2, batch.failed)
        paths = sorted(item.output.relative_to(out).as_posix() for item in batch.ok)
        self.assertEqual(paths, ["dir/a/a.html", "dir/b/b.html"])
        for item in batch.ok:
            self.assertTrue((item.output.parent / "source_folder").is_dir())

    def test_html_source_folder_conflict(self):
        self.make_tree("Input/dir", "a.pdf")
        out = self.root / "out"
        (out / "dir" / "source_folder").mkdir(parents=True)
        batch, slots = self.run_batch(
            self.root / "Input", out_dir=out, recursive=True,
            format="html", force=False,
        )
        self.assertEqual(len(batch.ok), 0)
        self.assertEqual(len(batch.failed), 1)
        self.assertIn("source_folder", batch.failed[0].error)
        self.assertEqual(slots[0].calls, [])

    def test_duplicate_input(self):
        source = write_pdf(self.root / "a.pdf")
        batch, _ = self.run_batch(
            [source, source], out_dir=self.root / "out", force=True,
        )
        self.assertEqual(len(batch.ok), 1)
        self.assertEqual(len(batch.failed), 1)
        self.assertIn("Duplicate input", batch.failed[0].error)

    def test_empty_folder_failure(self):
        (self.root / "empty").mkdir()
        batch, _ = self.run_batch(
            self.root / "empty", out_dir=self.root / "out", force=True,
        )
        self.assertEqual(batch.total, 1)
        self.assertEqual(len(batch.failed), 1)
        self.assertIn("No PDF files", batch.failed[0].error)

    def test_bad_base_dir(self):
        source = write_pdf(self.root / "Input" / "a.pdf")
        other = self.root / "Other"
        other.mkdir()
        batch, _ = self.run_batch(
            [source], out_dir=self.root / "out", base_dir=other, force=True,
        )
        self.assertEqual(len(batch.failed), 1)
        self.assertIn("not inside base_dir", batch.failed[0].error)

    def test_invalid_layout(self):
        with self.assertRaises(pdf2word.ValidationError):
            self.run_batch(
                write_pdf(self.root / "a.pdf"), out_dir=self.root / "out",
                layout="unknown", force=True,
            )

    def test_mirror_file_list_without_base_dir(self):
        source = write_pdf(self.root / "a.pdf")
        out = self.root / "out"
        batch, _ = self.run_batch(
            [source], out_dir=out, layout="mirror", force=True,
        )
        self.assertEqual(batch.ok[0].output.relative_to(out).as_posix(), "a.docx")


if __name__ == "__main__":
    unittest.main()

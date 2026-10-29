from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import unittest
from pathlib import Path

import pdf2word
from pdf2word import PdfConverter, ValidationError
from tests.helpers import PDF2WORD_SCRIPT, REPO_ROOT
from tests.integration.base import WordTestCase


class HierarchyTests(WordTestCase):
    def test_exact_html_hierarchy(self):
        source = self.copy_sample("Input/blockchain/NP 2025-32/NP 2025-32.pdf")
        out = self.root / "Output"
        converter = PdfConverter(workers=1, format="html", retries=0)
        with converter:
            batch = converter.convert_many(
                self.root / "Input", out_dir=out, recursive=True, force=True,
            )
        self.assertEqual(len(batch.ok), 1, batch.failed)
        output = batch.ok[0].output
        self.assertEqual(
            output.relative_to(out).as_posix(),
            "blockchain/NP 2025-32/NP 2025-32.html",
        )
        self.assertTrue((output.parent / "source_folder" / "image001.png").is_file())
        self.assertFalse((out / "Input").exists())
        self.assertFalse((output.parent / "NP 2025-32").exists())
        self.assertFalse(any(output.parent.glob("*_fichiers")))
        self.assertFalse(any(output.parent.glob("*_files")))
        self.assert_no_orphan_word(converter)

    def test_exact_docx_hierarchy(self):
        self.copy_sample("Input/blockchain/NP 2025-32/NP 2025-32.pdf")
        out = self.root / "Output"
        converter = PdfConverter(workers=1, retries=0)
        with converter:
            batch = converter.convert_many(
                self.root / "Input", out_dir=out, recursive=True, force=True,
            )
        self.assertEqual(len(batch.ok), 1, batch.failed)
        self.assertEqual(
            batch.ok[0].output.relative_to(out).as_posix(),
            "blockchain/NP 2025-32/NP 2025-32.docx",
        )
        self.assertFalse((out / "Input").exists())
        self.assert_no_orphan_word(converter)

    def test_base_dir_includes_input_level(self):
        self.copy_sample("Input/blockchain/NP 2025-32/NP 2025-32.pdf")
        out = self.root / "Output"
        converter = PdfConverter(workers=1, retries=0)
        with converter:
            batch = converter.convert_many(
                self.root / "Input", out_dir=out, recursive=True,
                base_dir=self.root, force=True,
            )
        self.assertEqual(len(batch.ok), 1, batch.failed)
        self.assertEqual(
            batch.ok[0].output.relative_to(out).as_posix(),
            "Input/blockchain/NP 2025-32/NP 2025-32.docx",
        )

    def test_multiple_html_documents_in_one_folder(self):
        folder = self.root / "Input" / "dir"
        self.copy_sample("Input/dir/a.pdf")
        self.copy_sample("Input/dir/b.pdf")
        out = self.root / "Output"
        converter = PdfConverter(workers=2, format="html", retries=0)
        with converter:
            batch = converter.convert_many(
                self.root / "Input", out_dir=out, recursive=True, force=True,
            )
        self.assertEqual(len(batch.ok), 2, batch.failed)
        paths = sorted(item.output.relative_to(out).as_posix() for item in batch.ok)
        self.assertEqual(paths, ["dir/a/a.html", "dir/b/b.html"])
        for item in batch.ok:
            self.assertTrue((item.output.parent / "source_folder").is_dir())
        self.assert_no_orphan_word(converter)

    def test_unicode_and_spaces(self):
        source = self.copy_sample("Input/entrée des fichiers/café résumé.pdf")
        out = self.root / "Output"
        converter = PdfConverter(workers=1, retries=0)
        with converter:
            batch = converter.convert_many(
                self.root / "Input", out_dir=out, recursive=True, force=True,
            )
        self.assertEqual(len(batch.ok), 1, batch.failed)
        self.assertEqual(
            batch.ok[0].output.relative_to(out).as_posix(),
            "entrée des fichiers/café résumé.docx",
        )
        self.assert_no_orphan_word(converter)

    def test_source_hash_unchanged(self):
        source = self.copy_sample("Input/entrée.pdf")
        before = hashlib.sha256(source.read_bytes()).hexdigest()
        converter = PdfConverter(workers=1, retries=0)
        with converter:
            converter.convert(source, out=self.root / "out" / "entrée.docx", force=True)
        after = hashlib.sha256(source.read_bytes()).hexdigest()
        self.assertEqual(before, after)
        self.assert_no_orphan_word(converter)

    def test_progress_events(self):
        self.copy_sample("Input/a.pdf")
        self.copy_sample("Input/b.pdf")
        events = []
        converter = PdfConverter(workers=2, retries=0)
        with converter:
            batch = converter.convert_many(
                self.root / "Input", out_dir=self.root / "Output",
                recursive=True, force=True, on_event=events.append,
            )
        self.assertEqual(len(batch.ok), 2, batch.failed)
        phases = [event.phase for event in events]
        self.assertEqual(phases.count("started"), 2)
        self.assertEqual(phases.count("completed"), 2)
        self.assertEqual(phases.count("finished"), 1)
        finished = events[-1]
        self.assertEqual(finished.done, 2)
        self.assertEqual(finished.total, 2)
        self.assertEqual(finished.cancelled, 0)
        self.assertEqual(finished.percent, 100.0)
        self.assert_no_orphan_word(converter)

    def test_force_rerun_html(self):
        source = self.copy_sample("Input/dir/doc.pdf")
        out = self.root / "Output"
        converter = PdfConverter(workers=1, format="html", retries=0)
        with converter:
            first = converter.convert(source, out=out / "doc.html", force=True)
            marker = first.output.parent / "source_folder" / "marker.txt"
            marker.write_text("old", encoding="utf-8")
            second = converter.convert(source, out=out / "doc.html", force=True)
        self.assertTrue(second.output.is_file())
        self.assertTrue((second.output.parent / "source_folder" / "image001.png").is_file())
        self.assertFalse(marker.exists())
        self.assertFalse(any(second.output.parent.glob("*_fichiers")))
        self.assert_no_orphan_word(converter)

    def test_existing_output_refused(self):
        source = self.copy_sample()
        out = self.root / "out" / "sample.docx"
        converter = PdfConverter(workers=1, retries=0)
        with converter:
            converter.convert(source, out=out, force=True)
            with self.assertRaises(ValidationError):
                converter.convert(source, out=out, force=False)
            converter.convert(source, out=out, force=True)
        self.assert_no_orphan_word(converter)

    def test_html_source_folder_conflict(self):
        source = self.copy_sample()
        out = self.root / "out"
        (out / "source_folder").mkdir(parents=True)
        converter = PdfConverter(workers=1, format="html", retries=0)
        with converter:
            with self.assertRaises(ValidationError):
                converter.convert(source, out=out / "sample.html", force=False)
        self.assertFalse((out / "sample.html").exists())
        self.assert_no_orphan_word(converter)

    def test_cli_json_end_to_end(self):
        self.copy_sample("Input/blockchain/NP 2025-32/NP 2025-32.pdf")
        self.copy_sample("Input/other/second.pdf")
        out = self.root / "Output"
        command = [
            sys.executable, str(PDF2WORD_SCRIPT),
            "convert", str(self.root / "Input"),
            "-r", "--out", str(out), "--json", "--force",
        ]
        completed = subprocess.run(
            command, cwd=str(REPO_ROOT), capture_output=True, text=True, timeout=300,
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)
        payloads = [
            json.loads(line) for line in completed.stdout.splitlines() if line.strip()
        ]
        self.assertEqual(len(payloads), 2, completed.stdout)
        self.assertTrue(all(payload["ok"] for payload in payloads))
        outputs = sorted(Path(payload["output"]).relative_to(out).as_posix()
                         for payload in payloads)
        self.assertEqual(outputs, [
            "blockchain/NP 2025-32/NP 2025-32.docx",
            "other/second.docx",
        ])

    def test_cli_json_validation_failure(self):
        missing = [self.root / "missing-a.pdf", self.root / "missing-b.pdf"]
        out = self.root / "Output"
        command = [
            sys.executable, str(PDF2WORD_SCRIPT),
            "convert", *map(str, missing), "--out", str(out), "--json",
        ]
        completed = subprocess.run(
            command, cwd=str(REPO_ROOT), capture_output=True, text=True, timeout=120,
        )
        self.assertEqual(completed.returncode, 1, completed.stderr)
        lines = [line for line in completed.stdout.splitlines() if line.strip()]
        self.assertEqual(len(lines), 2, completed.stdout)
        payloads = [json.loads(line) for line in lines]
        self.assertTrue(all(not payload["ok"] for payload in payloads))
        self.assertTrue(all(payload["error_type"] == "validation" for payload in payloads))


if __name__ == "__main__":
    unittest.main()

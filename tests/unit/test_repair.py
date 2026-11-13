from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from pdf2word import repair_html_outputs

E_ACUTE = "\u00e9"      # é
E_GRAVE = "\u00e8"      # è
A_GRAVE = "\u00e0"      # à
E_ACUTE_UP = "\u00c9"   # É
A_GRAVE_UP = "\u00c0"   # À
E_GRAVE_UP = "\u00c8"   # È


class RepairHtmlOutputTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def make_doc(self, stem: str, reference: str,
                 folder: str | None = None,
                 directory: Path | None = None) -> Path:
        """Write <stem>.html referencing <reference>/image001.png.

        ``folder`` is created when given; ``reference`` may use any encoding.
        """
        base = directory or self.root
        if folder is not None:
            images = base / folder
            images.mkdir(parents=True)
            (images / "image001.png").write_bytes(b"\x89PNG fake")
            (images / "filelist.xml").write_text("<xml/>", encoding="utf-8")

        html = base / f"{stem}.html"
        html.write_text(
            "<html><head>"
            f'<link rel=File-List href="{reference}/filelist.xml">'
            "</head><body>"
            f'<img src="{reference}/image001.png">'
            "</body></html>",
            encoding="utf-8",
        )
        return html

    def test_repairs_page_when_folder_was_already_renamed(self) -> None:
        # The reported case: Word wrote %20 for the space but a raw é, the
        # folder was renamed to source_folder, and only the HTML is stale.
        stem = "Mesure applicative_FR_sign" + E_ACUTE
        reference = "Mesure%20applicative_FR_sign" + E_ACUTE + "_fichiers"
        html = self.make_doc(stem, reference, folder="source_folder")

        summary = repair_html_outputs(self.root)

        self.assertEqual(len(summary.repaired), 1)
        result = summary.repaired[0]
        self.assertEqual(result.links_replaced, 2)
        self.assertIsNone(result.renamed_from)
        text = html.read_text(encoding="utf-8")
        self.assertIn("source_folder/image001.png", text)
        self.assertIn("source_folder/filelist.xml", text)
        self.assertNotIn(reference, text)
        # The folder is untouched and the page still works.
        self.assertTrue((self.root / "source_folder" / "image001.png").is_file())

        # Idempotent: a second pass changes nothing.
        again = repair_html_outputs(self.root)
        self.assertEqual(len(again.repaired), 0)
        self.assertEqual(len(again.unchanged), 1)

    def test_renames_localized_companion_folder(self) -> None:
        html = self.make_doc("NP 2025-32", "NP%202025-32_fichiers",
                             folder="NP 2025-32_fichiers")

        summary = repair_html_outputs(self.root)

        self.assertEqual(len(summary.repaired), 1)
        self.assertEqual(summary.repaired[0].renamed_from, "NP 2025-32_fichiers")
        self.assertFalse((self.root / "NP 2025-32_fichiers").exists())
        self.assertTrue((self.root / "source_folder" / "image001.png").is_file())
        text = html.read_text(encoding="utf-8")
        self.assertIn("source_folder/image001.png", text)

    def test_decodes_uppercase_accents_in_cp1252(self) -> None:
        # Some Word versions escape with the ANSI code page: %C9 is É, not %C3%89.
        stem = "Test " + E_ACUTE_UP + "cole " + A_GRAVE_UP + " " + E_GRAVE_UP
        folder = stem + "_fichiers"
        reference = "Test%20%C9cole%20%C0%20%C8_fichiers"
        html = self.make_doc(stem, reference, folder=folder)

        repair_html_outputs(self.root)

        self.assertFalse((self.root / folder).exists())
        text = html.read_text(encoding="utf-8")
        self.assertIn("source_folder/image001.png", text)

    def test_decodes_utf8_escaped_uppercase_accents(self) -> None:
        stem = "Test " + E_ACUTE_UP + A_GRAVE_UP + E_GRAVE_UP
        folder = stem + "_fichiers"
        reference = "Test%20%C3%89%C3%80%C3%88_fichiers"
        html = self.make_doc(stem, reference, folder=folder)

        repair_html_outputs(self.root)

        text = html.read_text(encoding="utf-8")
        self.assertIn("source_folder/image001.png", text)
        self.assertTrue((self.root / "source_folder" / "image001.png").is_file())

    def test_decodes_html_entities(self) -> None:
        html = self.make_doc("A & B", "A%20&amp;%20B_files", folder="A & B_files")

        repair_html_outputs(self.root)

        self.assertFalse((self.root / "A & B_files").exists())
        text = html.read_text(encoding="utf-8")
        self.assertIn("source_folder/image001.png", text)

    def test_rewrites_absolute_and_css_links(self) -> None:
        stem = "Doc " + E_ACUTE + E_GRAVE + A_GRAVE
        reference = "Doc%20" + E_ACUTE + E_GRAVE + A_GRAVE + "_fichiers"
        (self.root / "source_folder").mkdir()
        (self.root / "source_folder" / "image001.png").write_bytes(b"png")
        html = self.root / f"{stem}.html"
        html.write_text(
            f'<img src="file:///C:/tmp/{reference}/image001.png">'
            f"<style>p {{ background-image: url('{reference}/bg.png'); }}</style>",
            encoding="utf-8",
        )

        summary = repair_html_outputs(self.root)

        self.assertEqual(len(summary.repaired), 1)
        self.assertEqual(summary.repaired[0].links_replaced, 2)
        text = html.read_text(encoding="utf-8")
        self.assertIn("file:///C:/tmp/source_folder/image001.png", text)
        self.assertIn("url('source_folder/bg.png')", text)

    def test_real_world_names_with_parentheses_and_apostrophes(self) -> None:
        # Names taken from names.txt: spaces encoded, accents raw, and literal
        # parentheses/apostrophes inside the HTML attribute value.
        names = [
            "ARCHER DESIGN PATTERNS AWS S3 (EN)",
            "ARCHER DESIGN PATTERNS AWS S3 (FR)",
            "CLOUD ASSET INVENTORY (EN)",
            "CAM0036-V2-Origine de l'Etablissement",
            "CAM0039-V4-Tranche d'Effectif Salari" + E_ACUTE,
            "CAM0054-V1-Mode d'utilisation du taux",
            "CAM0056-V1-Source de l'index",
            "CAM0067-V2-Tranches d'" + "\u00e2" + "ges",
            "CAM0070-V1-Relev" + E_ACUTE + " d'identit" + E_ACUTE + " bancaire",
            "CAM0128-V1-Mode d'amortissement Cr" + E_ACUTE + "dit",
            "CAM0129-V1-Type d'" + E_ACUTE + "ch" + E_ACUTE + "ance Cr" + E_ACUTE + "dit",
            "CAM0180-V1-Zones d'emploi",
        ]

        for index, stem in enumerate(names):
            directory = self.root / f"case{index:02d}"
            directory.mkdir()
            reference = stem.replace(" ", "%20") + "_fichiers"
            self.make_doc(stem, reference, folder="source_folder",
                          directory=directory)

        summary = repair_html_outputs(self.root)

        self.assertEqual(len(summary.repaired), len(names))
        for index, stem in enumerate(names):
            with self.subTest(stem=stem):
                html = self.root / f"case{index:02d}" / f"{stem}.html"
                text = html.read_text(encoding="utf-8")
                self.assertIn("source_folder/image001.png", text)
                self.assertIn("source_folder/filelist.xml", text)
                self.assertNotIn("_fichiers/", text)

    def test_renames_real_world_folders_still_named_after_word(self) -> None:
        cases = [
            ("ARCHER DESIGN PATTERNS AWS S3 (EN)",
             "ARCHER%20DESIGN%20PATTERNS%20AWS%20S3%20(EN)_fichiers"),
            ("CAM0036-V2-Origine de l'Etablissement",
             "CAM0036-V2-Origine%20de%20l'Etablissement_fichiers"),
            ("CLOUD ASSET INVENTORY (EN)",
             "CLOUD%20ASSET%20INVENTORY%20(EN)_fichiers"),
        ]

        for index, (stem, reference) in enumerate(cases):
            directory = self.root / f"case{index:02d}"
            directory.mkdir()
            self.make_doc(stem, reference, folder=stem + "_fichiers",
                          directory=directory)

        summary = repair_html_outputs(self.root)

        self.assertEqual(len(summary.repaired), len(cases))
        for index, (stem, _reference) in enumerate(cases):
            with self.subTest(stem=stem):
                directory = self.root / f"case{index:02d}"
                self.assertFalse((directory / (stem + "_fichiers")).exists())
                self.assertTrue((directory / "source_folder" / "image001.png").is_file())
                text = (directory / f"{stem}.html").read_text(encoding="utf-8")
                self.assertIn("source_folder/image001.png", text)

    def test_dry_run_writes_nothing(self) -> None:
        html = self.make_doc("doc", "doc_fichiers", folder="doc_fichiers")
        before = html.read_bytes()

        summary = repair_html_outputs(self.root, dry_run=True)

        self.assertTrue(summary.dry_run)
        self.assertEqual(len(summary.repaired), 1)
        self.assertTrue((self.root / "doc_fichiers").is_dir())
        self.assertFalse((self.root / "source_folder").exists())
        self.assertEqual(html.read_bytes(), before)

    def test_skips_document_without_image_folder(self) -> None:
        html = self.root / "plain.html"
        html.write_text("<html><body>hello</body></html>", encoding="utf-8")

        summary = repair_html_outputs(self.root)

        self.assertEqual(len(summary.skipped), 1)
        self.assertEqual(len(summary.repaired), 0)

    def test_merges_leftover_folder_without_deleting(self) -> None:
        # Both folders exist: same-name files stay in the old one and are kept.
        stem = "doc"
        reference = "doc_fichiers"
        html = self.make_doc(stem, reference, folder="source_folder")
        old = self.root / reference
        old.mkdir()
        (old / "image001.png").write_bytes(b"other")
        (old / "extra.png").write_bytes(b"extra")

        summary = repair_html_outputs(self.root)

        self.assertEqual(len(summary.repaired), 1)
        # extra.png was moved, the conflicting image001.png was left in place.
        self.assertTrue((self.root / "source_folder" / "extra.png").is_file())
        self.assertTrue((old / "image001.png").is_file())
        text = html.read_text(encoding="utf-8")
        self.assertIn("source_folder/image001.png", text)

        # The page no longer references the leftover folder: a second pass is a
        # no-op even though the old folder could not be fully emptied.
        again = repair_html_outputs(self.root)
        self.assertEqual(len(again.repaired), 0)
        self.assertEqual(len(again.unchanged), 1)


if __name__ == "__main__":
    unittest.main()

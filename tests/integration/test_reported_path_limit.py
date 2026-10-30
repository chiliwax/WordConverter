from __future__ import annotations

import shutil
import unittest
from pathlib import Path

from pdf2word import PdfConverter
from tests.helpers import SAMPLE_PDF
from tests.integration.base import WordTestCase

# Path reported by the user (237 characters on the original machine).
REPORTED_FILENAME = (
    "LJ 2025-106 Norme d'échange entre les systemes d'information du groupe - API.pdf"
)
REPORTED_FOLDER = Path(REPORTED_FILENAME).stem
# The original path is 237 characters. On English Word the companion folder is
# "_files"; 245 reproduces the same "supporting files path is too long" error.
# French Word uses "_fichiers" (3 characters longer), which explains why 237 can
# fail on the user's machine.
TARGET_LENGTH = 245


class ReportedPathLimitTests(WordTestCase):
    """Reproducer for the Word SaveAs error on a 218-249 character path."""

    def test_reported_long_path_to_html(self):
        # The reported path is 237 characters. The temporary root can shift the
        # final length slightly, so the test asserts the failing 218-249 range
        # rather than one exact number.
        relative = Path("LAC - MD Transformation") / REPORTED_FOLDER / REPORTED_FILENAME
        candidate = self.root / relative

        # The temporary root can be longer than the original user profile.
        # Drop the top folder if needed, then pad to exactly the reported length.
        if len(str(candidate)) > 249:
            relative = Path(REPORTED_FOLDER) / REPORTED_FILENAME
            candidate = self.root / relative

        deficit = TARGET_LENGTH - len(str(candidate))
        if deficit > 1:
            candidate = self.root / ("p" * (deficit - 1)) / relative

        if len(str(candidate)) >= 250:
            self.skipTest(f"temporary path is too long for this reproducer: {candidate}")

        self.assertGreaterEqual(len(str(candidate)), 218)
        self.assertLess(len(str(candidate)), 250)

        candidate.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(SAMPLE_PDF, candidate)
        output = candidate.with_suffix(".html")

        converter = PdfConverter(workers=1, format="html", retries=0, timeout=180)
        with converter:
            result = converter.convert(candidate, out=output, force=True)

        self.assertTrue(result.output.is_file())
        self.assertTrue((result.output.parent / "source_folder").is_dir())

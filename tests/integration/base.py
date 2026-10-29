from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path

from pdf2word import PdfConverter
from tests.helpers import SAMPLE_PDF, remove_long_tree, slow_tests_enabled, word_tests_enabled

slow = unittest.skipUnless(
    slow_tests_enabled(),
    "set PDF2WORD_SLOW_TESTS=1 to run slow tests",
)


class WordTestCase(unittest.TestCase):
    """Base class: skips the whole class when Word tests are not enabled."""

    @classmethod
    def setUpClass(cls):
        if not word_tests_enabled():
            raise unittest.SkipTest(
                "set PDF2WORD_RUN_WORD_TESTS=1 to run Microsoft Word tests"
            )
        if not SAMPLE_PDF.is_file():
            raise unittest.SkipTest(f"sample PDF not found: {SAMPLE_PDF}")
        try:
            with PdfConverter(workers=1) as converter:
                converter.ping()
        except Exception as exc:  # noqa: BLE001 - skip, do not fail the suite
            raise unittest.SkipTest(f"Microsoft Word is not available: {exc}")

    def setUp(self):
        self.root = Path(tempfile.mkdtemp(prefix="pdf2word-it-"))
        self.addCleanup(lambda: remove_long_tree(self.root))

    def copy_sample(self, relative: str = "sample.pdf") -> Path:
        target = self.root / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(SAMPLE_PDF, target)
        return target

    def make_tree(self, relative: str, name: str = "file.pdf") -> Path:
        return self.copy_sample(relative.rstrip("\\/") + "/" + name)

    def assert_no_orphan_word(self, converter: PdfConverter) -> None:
        self.assertEqual(converter.word_pids(), [])

"""Reject compressed files that would expand beyond the DOCX safety limit."""

import tempfile
import unittest
import zipfile
from pathlib import Path

from app import is_docx_file


PROJECT_DIR = Path(__file__).resolve().parents[1]


class UploadLimitTests(unittest.TestCase):
    def test_rejects_docx_zip_bomb(self):
        temporary_root = PROJECT_DIR / "data"
        temporary_root.mkdir(exist_ok=True)
        with tempfile.TemporaryDirectory(dir=temporary_root) as temporary_dir:
            path = Path(temporary_dir) / "oversized.docx"
            with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                archive.writestr("[Content_Types].xml", "<Types/>")
                with archive.open("word/document.xml", "w") as document:
                    chunk = b"0" * (1024 * 1024)
                    for _ in range(51):
                        document.write(chunk)

            self.assertLess(path.stat().st_size, 1024 * 1024)
            self.assertFalse(is_docx_file(path))


if __name__ == "__main__":
    unittest.main()

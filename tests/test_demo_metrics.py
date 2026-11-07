"""Keep the published demonstration numbers reproducible."""

import tempfile
import unittest
from pathlib import Path

from utils.format_checker import FormatChecker


FIXTURE = Path(__file__).resolve().parents[1] / "examples" / "演示用_多类型格式错误.docx"


class DemoMetricsTests(unittest.TestCase):
    def test_general_template_check_and_fix(self):
        self.assertTrue(FIXTURE.is_file(), "The packaged demo document is missing")

        checker = FormatChecker(template_id="general")
        before, _ = checker.check_file(FIXTURE)
        self.assertEqual(len(before), 64)
        self.assertEqual(sum(issue.fixable for issue in before), 61)

        temporary_root = FIXTURE.parents[1] / "data"
        temporary_root.mkdir(exist_ok=True)
        with tempfile.TemporaryDirectory(dir=temporary_root) as temporary_dir:
            fixed = Path(temporary_dir) / "fixed.docx"
            changes = checker.fix_file(FIXTURE, fixed)
            after, _ = FormatChecker(template_id="general").check_file(fixed)

        self.assertEqual(len(changes), 91)
        self.assertEqual(len(after), 3)


if __name__ == "__main__":
    unittest.main()

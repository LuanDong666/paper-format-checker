# -*- coding: utf-8 -*-
"""端到端接口测试：不依赖外网或真实 API Key。"""

import io
import shutil
import unittest
from pathlib import Path

from app import create_app


PROJECT_DIR = Path(__file__).resolve().parents[1]
FIXTURE = PROJECT_DIR / "test_document.docx"
TEST_DATA = PROJECT_DIR / "data" / ".test-data"


class PaperCheckerApiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        shutil.rmtree(TEST_DATA, ignore_errors=True)
        cls.app = create_app({"TESTING": True, "DATA_FOLDER": TEST_DATA})
        cls.client = cls.app.test_client()

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(TEST_DATA, ignore_errors=True)

    def upload_document(self, endpoint, **fields):
        fields["file"] = (io.BytesIO(FIXTURE.read_bytes()), "测试论文.docx")
        return self.client.post(endpoint, data=fields, content_type="multipart/form-data")

    def test_health_and_templates(self):
        self.assertEqual(self.client.get("/api/health").status_code, 200)
        data = self.client.get("/api/templates").get_json()
        self.assertEqual(len(data["templates"]), 4)

    def test_rejects_missing_or_invalid_file(self):
        self.assertEqual(self.client.post("/api/check").status_code, 400)
        response = self.client.post(
            "/api/check",
            data={"file": (io.BytesIO(b"not a Word document"), "fake.docx")},
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 400)

    def test_check_creates_isolated_report(self):
        first = self.upload_document("/api/check", template="general")
        second = self.upload_document("/api/check", template="general")
        self.assertEqual(first.status_code, 200)
        self.assertEqual(second.status_code, 200)
        first_data, second_data = first.get_json(), second.get_json()
        self.assertNotEqual(first_data["job_id"], second_data["job_id"])
        self.assertGreater(first_data["summary"]["total_issues"], 0)
        report = self.client.get(first_data["report_url"])
        self.assertEqual(report.status_code, 200)
        self.assertIn("格式检查报告", report.get_data(as_text=True))
        report.close()

    def test_fix_produces_download(self):
        response = self.upload_document("/api/fix", template="general", issues_before="0")
        self.assertEqual(response.status_code, 200)
        data = response.get_json()
        self.assertGreater(data["total_fixes"], 0)
        download = self.client.get(data["download_url"])
        self.assertEqual(download.status_code, 200)
        self.assertEqual(download.mimetype, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        download.close()

    def test_parse_requires_text(self):
        self.assertEqual(self.client.post("/api/parse-requirements", json={}).status_code, 400)


if __name__ == "__main__":
    unittest.main(verbosity=2)

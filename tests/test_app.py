# -*- coding: utf-8 -*-
"""端到端接口测试：不依赖外网或真实 API Key。"""

import io
import os
import shutil
import time
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
        cls.client.get("/register")
        with cls.client.session_transaction() as auth_session:
            csrf = auth_session["csrf_token"]
        registration = cls.client.post("/register", data={"csrf_token": csrf,
            "email": "tester@example.edu", "password": "test-password-123"})
        assert registration.status_code == 302

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(TEST_DATA, ignore_errors=True)

    def upload_document(self, endpoint, **fields):
        fields["file"] = (io.BytesIO(FIXTURE.read_bytes()), "测试论文.docx")
        return self.client.post(endpoint, data=fields, content_type="multipart/form-data",
                                headers=self.csrf_headers())

    def csrf_headers(self):
        self.client.get("/")
        with self.client.session_transaction() as auth_session:
            return {"X-CSRF-Token": auth_session.get("csrf_token") or ""}

    def test_health_and_templates(self):
        self.assertEqual(self.client.get("/api/health").status_code, 200)
        data = self.client.get("/api/templates").get_json()
        self.assertEqual(len(data["templates"]), 4)

    def test_multiple_pages_and_official_resource_links(self):
        pages = (("/", "格式检查"), ("/guide", "格式指南"),
                 ("/similarity", "查重资源"), ("/profile", "我的资料"),
                 ("/history", "我的历史"))
        for path, title in pages:
            with self.subTest(path=path):
                response = self.client.get(path)
                self.assertEqual(response.status_code, 200)
                self.assertIn(title, response.get_data(as_text=True))
                self.assertIn("site-shell.css", response.get_data(as_text=True))
        links = self.client.get("/similarity").get_data(as_text=True)
        self.assertIn("https://cx.cnki.net/", links)
        self.assertIn("https://check.wanfangdata.com.cn/", links)
        self.assertIn("https://dsa.dayainfo.com/", links)
        self.assertIn("https://www.eshukan.com/", links)
        self.assertIn("https://opaj.napstic.cn/periodicals", links)
        self.assertIn("https://www.letpub.com.cn/index.php?page=journalapp", links)
        self.assertIn('rel="noopener noreferrer"', links)

    def test_rejects_missing_or_invalid_file(self):
        self.assertEqual(self.client.post("/api/check", headers=self.csrf_headers()).status_code, 400)
        response = self.client.post(
            "/api/check",
            data={"file": (io.BytesIO(b"not a Word document"), "fake.docx")},
            content_type="multipart/form-data", headers=self.csrf_headers(),
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
        self.assertGreaterEqual(data["issues_before"], data["issues_after"])
        comparison = self.client.get(data["comparison_url"])
        self.assertEqual(comparison.status_code, 200)
        self.assertIn("修复前后对照", comparison.get_data(as_text=True))
        download = self.client.get(data["download_url"])
        self.assertEqual(download.status_code, 200)
        self.assertEqual(download.mimetype, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        download.close()

    def test_profile_and_history(self):
        csrf = self.csrf_headers()["X-CSRF-Token"]
        update = self.client.post("/profile", data={"csrf_token": csrf,
            "display_name": "蓝鲸同学", "avatar": "book"})
        self.assertEqual(update.status_code, 302)
        profile = self.client.get("/profile").get_data(as_text=True)
        self.assertIn("蓝鲸同学", profile)
        self.assertIn("value=\"book\" checked", profile)
        self.assertEqual(self.client.post("/profile", data={"csrf_token": csrf,
            "display_name": "测试", "avatar": "unknown"}).status_code, 200)
        check = self.upload_document("/api/check", template="general").get_json()
        fix = self.upload_document("/api/fix", template="general").get_json()
        history = self.client.get("/history").get_data(as_text=True)
        self.assertIn(check["job_id"], history)
        self.assertIn(fix["job_id"], history)
        self.assertIn("查看对照", history)
        deletion = self.client.post("/history/delete/" + check["job_id"],
                                    data={"csrf_token": self.csrf_headers()["X-CSRF-Token"]})
        self.assertEqual(deletion.status_code, 302)
        self.assertNotIn(check["job_id"], self.client.get("/history").get_data(as_text=True))
        self.assertEqual(self.client.get(check["report_url"]).status_code, 404)
        expired = self.upload_document("/api/check", template="general").get_json()
        expired_dir = TEST_DATA / expired["job_id"]
        old_time = time.time() - self.app.config["JOB_TTL_SECONDS"] - 60
        os.utime(expired_dir, (old_time, old_time))
        self.assertEqual(self.client.get(expired["report_url"]).status_code, 404)
        history = self.client.get("/history").get_data(as_text=True)
        self.assertIn(expired["job_id"], history)
        self.assertIn("文件已到期清理", history)

    def test_parse_requires_text(self):
        self.assertEqual(self.client.post("/api/parse-requirements", json={}, headers=self.csrf_headers()).status_code, 400)

    def test_login_csrf_and_job_ownership(self):
        first = self.upload_document("/api/fix", template="general").get_json()
        self.assertEqual(self.client.post("/api/check").status_code, 403)
        self.client.post("/logout", data={"csrf_token": self.csrf_headers()["X-CSRF-Token"]})
        self.assertEqual(self.client.get(first["comparison_url"]).status_code, 302)
        self.client.get("/register")
        with self.client.session_transaction() as auth_session:
            csrf = auth_session["csrf_token"]
        self.client.post("/register", data={"csrf_token": csrf,
            "email": "other@example.edu", "password": "another-password-123"})
        self.assertEqual(self.client.get(first["comparison_url"]).status_code, 404)
        self.assertEqual(self.client.get(first["download_url"]).status_code, 404)
        self.assertNotIn(first["job_id"], self.client.get("/history").get_data(as_text=True))
        self.client.post("/logout", data={"csrf_token": self.csrf_headers()["X-CSRF-Token"]})
        self.client.get("/login")
        with self.client.session_transaction() as auth_session:
            csrf = auth_session["csrf_token"]
        response = self.client.post("/login", data={"csrf_token": csrf,
            "email": "tester@example.edu", "password": "test-password-123"})
        self.assertEqual(response.status_code, 302)


if __name__ == "__main__":
    unittest.main(verbosity=2)

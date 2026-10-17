# -*- coding: utf-8 -*-
"""论文/作业格式检查与修复系统的 Flask 应用。"""

import json
import os
import shutil
import time
import uuid
import zipfile
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.exceptions import RequestEntityTooLarge

from utils.format_checker import FORMAT_TEMPLATES, FormatChecker, get_font_options, get_template_list
from utils.report_generator import generate_html_report, generate_report
from utils.requirement_parser import merge_with_default, parse_requirements

ALLOWED_EXTENSIONS = {"docx"}
MAX_UPLOAD_BYTES = 50 * 1024 * 1024
JOB_TTL_SECONDS = 24 * 60 * 60


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def is_docx_file(file_path: Path) -> bool:
    """Reject renamed non-Office files before python-docx tries to parse them."""
    try:
        with zipfile.ZipFile(file_path) as archive:
            names = set(archive.namelist())
            return "[Content_Types].xml" in names and "word/document.xml" in names
    except (OSError, zipfile.BadZipFile):
        return False


def safe_original_name(filename: str) -> str:
    """Keep a display name without permitting path components."""
    name = Path(filename or "").name.strip()
    return name if name and allowed_file(name) else "document.docx"


def create_app(test_config=None):
    app = Flask(__name__)
    base_dir = Path(__file__).resolve().parent
    app.config.from_mapping(
        DATA_FOLDER=base_dir / "data",
        MAX_CONTENT_LENGTH=MAX_UPLOAD_BYTES,
        JOB_TTL_SECONDS=JOB_TTL_SECONDS,
    )
    if test_config:
        app.config.update(test_config)
    app.config["DATA_FOLDER"] = Path(app.config["DATA_FOLDER"])
    app.config["DATA_FOLDER"].mkdir(parents=True, exist_ok=True)

    @app.after_request
    def add_security_headers(response):
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["Referrer-Policy"] = "same-origin"
        return response

    @app.errorhandler(RequestEntityTooLarge)
    def file_too_large(_error):
        return jsonify({"error": "文件不能超过 50MB"}), 413

    def remove_expired_jobs():
        cutoff = time.time() - app.config["JOB_TTL_SECONDS"]
        for job_dir in app.config["DATA_FOLDER"].iterdir():
            if job_dir.is_dir() and job_dir.stat().st_mtime < cutoff:
                shutil.rmtree(job_dir, ignore_errors=True)

    def save_upload():
        if "file" not in request.files:
            return None, (jsonify({"error": "未选择文件"}), 400)
        upload = request.files["file"]
        if not upload.filename:
            return None, (jsonify({"error": "未选择文件"}), 400)
        if not allowed_file(upload.filename):
            return None, (jsonify({"error": "仅支持 .docx 格式文件"}), 400)

        remove_expired_jobs()
        job_id = uuid.uuid4().hex
        job_dir = app.config["DATA_FOLDER"] / job_id
        job_dir.mkdir()
        upload_path = job_dir / "source.docx"
        upload.save(upload_path)
        if not is_docx_file(upload_path):
            shutil.rmtree(job_dir, ignore_errors=True)
            return None, (jsonify({"error": "文件不是有效的 Word .docx 文档"}), 400)
        return {"id": job_id, "dir": job_dir, "path": upload_path,
                "original_name": safe_original_name(upload.filename)}, None

    def load_config():
        template_id = request.form.get("template", "general")
        custom_config = request.form.get("custom_config")
        if template_id not in FORMAT_TEMPLATES:
            raise ValueError("未知的格式模板")
        if not custom_config:
            return template_id, None
        config = json.loads(custom_config)
        if not isinstance(config, dict):
            raise ValueError("自定义配置必须是对象")
        missing = set(FORMAT_TEMPLATES["general"]) - set(config)
        if missing:
            raise ValueError("自定义配置缺少必要字段")
        return template_id, config

    @app.route("/")
    def index():
        return render_template("index.html")

    @app.route("/api/health")
    def health():
        return jsonify({"success": True, "service": "paper-format-checker"})

    @app.route("/api/templates")
    def get_templates():
        return jsonify({"templates": get_template_list()})

    @app.route("/api/template/<template_id>")
    def get_template_detail(template_id):
        if template_id in FORMAT_TEMPLATES:
            return jsonify({"template": FORMAT_TEMPLATES[template_id]})
        return jsonify({"error": "模板不存在"}), 404

    @app.route("/api/font-options")
    def font_options():
        return jsonify(get_font_options())

    @app.route("/api/parse-requirements", methods=["POST"])
    def parse_requirements_api():
        data = request.get_json(silent=True) or {}
        text = str(data.get("text", "")).strip()
        if not text:
            return jsonify({"error": "请输入格式要求文本"}), 400
        if len(text) > 20_000:
            return jsonify({"error": "格式要求文本不能超过 20000 个字符"}), 400
        try:
            # Keep requirement parsing local. A format checker should never send
            # a student's assignment or thesis requirements to a remote model by default.
            result = parse_requirements(text)
            return jsonify({"success": True, "config": merge_with_default(result["config"]),
                            "confidence": result["confidence"], "unclear": result["unclear"],
                            "used_ai": result["used_ai"], "extracted_count": len(result["config"])})
        except Exception:
            app.logger.exception("Requirement parsing failed")
            return jsonify({"error": "解析失败，请检查输入后重试"}), 500

    @app.route("/api/check", methods=["POST"])
    def check():
        upload, error = save_upload()
        if error:
            return error
        try:
            template_id, config = load_config()
            checker = FormatChecker(template_id=template_id, custom_config=config)
            issues, stats = checker.check_file(upload["path"])
            report_data = generate_report(issues, stats, upload["original_name"])
            report_data["template"] = template_id
            report_data["template_name"] = FORMAT_TEMPLATES[template_id]["name"]
            (upload["dir"] / "report.html").write_text(generate_html_report(report_data), encoding="utf-8")
            return jsonify({"success": True, "job_id": upload["id"], "report": report_data,
                            "summary": checker.get_summary(), "report_url": f"/report/{upload['id']}",
                            "fixable_count": sum(issue.fixable for issue in issues)})
        except (ValueError, json.JSONDecodeError) as error:
            return jsonify({"error": str(error)}), 400
        except Exception:
            app.logger.exception("Document check failed")
            return jsonify({"error": "检查失败：文档可能损坏或包含不受支持的内容"}), 500

    @app.route("/api/fix", methods=["POST"])
    def fix():
        upload, error = save_upload()
        if error:
            return error
        try:
            template_id, config = load_config()
            checker = FormatChecker(template_id=template_id, custom_config=config)
            output_path = upload["dir"] / "fixed.docx"
            fix_records = checker.fix_file(upload["path"], output_path)
            issues_after, _ = FormatChecker(template_id, config).check_file(output_path)
            fix_stats = {}
            for record in fix_records:
                fix_stats[record.category] = fix_stats.get(record.category, 0) + 1
            return jsonify({"success": True, "job_id": upload["id"],
                            "download_url": f"/download/{upload['id']}",
                            "issues_before": request.form.get("issues_before", 0),
                            "issues_after": len(issues_after), "total_fixes": len(fix_records),
                            "fix_stats": fix_stats, "fix_records": [record.to_dict() for record in fix_records],
                            "message": f"修复完成：已修改 {len(fix_records)} 处，仍有 {len(issues_after)} 个需人工确认的问题。"})
        except (ValueError, json.JSONDecodeError) as error:
            return jsonify({"error": str(error)}), 400
        except Exception:
            app.logger.exception("Document fix failed")
            return jsonify({"error": "修复失败：文档可能损坏或包含不受支持的内容"}), 500

    @app.route("/api/polish", methods=["POST"])
    def polish():
        return jsonify({"error": "AI 润色在本地版本中默认关闭；请部署受控的模型代理后再启用。"}), 501

    @app.route("/report/<job_id>")
    def view_report(job_id):
        report_path = app.config["DATA_FOLDER"] / job_id / "report.html"
        if report_path.is_file() and len(job_id) == 32:
            return send_file(report_path, mimetype="text/html; charset=utf-8")
        return "报告不存在或已过期", 404

    @app.route("/download/<job_id>")
    def download_file(job_id):
        path = app.config["DATA_FOLDER"] / job_id / "fixed.docx"
        if path.is_file() and len(job_id) == 32:
            return send_file(path, as_attachment=True, download_name="论文_已修复.docx")
        return "文件不存在或已过期", 404

    return app


app = create_app()

if __name__ == "__main__":
    # Some managed browser environments cannot reach a process bound only to
    # loopback. Binding all local interfaces also supports classroom LAN demos.
    app.run(
        debug=False,
        host=os.getenv("PAPER_CHECKER_HOST", "127.0.0.1"),
        port=int(os.getenv("PAPER_CHECKER_PORT", "5000")),
    )

# -*- coding: utf-8 -*-
"""论文/作业格式检查与修复系统的 Flask 应用。"""

import json
import os
import re
import secrets
import shutil
import time
import uuid
import zipfile
from datetime import timedelta
from pathlib import Path

from flask import Flask, abort, g, jsonify, redirect, render_template, request, send_file, session, url_for
from werkzeug.exceptions import RequestEntityTooLarge

from utils.auth_store import (AVATARS, authenticate, create_user, delete_job_record, get_user,
                              list_jobs, normalize_email, prepare_auth, record_job,
                              update_profile, valid_email)
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
            entries = archive.infolist()
            if len(entries) > 5000 or sum(item.file_size for item in entries) > 200 * 1024 * 1024:
                return False
            if any(item.file_size > 50 * 1024 * 1024 for item in entries):
                return False
            names = {item.filename for item in entries}
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
    db_path, saved_secret = prepare_auth(app.config["DATA_FOLDER"])
    app.config["SECRET_KEY"] = app.config.get("SECRET_KEY") or saved_secret
    app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax",
                      PERMANENT_SESSION_LIFETIME=timedelta(hours=8))
    login_attempts = {}

    def csrf_token():
        if "csrf_token" not in session:
            session["csrf_token"] = secrets.token_urlsafe(32)
        return session["csrf_token"]

    app.jinja_env.globals["csrf_token"] = csrf_token

    @app.before_request
    def require_account():
        g.user = None
        user_id = session.get("user_id")
        if isinstance(user_id, int):
            g.user = get_user(db_path, user_id)
        if request.endpoint in {"static", "health", "login", "register"}:
            return None
        if g.user is None:
            if request.path.startswith("/api/"):
                return jsonify({"error": "请先登录"}), 401
            return redirect(url_for("login"))
        if request.method in {"POST", "PUT", "PATCH", "DELETE"}:
            supplied = request.headers.get("X-CSRF-Token") or request.form.get("csrf_token", "")
            if not secrets.compare_digest(supplied, csrf_token()):
                return jsonify({"error": "页面已过期，请刷新后重试"}), 403
        return None

    @app.context_processor
    def account_context():
        return {"current_user": getattr(g, "user", None), "avatar_options": AVATARS}

    def auth_page(mode, error="", status=200):
        return render_template("auth.html", mode=mode, error=error), status

    @app.route("/login", methods=["GET", "POST"])
    def login():
        if g.user:
            return redirect(url_for("index"))
        if request.method == "GET":
            return auth_page("login")
        if not secrets.compare_digest(request.form.get("csrf_token", ""), csrf_token()):
            return auth_page("login", "页面已过期，请刷新后重试", 403)
        email = normalize_email(request.form.get("email", ""))
        password = request.form.get("password", "")
        if len(email) > 254 or len(password) > 128:
            return auth_page("login", "邮箱或密码不正确", 401)
        key = (request.remote_addr or "", email)
        now = time.monotonic()
        attempts = [t for t in login_attempts.get(key, []) if now - t < 900]
        if len(attempts) >= 5:
            return auth_page("login", "尝试次数过多，请 15 分钟后重试", 429)
        user = authenticate(db_path, email, password)
        if user is None:
            attempts.append(now)
            login_attempts[key] = attempts
            return auth_page("login", "邮箱或密码不正确", 401)
        login_attempts.pop(key, None)
        session.clear()
        session["user_id"] = user["id"]
        session.permanent = True
        return redirect(url_for("index"))

    @app.route("/register", methods=["GET", "POST"])
    def register():
        if g.user:
            return redirect(url_for("index"))
        if request.method == "GET":
            return auth_page("register")
        if not secrets.compare_digest(request.form.get("csrf_token", ""), csrf_token()):
            return auth_page("register", "页面已过期，请刷新后重试", 403)
        email = normalize_email(request.form.get("email", ""))
        password = request.form.get("password", "")
        if not valid_email(email):
            return auth_page("register", "请输入有效的邮箱地址", 400)
        if not 10 <= len(password) <= 128:
            return auth_page("register", "密码需为 10～128 个字符", 400)
        user_id = create_user(db_path, email, password)
        if user_id is None:
            return auth_page("register", "该邮箱已注册", 409)
        session.clear()
        session["user_id"] = user_id
        session.permanent = True
        return redirect(url_for("index"))

    @app.route("/logout", methods=["POST"])
    def logout():
        session.clear()
        return redirect(url_for("login"))

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
            if (job_dir.is_dir() and re.fullmatch(r"[0-9a-f]{32}", job_dir.name)
                    and job_dir.stat().st_mtime < cutoff):
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
        (job_dir / "owner.json").write_text(json.dumps({"user_id": g.user["id"]}), encoding="utf-8")
        upload_path = job_dir / "source.docx"
        upload.save(upload_path)
        if not is_docx_file(upload_path):
            shutil.rmtree(job_dir, ignore_errors=True)
            return None, (jsonify({"error": "文件不是有效的 Word .docx 文档"}), 400)
        return {"id": job_id, "dir": job_dir, "path": upload_path,
                "original_name": safe_original_name(upload.filename)}, None

    def owned_job(job_id):
        if not re.fullmatch(r"[0-9a-f]{32}", job_id):
            abort(404)
        job_dir = app.config["DATA_FOLDER"] / job_id
        try:
            owner = json.loads((job_dir / "owner.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            abort(404)
        if owner.get("user_id") != g.user["id"]:
            abort(404)
        if job_dir.stat().st_mtime < time.time() - app.config["JOB_TTL_SECONDS"]:
            abort(404)
        return job_dir

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
        return render_template("index.html", active_page="checker")

    @app.route("/guide")
    def guide():
        return render_template("guide.html", active_page="guide", templates=FORMAT_TEMPLATES)

    @app.route("/similarity")
    def similarity():
        return render_template("similarity.html", active_page="similarity")

    @app.route("/profile", methods=["GET", "POST"])
    def profile():
        error = ""
        saved = request.args.get("saved") == "1"
        if request.method == "POST":
            display_name = request.form.get("display_name", "").strip()
            avatar = request.form.get("avatar", "whale")
            if not 1 <= len(display_name) <= 24 or any(ord(c) < 32 for c in display_name):
                error = "昵称需为 1～24 个可见字符"
            elif avatar not in AVATARS:
                error = "请选择列表中的头像"
            else:
                update_profile(db_path, g.user["id"], display_name, avatar)
                return redirect(url_for("profile", saved=1))
        return render_template("profile.html", active_page="profile", avatars=AVATARS,
                               error=error, saved=saved)

    @app.route("/history")
    def history():
        remove_expired_jobs()
        jobs = list_jobs(db_path, g.user["id"])
        for job in jobs:
            job_dir = app.config["DATA_FOLDER"] / job["job_id"]
            artifact = "report.html" if job["action"] == "check" else "fixed.docx"
            job["available"] = (job_dir / artifact).is_file()
        return render_template("history.html", active_page="history", jobs=jobs)

    @app.route("/history/delete/<job_id>", methods=["POST"])
    def delete_history(job_id):
        if not re.fullmatch(r"[0-9a-f]{32}", job_id):
            abort(404)
        job_dir = app.config["DATA_FOLDER"] / job_id
        if job_dir.exists():
            owned_job(job_id)
        if not delete_job_record(db_path, g.user["id"], job_id):
            abort(404)
        if job_dir.exists():
            shutil.rmtree(job_dir)
        return redirect(url_for("history"))

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
            record_job(db_path, job_id=upload["id"], user_id=g.user["id"],
                       filename=upload["original_name"], action="check",
                       template_name=FORMAT_TEMPLATES[template_id]["name"],
                       before_count=len(issues))
            return jsonify({"success": True, "job_id": upload["id"], "report": report_data,
                            "summary": checker.get_summary(), "report_url": f"/report/{upload['id']}",
                            "fixable_count": sum(issue.fixable for issue in issues)})
        except (ValueError, json.JSONDecodeError) as error:
            shutil.rmtree(upload["dir"], ignore_errors=True)
            return jsonify({"error": str(error)}), 400
        except Exception:
            shutil.rmtree(upload["dir"], ignore_errors=True)
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
            issues_before, _ = checker.check_file(upload["path"])
            output_path = upload["dir"] / "fixed.docx"
            fix_records = checker.fix_file(upload["path"], output_path)
            issues_after, _ = FormatChecker(template_id, config).check_file(output_path)
            fix_stats = {}
            for record in fix_records:
                fix_stats[record.category] = fix_stats.get(record.category, 0) + 1
            comparison = {
                "filename": upload["original_name"],
                "template_name": FORMAT_TEMPLATES[template_id]["name"],
                "before_count": len(issues_before),
                "after_count": len(issues_after),
                "fix_records": [record.to_dict() for record in fix_records],
                "remaining_issues": [issue.to_dict() for issue in issues_after],
            }
            (upload["dir"] / "comparison.json").write_text(
                json.dumps(comparison, ensure_ascii=False), encoding="utf-8")
            record_job(db_path, job_id=upload["id"], user_id=g.user["id"],
                       filename=upload["original_name"], action="fix",
                       template_name=FORMAT_TEMPLATES[template_id]["name"],
                       before_count=len(issues_before), after_count=len(issues_after),
                       fix_count=len(fix_records))
            return jsonify({"success": True, "job_id": upload["id"],
                            "download_url": f"/download/{upload['id']}",
                            "comparison_url": f"/comparison/{upload['id']}",
                            "issues_before": len(issues_before),
                            "issues_after": len(issues_after), "total_fixes": len(fix_records),
                            "fix_stats": fix_stats, "fix_records": [record.to_dict() for record in fix_records],
                            "message": f"修复完成：已修改 {len(fix_records)} 处，仍有 {len(issues_after)} 个需人工确认的问题。"})
        except (ValueError, json.JSONDecodeError) as error:
            shutil.rmtree(upload["dir"], ignore_errors=True)
            return jsonify({"error": str(error)}), 400
        except Exception:
            shutil.rmtree(upload["dir"], ignore_errors=True)
            app.logger.exception("Document fix failed")
            return jsonify({"error": "修复失败：文档可能损坏或包含不受支持的内容"}), 500

    @app.route("/api/polish", methods=["POST"])
    def polish():
        return jsonify({"error": "AI 润色在本地版本中默认关闭；请部署受控的模型代理后再启用。"}), 501

    @app.route("/report/<job_id>")
    def view_report(job_id):
        report_path = owned_job(job_id) / "report.html"
        if report_path.is_file():
            return send_file(report_path, mimetype="text/html; charset=utf-8")
        return "报告不存在或已过期", 404

    @app.route("/comparison/<job_id>")
    def comparison(job_id):
        path = owned_job(job_id) / "comparison.json"
        if not path.is_file():
            abort(404)
        data = json.loads(path.read_text(encoding="utf-8"))
        return render_template("comparison.html", active_page="comparison", job_id=job_id, data=data)

    @app.route("/download/<job_id>")
    def download_file(job_id):
        path = owned_job(job_id) / "fixed.docx"
        if path.is_file():
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

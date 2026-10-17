# -*- coding: utf-8 -*-
"""
检查报告生成模块
生成HTML格式的检查报告
"""

import json
from html import escape
from datetime import datetime


def generate_report(issues, stats, filename):
    """生成HTML检查报告"""
    summary = {
        "total_issues": len(issues),
        "error_count": sum(1 for i in issues if i.severity == "error"),
        "warning_count": sum(1 for i in issues if i.severity == "warning"),
        "info_count": sum(1 for i in issues if i.severity == "info"),
    }

    # 按类别分组
    categories = {}
    for issue in issues:
        cat = issue.category
        if cat not in categories:
            categories[cat] = []
        categories[cat].append(issue.to_dict())

    # 按严重程度排序
    severity_order = {"error": 0, "warning": 1, "info": 2}
    for cat in categories:
        categories[cat].sort(key=lambda x: severity_order.get(x["severity"], 3))

    report_data = {
        "filename": filename,
        "check_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "summary": summary,
        "stats": stats,
        "categories": categories,
        "issues": [i.to_dict() for i in issues]
    }

    return report_data


def generate_html_report(report_data):
    """生成完整HTML报告页面"""
    issues_json = json.dumps(report_data, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>格式检查报告 - {escape(report_data['filename'])}</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: -apple-system, "Microsoft YaHei", sans-serif; background: #f5f7fa; color: #333; }}
        .container {{ max-width: 900px; margin: 0 auto; padding: 20px; }}
        .header {{ background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 12px; margin-bottom: 20px; }}
        .header h1 {{ font-size: 24px; margin-bottom: 8px; }}
        .header .meta {{ font-size: 14px; opacity: 0.9; }}
        .summary-cards {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 20px; }}
        .card {{ background: white; border-radius: 10px; padding: 20px; text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }}
        .card .number {{ font-size: 32px; font-weight: bold; }}
        .card .label {{ font-size: 13px; color: #888; margin-top: 4px; }}
        .card.error .number {{ color: #e74c3c; }}
        .card.warning .number {{ color: #f39c12; }}
        .card.info .number {{ color: #3498db; }}
        .card.total .number {{ color: #2c3e50; }}
        .stats-bar {{ background: white; border-radius: 10px; padding: 16px 20px; margin-bottom: 20px; display: flex; flex-wrap: wrap; gap: 20px; font-size: 14px; }}
        .stats-bar span {{ color: #666; }}
        .stats-bar strong {{ color: #333; }}
        .category-section {{ background: white; border-radius: 10px; margin-bottom: 16px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.06); }}
        .category-header {{ padding: 14px 20px; background: #f8f9fa; border-bottom: 1px solid #eee; cursor: pointer; display: flex; justify-content: space-between; align-items: center; }}
        .category-header h3 {{ font-size: 16px; }}
        .category-header .count {{ background: #e74c3c; color: white; font-size: 12px; padding: 2px 8px; border-radius: 10px; }}
        .issue-list {{ padding: 0; }}
        .issue-item {{ padding: 14px 20px; border-bottom: 1px solid #f0f0f0; }}
        .issue-item:last-child {{ border-bottom: none; }}
        .issue-header {{ display: flex; align-items: center; gap: 8px; margin-bottom: 6px; }}
        .severity-badge {{ font-size: 11px; padding: 2px 8px; border-radius: 4px; font-weight: bold; }}
        .severity-badge.error {{ background: #fde8e8; color: #e74c3c; }}
        .severity-badge.warning {{ background: #fef3e2; color: #f39c12; }}
        .severity-badge.info {{ background: #e8f4fd; color: #3498db; }}
        .issue-location {{ font-size: 12px; color: #999; }}
        .issue-desc {{ font-size: 14px; color: #333; margin-bottom: 4px; }}
        .issue-suggestion {{ font-size: 13px; color: #27ae60; background: #eafaf1; padding: 6px 10px; border-radius: 6px; }}
        .issue-suggestion::before {{ content: "💡 建议："; font-weight: bold; }}
        .no-issues {{ text-align: center; padding: 40px; color: #27ae60; font-size: 18px; }}
        .back-btn {{ display: inline-block; margin-top: 20px; padding: 10px 24px; background: #667eea; color: white; text-decoration: none; border-radius: 8px; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📄 格式检查报告</h1>
            <div class="meta">文件：{escape(report_data['filename'])} | 检查时间：{escape(report_data['check_time'])}</div>
        </div>

        <div class="summary-cards">
            <div class="card total">
                <div class="number">{report_data['summary']['total_issues']}</div>
                <div class="label">问题总数</div>
            </div>
            <div class="card error">
                <div class="number">{report_data['summary']['error_count']}</div>
                <div class="label">严重问题</div>
            </div>
            <div class="card warning">
                <div class="number">{report_data['summary']['warning_count']}</div>
                <div class="label">警告</div>
            </div>
            <div class="card info">
                <div class="number">{report_data['summary']['info_count']}</div>
                <div class="label">建议</div>
            </div>
        </div>

        <div class="stats-bar">
            <span>总段落：<strong>{report_data['stats']['total_paragraphs']}</strong></span>
            <span>总字数：<strong>{report_data['stats']['total_words']}</strong></span>
            <span>一级标题：<strong>{report_data['stats']['h1_count']}</strong></span>
            <span>二级标题：<strong>{report_data['stats']['h2_count']}</strong></span>
            <span>三级标题：<strong>{report_data['stats']['h3_count']}</strong></span>
            <span>正文段落：<strong>{report_data['stats']['body_count']}</strong></span>
            <span>参考文献：<strong>{report_data['stats']['ref_count']}</strong></span>
            <span>表格：<strong>{report_data['stats']['table_count']}</strong></span>
        </div>
    """

    if not report_data["issues"]:
        html += """
        <div class="category-section">
            <div class="no-issues">🎉 未检测到格式问题，格式规范！</div>
        </div>
        """
    else:
        for cat, cat_issues in report_data["categories"].items():
            html += f"""
        <div class="category-section">
            <div class="category-header" onclick="this.nextElementSibling.style.display=this.nextElementSibling.style.display==='none'?'block':'none'">
                <h3>{escape(cat)}</h3>
                <span class="count">{len(cat_issues)}</span>
            </div>
            <div class="issue-list">
            """
            for issue in cat_issues:
                severity_label = {"error": "严重", "warning": "警告", "info": "建议"}.get(issue["severity"], issue["severity"])
                html += f"""
                <div class="issue-item">
                    <div class="issue-header">
                        <span class="severity-badge {issue['severity']}">{severity_label}</span>
                        <span class="issue-location">{escape(issue['location'])}</span>
                    </div>
                    <div class="issue-desc">{escape(issue['description'])}</div>
                    <div class="issue-suggestion">{escape(issue['suggestion'])}</div>
                </div>
                """
            html += """
            </div>
        </div>
            """

    html += """
        <a href="/" class="back-btn">← 返回重新上传</a>
    </div>
</body>
</html>"""

    return html

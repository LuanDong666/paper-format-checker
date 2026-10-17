# -*- coding: utf-8 -*-
"""
格式要求智能解析模块
将自然语言描述的格式要求解析为结构化配置参数
支持规则匹配（免费）和AI增强解析（需API Key）
"""

import re
import json
import urllib.request


# 字号名称到pt的映射
SIZE_NAME_MAP = {
    "初号": 42, "小初": 36,
    "一号": 26, "小一": 24,
    "二号": 22, "小二": 18,
    "三号": 16, "小三": 15,
    "四号": 14, "小四": 12,
    "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5,
}

# 纸张大小映射（宽, 高）单位cm
PAPER_SIZE_MAP = {
    "A4": (21.0, 29.7),
    "A3": (29.7, 42.0),
    "A5": (14.8, 21.0),
    "B4": (25.0, 35.3),
    "B5": (17.6, 25.0),
    "16开": (18.4, 26.0),
    "32开": (13.0, 18.4),
}

# 常见中文字体
CN_FONTS = ["宋体", "黑体", "楷体", "仿宋", "微软雅黑", "隶书", "幼圆",
            "华文行楷", "华文新魏", "华文中宋", "方正姚体", "方正舒体"]
EN_FONTS = ["Times New Roman", "Arial", "Calibri", "Cambria", "Georgia",
            "Verdana", "Courier New", "Tahoma", "Garamond"]


def parse_size(text):
    """从文本中提取字号，返回pt值。优先匹配带'小'的字号（如小四优先于四号）"""
    # 先匹配带"小"的字号
    small_sizes = [(k, v) for k, v in SIZE_NAME_MAP.items() if k.startswith("小")]
    for name, pt in small_sizes:
        if name in text:
            return pt, name
    # 再匹配其他字号，按长度降序
    other_sizes = sorted(
        [(k, v) for k, v in SIZE_NAME_MAP.items() if not k.startswith("小")],
        key=lambda x: len(x[0]), reverse=True
    )
    for name, pt in other_sizes:
        if name in text:
            return pt, name
    # 匹配数字+pt
    m = re.search(r'(\d+\.?\d*)\s*(?:pt|磅|号字)', text, re.IGNORECASE)
    if m:
        return float(m.group(1)), f"{m.group(1)}pt"
    return None, None


def parse_requirements_rule_based(text):
    """
    基于规则匹配解析格式要求（免费，不需要API Key）
    返回配置字典、置信度信息、未明确项列表
    """
    config = {}
    confidence = {}  # 每个参数的置信度 high/medium/low
    unclear = []    # 未明确的项

    text_lower = text.lower()

    # 1. 纸张大小
    for paper_name, (w, h) in PAPER_SIZE_MAP.items():
        if paper_name in text:
            config["paper_width"] = w
            config["paper_height"] = h
            confidence["paper_size"] = "high"
            break
    if "paper_size" not in confidence:
        unclear.append("纸张大小（默认A4）")

    # 2. 页边距（先分别匹配上下左右，再用统一值填充缺失的）
    margin_keys = ["top_margin", "bottom_margin", "left_margin", "right_margin"]
    for direction, key in [("上", "top_margin"), ("下", "bottom_margin"),
                             ("左", "left_margin"), ("右", "right_margin")]:
        m = re.search(rf'{direction}(?:边距)?[^0-9]*(\d+\.?\d*)\s*(?:cm|厘米|公分)', text)
        if m:
            config[key] = float(m.group(1))
            confidence[key] = "high"
    # 如果有统一的页边距说明，填充未单独指定的
    m = re.search(r'页边距[^0-9]*(\d+\.?\d*)\s*(?:cm|厘米|公分)', text)
    if m:
        v = float(m.group(1))
        for key in margin_keys:
            if key not in config:
                config[key] = v
                confidence[key] = "medium"
        confidence["margins"] = "high"
    if not any(k in config for k in margin_keys):
        unclear.append("页边距（默认上下2.54cm，左右3.17cm）")

    # 3. 正文字体
    for font in CN_FONTS:
        # 匹配"正文用X字体"或"正文X体"
        if re.search(rf'正文[^。；\n]*{font}', text) or re.search(rf'{font}[^。；\n]*正文', text):
            config["body_font"] = font
            confidence["body_font"] = "high"
            break
    if "body_font" not in confidence:
        # 全局字体要求
        for font in CN_FONTS:
            if re.search(rf'字体[^。；\n]*{font}|{font}[^。；\n]*字体', text):
                config["body_font"] = font
                confidence["body_font"] = "medium"
                break
    if "body_font" not in confidence:
        unclear.append("正文字体（默认宋体）")

    # 4. 正文英文/数字字体
    for font in EN_FONTS:
        if re.search(rf'英文[^。；\n]*{re.escape(font)}|数字[^。；\n]*{re.escape(font)}', text, re.IGNORECASE):
            config["body_en_font"] = font
            confidence["body_en_font"] = "high"
            break
    if "body_en_font" not in confidence:
        unclear.append("英文/数字字体（默认Times New Roman）")

    # 5. 正文字号
    # 匹配"正文X号"或"正文Xpt"
    body_section = re.search(r'正文[^。；\n]{0,30}', text)
    if body_section:
        size_pt, size_name = parse_size(body_section.group())
        if size_pt:
            config["body_size"] = size_pt
            confidence["body_size"] = "high"
    if "body_size" not in confidence:
        # 全局字号
        size_pt, size_name = parse_size(text)
        if size_pt and "字号" in text:
            config["body_size"] = size_pt
            confidence["body_size"] = "low"
    if "body_size" not in confidence:
        unclear.append("正文字号（默认小四12pt）")

    # 6. 行距（支持"X倍行距"和"行距X倍"两种顺序）
    m = re.search(r'(\d+\.?\d*)\s*倍[^。；\n]{0,5}行距', text)
    if not m:
        m = re.search(r'行距[^0-9]*(\d+\.?\d*)\s*倍', text)
    if m:
        config["body_line_spacing"] = float(m.group(1))
        confidence["line_spacing"] = "high"
    else:
        m = re.search(r'行距[^0-9]*(\d+\.?\d*)\s*(?:磅|pt)', text)
        if m:
            config["body_line_spacing"] = float(m.group(1)) / 12
            confidence["line_spacing"] = "medium"
        elif "单倍行距" in text or "固定行距" in text:
            config["body_line_spacing"] = 1.0
            confidence["line_spacing"] = "medium"
    if "line_spacing" not in confidence:
        unclear.append("行距（默认1.5倍）")

    # 7. 首行缩进
    if re.search(r'首行缩进|缩进\s*2\s*字符|缩进两字符', text):
        m = re.search(r'缩进[^0-9]*(\d+)\s*字符', text)
        chars = int(m.group(1)) if m else 2
        config["body_first_line_indent"] = (config.get("body_size", 12)) * chars
        confidence["first_line_indent"] = "high"
    elif "不缩进" in text or "顶格" in text:
        config["body_first_line_indent"] = 0
        confidence["first_line_indent"] = "high"
    else:
        unclear.append("首行缩进（默认2字符）")

    # 8. 一级标题
    h1_section = re.search(r'(?:一级标题|标题一|第一级标题|章标题)[^。；\n]{0,50}', text)
    if h1_section:
        sec = h1_section.group()
        for font in CN_FONTS:
            if font in sec:
                config["h1_font"] = font
                confidence["h1_font"] = "high"
                break
        size_pt, size_name = parse_size(sec)
        if size_pt:
            config["h1_size"] = size_pt
            confidence["h1_size"] = "high"
        if "居中" in sec or "中心" in sec:
            config["h1_alignment"] = "center"
            confidence["h1_alignment"] = "high"
        elif "左对齐" in sec or "顶格" in sec:
            config["h1_alignment"] = "left"
            confidence["h1_alignment"] = "high"

    # 9. 二级标题
    h2_section = re.search(r'(?:二级标题|标题二|第二级标题|节标题)[^。；\n]{0,50}', text)
    if h2_section:
        sec = h2_section.group()
        for font in CN_FONTS:
            if font in sec:
                config["h2_font"] = font
                confidence["h2_font"] = "high"
                break
        size_pt, size_name = parse_size(sec)
        if size_pt:
            config["h2_size"] = size_pt
            confidence["h2_size"] = "high"

    # 10. 参考文献字号
    ref_section = re.search(r'参考文献[^。；\n]{0,40}', text)
    if ref_section:
        sec = ref_section.group()
        size_pt, size_name = parse_size(sec)
        if size_pt:
            config["ref_size"] = size_pt
            confidence["ref_size"] = "high"

    # 11. 加粗要求
    if re.search(r'标题[^。；\n]*加粗|加粗[^。；\n]*标题', text):
        config["h1_bold"] = True
        config["h2_bold"] = True
        confidence["heading_bold"] = "medium"

    return config, confidence, unclear


def parse_requirements_ai(text, api_key, api_base="https://api.deepseek.com",
                          model="deepseek-chat"):
    """
    使用AI大模型深度解析格式要求（需要API Key）
    返回结构化配置
    """
    prompt = f"""你是一个论文格式解析专家。请从以下格式要求文本中提取结构化的格式参数。

要求文本：
\"\"\"
{text}
\"\"\"

请提取以下参数（如果文本中没有明确说明，该字段填null）：
- paper_width: 纸张宽度cm（A4为21.0）
- paper_height: 纸张高度cm（A4为29.7）
- top_margin: 上边距cm
- bottom_margin: 下边距cm
- left_margin: 左边距cm
- right_margin: 右边距cm
- body_font: 正文中文字体（如宋体、黑体）
- body_en_font: 正文英文/数字字体（如Times New Roman）
- body_size: 正文字号pt（小四=12，四号=14，五号=10.5）
- body_line_spacing: 正文行距倍数（如1.5）
- body_first_line_indent: 首行缩进pt（2字符=字号*2）
- h1_font: 一级标题中文字体
- h1_en_font: 一级标题英文/数字字体
- h1_size: 一级标题字号pt
- h1_bold: 一级标题是否加粗（true/false）
- h1_alignment: 一级标题对齐（center/left）
- h2_font: 二级标题中文字体
- h2_size: 二级标题字号pt
- h2_bold: 二级标题是否加粗
- ref_size: 参考文献字号pt

只返回JSON，不要解释，不要markdown代码块。"""

    payload = json.dumps({
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.1,
    })
    req = urllib.request.Request(
        f"{api_base}/chat/completions",
        data=payload.encode('utf-8'),
        headers={
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {api_key}'
        }
    )
    resp = urllib.request.urlopen(req, timeout=30)
    result = json.loads(resp.read())
    ai_text = result['choices'][0]['message']['content']

    # 清理可能的markdown代码块
    ai_text = ai_text.strip()
    if ai_text.startswith('```'):
        ai_text = re.sub(r'^```\w*\n', '', ai_text)
        ai_text = re.sub(r'\n```$', '', ai_text)

    try:
        config = json.loads(ai_text)
        # 过滤掉null值
        config = {k: v for k, v in config.items() if v is not None}
        return config
    except json.JSONDecodeError:
        return {}


def parse_requirements(text, api_key=None, api_base="https://api.deepseek.com",
                       model="deepseek-chat"):
    """
    统一入口：先规则匹配，有API Key则用AI增强
    返回 {config, confidence, unclear, used_ai}
    """
    # 规则匹配
    rule_config, confidence, unclear = parse_requirements_rule_based(text)

    used_ai = False
    if api_key:
        try:
            ai_config = parse_requirements_ai(text, api_key, api_base, model)
            # AI结果覆盖规则结果（AI更准确）
            for k, v in ai_config.items():
                rule_config[k] = v
                confidence[k] = "high"
            # 更新未明确项
            unclear = [u for u in unclear if not any(
                keyword in u for keyword in
                [k.replace("_", " ") for k in ai_config.keys()]
            )]
            used_ai = True
        except Exception as e:
            # AI失败时回退到规则匹配
            pass

    return {
        "config": rule_config,
        "confidence": confidence,
        "unclear": unclear,
        "used_ai": used_ai,
    }


def merge_with_default(config):
    """将解析到的配置与默认模板合并，补全缺失项"""
    default = {
        "name": "智能解析",
        "paper_width": 21.0, "paper_height": 29.7,
        "top_margin": 2.54, "bottom_margin": 2.54,
        "left_margin": 3.17, "right_margin": 3.17,
        "body_font": "宋体", "body_en_font": "Times New Roman",
        "body_size": 12, "body_line_spacing": 1.5,
        "body_first_line_indent": 24, "body_alignment": "justify",
        "h1_font": "黑体", "h1_en_font": "Times New Roman",
        "h1_size": 16, "h1_bold": False, "h1_alignment": "center",
        "h1_space_before": 12, "h1_space_after": 12,
        "h2_font": "黑体", "h2_en_font": "Times New Roman",
        "h2_size": 14, "h2_bold": False, "h2_alignment": "left",
        "h2_space_before": 6, "h2_space_after": 6,
        "h3_font": "黑体", "h3_en_font": "Times New Roman",
        "h3_size": 12, "h3_bold": True, "h3_alignment": "left",
        "ref_font": "宋体", "ref_en_font": "Times New Roman",
        "ref_size": 10.5, "ref_line_spacing": 1.25,
        "keyword_font": "黑体", "keyword_en_font": "Times New Roman",
        "keyword_size": 12,
    }
    # 处理首行缩进：如果配置了body_size但没配indent，自动计算
    if "body_size" in config and "body_first_line_indent" not in config:
        config["body_first_line_indent"] = config["body_size"] * 2
    default.update(config)
    return default

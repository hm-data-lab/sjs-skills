#!/usr/bin/env python3
"""
大纲解析器 — 统一的 Markdown 大纲解析

消除 gen_ppt.py、content_layout.py、gen_visuals.py 中的 3 处重复解析逻辑。
所有脚本共用此模块。

输出格式：
{
    "title": str,
    "meta": str,
    "sections": [
        {
            "heading": str,
            "slides": [
                {"subtitle": str|None, "points": [str], "chart": str|None}
            ]
        }
    ],
    "notes_map": {heading: notes_text}
}
"""

import re


def parse_outline(md_path: str) -> dict:
    """
    解析大纲 Markdown 文件，返回结构化数据。

    Args:
        md_path: 大纲 Markdown 文件路径

    Returns:
        dict: 包含 title, meta, sections, notes_map
    """
    with open(md_path, "r", encoding="utf-8") as f:
        content = f.read()

    lines = content.split("\n")
    result = {"title": "", "meta": "", "sections": [], "notes_map": {}}

    in_notes_section = False
    current_section = None
    current_slide = None
    last_chart_hint = None

    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        if not line:
            i += 1
            continue

        stripped = line.strip()

        # 图表标注 <!-- chart: xxx -->
        if stripped.startswith("<!--"):
            chart_match = re.search(r"chart:\s*(\w+)", stripped)
            if chart_match:
                last_chart_hint = chart_match.group(1)
            i += 1
            continue

        # ── 备注区 ──
        if in_notes_section:
            if stripped.startswith("## ") and not stripped.startswith("### "):
                heading = stripped[3:].strip()
                notes_lines = []
                i += 1
                while i < len(lines):
                    nl = lines[i].rstrip()
                    if nl.startswith("## ") or nl.startswith("# "):
                        break
                    if nl.startswith("- "):
                        notes_lines.append(nl[2:].strip())
                    elif nl and not nl.strip().startswith("<!--"):
                        notes_lines.append(nl.strip())
                    i += 1
                result["notes_map"][heading] = "\n".join(notes_lines)
            else:
                i += 1
            continue

        # ── 一级标题 ──
        if stripped.startswith("# ") and not stripped.startswith("## "):
            result["title"] = stripped[2:].strip()
            i += 1
            continue

        # ── 元信息行 ──
        if stripped.startswith("> ") and "产品线" in stripped:
            result["meta"] = stripped[2:].strip()
            i += 1
            continue

        # ── 备注区入口 ──
        if stripped.startswith("## 备注区") or stripped == "## 备注":
            in_notes_section = True
            i += 1
            continue

        # ── 二级标题（章节） ──
        if stripped.startswith("## ") and not stripped.startswith("### "):
            heading = stripped[3:].strip()
            current_section = {"heading": heading, "slides": []}
            result["sections"].append(current_section)
            current_slide = None
            last_chart_hint = None
            i += 1
            continue

        # ── 三级标题（子页面） ──
        if stripped.startswith("### "):
            subtitle = stripped[4:].strip()
            current_slide = {"subtitle": subtitle, "points": [], "chart": last_chart_hint}
            if current_section is not None:
                current_section["slides"].append(current_slide)
            last_chart_hint = None
            i += 1
            continue

        # ── 表格行 ──
        if stripped.startswith("|"):
            if current_slide is None and current_section is not None:
                current_slide = {"subtitle": None, "points": [], "chart": last_chart_hint}
                current_section["slides"].append(current_slide)
                last_chart_hint = None
            table_rows = []
            while i < len(lines):
                tl = lines[i].rstrip().strip()
                if tl.startswith("|"):
                    if not re.match(r"^\|[\s\-:|]+\|$", tl):
                        inner = tl.strip("|")
                        cells = [c.strip() for c in inner.split("|")]
                        table_rows.append(cells)
                    i += 1
                else:
                    break
            if table_rows and current_slide is not None:
                for row in table_rows:
                    current_slide["points"].append(" | ".join(row))
            continue

        # ── 列表项 ──
        if stripped.startswith("- "):
            point = stripped[2:].strip()
            if current_slide is None and current_section is not None:
                current_slide = {"subtitle": None, "points": [], "chart": last_chart_hint}
                current_section["slides"].append(current_slide)
                last_chart_hint = None
            if current_slide is not None:
                current_slide["points"].append(point)
            i += 1
            continue

        # ── 普通文本行 ──
        if current_slide is not None and not stripped.startswith("#"):
            if stripped and stripped != "...":
                current_slide["points"].append(stripped)

        i += 1

    return result


def parse_outline_simple(md_path: str) -> dict:
    """
    简化版大纲解析 — 只返回 {章节标题: [子页面列表]}，不含备注。

    用于 content_layout.py 和 gen_visuals.py 的图表推断。
    """
    full = parse_outline(md_path)
    sections = {}
    for section in full["sections"]:
        heading = section["heading"]
        if heading.startswith("备注"):
            continue
        sections[heading] = section["slides"]
    return sections

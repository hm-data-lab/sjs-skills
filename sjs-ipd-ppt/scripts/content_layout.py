#!/usr/bin/env python3
"""
内容布局增强（方案2改进版）
原则：不改背景，只在模板上添加清晰的内容布局
- 卡片/色块展示要点
- 关键数字突出
- 页面填满不留空白
"""

import sys
import re
from pathlib import Path
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# 品牌颜色
PRIMARY = RGBColor(0x00, 0x70, 0xC0)
ACCENT = RGBColor(0x00, 0xA3, 0xE0)
ORANGE = RGBColor(0xFF, 0x8C, 0x00)
GREEN = RGBColor(0x10, 0xB9, 0x81)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def add_card(slide, left, top, width, height, title, points, color):
    """添加一个内容卡片（半透明背景 + 色条 + 标题 + 要点）"""
    # 半透明背景
    bg = slide.shapes.add_shape(1, left, top, width, height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = WHITE
    bg.fill.fore_color.brightness = 0.12
    bg.line.fill.background()

    # 顶部色条
    bar = slide.shapes.add_shape(1, left, top, width, Pt(4))
    bar.fill.solid()
    bar.fill.fore_color.rgb = color
    bar.line.fill.background()

    # 标题
    txBox = slide.shapes.add_textbox(left + Pt(12), top + Pt(10), width - Pt(24), Pt(22))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    run = p.add_run()
    run.font.size = Pt(14)
    run.font.name = "Microsoft YaHei"
    run.font.bold = True
    run.font.color.rgb = color

    # 要点
    py = top + Pt(38)
    for pt in points[:4]:
        txBox = slide.shapes.add_textbox(left + Pt(12), py, width - Pt(24), Pt(18))
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = f"• {pt}"
        run = p.add_run()
        run.font.size = Pt(11)
        run.font.name = "Microsoft YaHei"
        run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x2E)
        py += Pt(20)


def add_kpi_card(slide, left, top, width, height, number, label, color):
    """添加 KPI 数字卡片"""
    # 背景
    bg = slide.shapes.add_shape(1, left, top, width, height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = WHITE
    bg.fill.fore_color.brightness = 0.12
    bg.line.fill.background()

    # 顶部色条
    bar = slide.shapes.add_shape(1, left, top, width, Pt(5))
    bar.fill.solid()
    bar.fill.fore_color.rgb = color
    bar.line.fill.background()

    # 大数字
    txBox = slide.shapes.add_textbox(left, top + Pt(20), width, Pt(45))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = number
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.font.size = Pt(30)
    run.font.name = "Arial"
    run.font.bold = True
    run.font.color.rgb = color

    # 标签
    txBox = slide.shapes.add_textbox(left + Pt(8), top + Pt(70), width - Pt(16), Pt(40))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = label
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.font.size = Pt(11)
    run.font.name = "Microsoft YaHei"
    run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x2E)


# ── 解析 ──

def parse_outline(md_path):
    with open(md_path, "r", encoding="utf-8") as f:
        content = f.read()
    sections = {}
    current_heading = ""
    current_sub = ""
    current_points = []
    for line in content.split("\n"):
        if line.startswith("## ") and not line.startswith("### "):
            if current_sub and current_points:
                sections.setdefault(current_heading, []).append({"subtitle": current_sub, "points": current_points})
            elif current_points and not current_sub:
                sections.setdefault(current_heading, []).append({"subtitle": None, "points": current_points})
            current_heading = line[3:].strip()
            current_sub = ""
            current_points = []
        elif line.startswith("### "):
            if current_sub and current_points:
                sections.setdefault(current_heading, []).append({"subtitle": current_sub, "points": current_points})
            elif current_points and not current_sub:
                sections.setdefault(current_heading, []).append({"subtitle": None, "points": current_points})
            current_sub = line[4:].strip()
            current_points = []
        elif line.startswith("- ") and current_sub:
            current_points.append(line[2:].strip())
        elif line.startswith("- ") and not current_sub and current_heading:
            current_points.append(line[2:].strip())
        elif current_sub and line.strip() and not line.startswith("#") and not line.startswith(">") and not line.startswith("|") and not line.startswith("<!"):
            if not current_points:
                current_points.append(line.strip())
    if current_sub and current_points:
        sections.setdefault(current_heading, []).append({"subtitle": current_sub, "points": current_points})
    elif current_points and not current_sub:
        sections.setdefault(current_heading, []).append({"subtitle": None, "points": current_points})
    return sections


# ── 主逻辑 ──

def enhance_content_pages(prs, outline):
    """内容页（2-8）：在模板上添加三列卡片"""
    colors = [PRIMARY, ACCENT, ORANGE]
    for i in range(2, 9):
        slide = prs.slides[i]
        shapes = list(slide.shapes)
        if len(shapes) < 2:
            continue

        subtitle = shapes[1].text_frame.text.strip().split("\n")[0] if shapes[1].has_text_frame else ""

        matched_points = []
        for heading, subs in outline.items():
            for sub in subs:
                if sub.get("subtitle") and sub["subtitle"] in subtitle:
                    matched_points = sub["points"]
                    break
            if matched_points:
                break

        if not matched_points:
            continue

        # 保存原文到备注
        if len(shapes) > 2 and shapes[2].has_text_frame:
            original_text = shapes[2].text_frame.text.strip()
            if original_text:
                notes = slide.notes_slide.notes_text_frame.text if slide.has_notes_slide else ""
                slide.notes_slide.notes_text_frame.text = (notes + "\n\n【原始内容】\n" + original_text).strip()

        # 清空正文 text box（Shape [2]）
        if len(shapes) > 2 and shapes[2].has_text_frame:
            tf = shapes[2].text_frame
            while len(tf.paragraphs) > 1:
                p = tf.paragraphs[-1]
                p._element.getparent().remove(p._element)
            tf.paragraphs[0].text = ""

        # 在空白区域添加三列卡片
        # 页面 960x540，子标题在 y=79，正文区从 y=130 开始
        card_w = Pt(280)
        card_h = Pt(160)
        gap = Pt(16)
        start_x = Pt(53)
        start_y = Pt(130)

        group_size = max(1, len(matched_points) // 3)
        for idx in range(0, min(len(matched_points), 9), group_size):
            group = matched_points[idx:idx + group_size]
            card_idx = idx // group_size
            if card_idx >= 3:
                break
            x = start_x + card_idx * (card_w + gap)
            color = colors[card_idx % len(colors)]
            title = group[0][:15]
            add_card(slide, x, start_y, card_w, card_h, title, group, color)


def enhance_value_page(prs, outline):
    """项目价值页（9）：添加四维度卡片"""
    slide = prs.slides[9]
    value_subs = outline.get("三、项目价值", [])
    if not value_subs:
        return

    # 清空现有内容（保留标题）
    shapes = list(slide.shapes)
    for shape in shapes[1:]:
        sp = shape._element
        sp.getparent().remove(sp)

    # 2x2 网格布局
    card_w = Pt(420)
    card_h = Pt(170)
    gap_x = Pt(16)
    gap_y = Pt(14)
    start_x = Pt(53)
    start_y = Pt(110)
    colors = [PRIMARY, ORANGE, ACCENT, GREEN]

    for idx, sub in enumerate(value_subs[:4]):
        row = idx // 2
        col = idx % 2
        x = start_x + col * (card_w + gap_x)
        y = start_y + row * (card_h + gap_y)
        title = sub.get("subtitle", "")
        points = sub.get("points", [])
        if title:
            add_card(slide, x, y, card_w, card_h, title, points, colors[idx])


def enhance_kpi_page(prs, outline):
    """产出收益页（15）：添加 KPI 大数字卡片"""
    slide = prs.slides[15]
    kpi_points = []
    for sub in outline.get("七、产出和收益", []):
        for pt in sub.get("points", []):
            if "⭐" in pt:
                kpi_points.append(pt)

    if not kpi_points:
        return

    # 清空现有内容（保留标题）
    shapes = list(slide.shapes)
    for shape in shapes[1:]:
        sp = shape._element
        sp.getparent().remove(sp)

    # 4 个 KPI 卡片
    kpis = []
    for pt in kpi_points[:4]:
        match = re.search(r"\*\*(.+?)\*\*[：:](.+)", pt)
        if match:
            kpis.append((match.group(1), match.group(2)[:20]))
        else:
            kpis.append((pt[:15], ""))

    card_w = Pt(210)
    card_h = Pt(150)
    gap = Pt(16)
    start_x = Pt(53)
    start_y = Pt(130)
    colors = [ORANGE, PRIMARY, ACCENT, GREEN]

    for idx, (number, label) in enumerate(kpis):
        x = start_x + idx * (card_w + gap)
        add_kpi_card(slide, x, start_y, card_w, card_h, number, label, colors[idx])


def main():
    if len(sys.argv) < 2:
        print("用法: python content_layout.py <pptx> [outline.md]")
        sys.exit(1)

    pptx_path = sys.argv[1]
    outline_path = sys.argv[2] if len(sys.argv) > 2 else None

    prs = Presentation(pptx_path)

    outline = {}
    if outline_path and Path(outline_path).exists():
        outline = parse_outline(outline_path)

    print("增强内容页（2-8）...")
    enhance_content_pages(prs, outline)

    print("增强项目价值页（9）...")
    enhance_value_page(prs, outline)

    print("增强产出收益页（15）...")
    enhance_kpi_page(prs, outline)

    prs.save(pptx_path)
    print(f"✅ 内容布局完成: {pptx_path}")


if __name__ == "__main__":
    main()

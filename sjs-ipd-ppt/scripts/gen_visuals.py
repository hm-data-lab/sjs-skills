#!/usr/bin/env python3
"""
视觉内容生成器 — 基于 ppt-master SVG 模板

为每个内容页生成高质量视觉 PNG（白色背景）。
输出到指定目录，文件名: slide_{N}.png

用法：
    python gen_visuals.py --outline <outline.md> --output <visual_dir>
"""

import os
import sys
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape

import cairosvg

from utils import get_template_dir, fix_chinese_font
from outline_parser import parse_outline_simple
from chart_engine import infer_chart_type

TEMPLATE_DIR = get_template_dir()

# ══════════════════════════════════════════════════════
#  内容分析
# ══════════════════════════════════════════════════════

def load_svg(name):
    """加载 SVG 模板"""
    path = TEMPLATE_DIR / f"{name}.svg"
    if not path.exists():
        raise FileNotFoundError(f"模板不存在: {path}")
    return path.read_text(encoding="utf-8")


def svg_to_white_bg_png(svg_content, width=1280, height=720):
    """SVG → PNG（白色背景）"""
    # 替换背景色为白色（不添加新属性，替换已有的 fill）
    svg_content = svg_content.replace('fill="url(#bgGrad)"', 'fill="#FFFFFF"')
    svg_content = svg_content.replace('fill="#F8FAFC"', 'fill="#FFFFFF"')
    svg_content = svg_content.replace('fill="#F1F5F9"', 'fill="#FFFFFF"')
    svg_content = svg_content.replace('fill="#F9FAFB"', 'fill="#FFFFFF"')

    png_data = cairosvg.svg2png(
        bytestring=svg_content.encode("utf-8"),
        output_width=width,
        output_height=height,
    )
    return png_data


# ══════════════════════════════════════════════════════
#  图表生成
# ══════════════════════════════════════════════════════

def gen_pyramid_svg(items, title=""):
    """生成金字塔 SVG"""
    svg = load_svg("pyramid_chart")
    items = list(reversed(items[:5]))  # 反转：用户第1项=顶层

    if title:
        svg = svg.replace("Strategic Capability Pyramid", xml_escape(title))
    svg = svg.replace("FIVE-TIER EVOLUTION MODEL · FROM FOUNDATION TO VISION", "HIERARCHY · FROM FOUNDATION TO VISION")

    # 替换层级文字
    labels = ["VISION", "INNOVATION", "BRAND POWER", "R&amp;D SYSTEM", "OPERATIONS"]
    for i, item in enumerate(items):
        if i < len(labels):
            svg = svg.replace(labels[i], xml_escape(item[:18]))

    # 替换右侧卡片
    card_titles = ["Vision &amp; Purpose", "Innovation Engine", "Brand Power", "R&amp;D System", "Operations Foundation"]
    for i, item in enumerate(items):
        if i < len(card_titles):
            svg = svg.replace(card_titles[i], xml_escape(item[:30]))

    return fix_chinese_font(svg)


def gen_flow_svg(items, title=""):
    """生成流程图 SVG"""
    svg = load_svg("process_flow")
    items = items[:6]

    if title:
        svg = svg.replace("Process Flow", xml_escape(title))

    # 查找并替换 step 标签
    step_patterns = re.findall(r'<text[^>]*>([^<]*(?:Step|STEP|Phase|PHASE)[^<]*)</text>', svg)
    for i, pattern in enumerate(step_patterns):
        if i < len(items):
            svg = svg.replace(pattern, xml_escape(items[i][:20]), 1)

    return fix_chinese_font(svg)


def gen_comparison_svg(items, title=""):
    """生成对比图 SVG"""
    svg = load_svg("comparison_columns")
    if title:
        svg = svg.replace("Comparison", xml_escape(title))
    return fix_chinese_font(svg)


def gen_timeline_svg(items, title=""):
    """生成时间轴 SVG"""
    svg = load_svg("timeline")
    if title:
        svg = svg.replace("Timeline", xml_escape(title))
    return fix_chinese_font(svg)


def gen_kpi_svg(kpis, title=""):
    """生成 KPI 卡片 SVG"""
    svg = load_svg("kpi_cards")
    if title:
        svg = svg.replace("Key Performance Indicators", xml_escape(title))
    return fix_chinese_font(svg)


# ══════════════════════════════════════════════════════
#  主流程
# ══════════════════════════════════════════════════════

# 页面映射：大纲章节 → slide 编号
# Charter 模板的页面结构
SLIDE_MAP = {
    "一、项目背景": [2, 3, 4, 5],      # 4 个子页面
    "二、研究内容": [6, 7, 8],          # 3 个子页面
    "三、项目价值": [9],
    "四、项目目标": [10],
    "五、关键资源计划": [13],
    "六、财务计划": [14],
    "七、产出和收益": [15],
    "八、风险管理": [16],
    "九、结论": [17, 18],
}


def generate_visuals(outline_path, output_dir):
    """为每个内容页生成视觉 PNG"""
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    sections = parse_outline_simple(outline_path)
    generated = 0

    for heading, subs in sections.items():
        if heading.startswith("备注"):
            continue
        if heading not in SLIDE_MAP:
            continue

        slide_indices = SLIDE_MAP[heading]

        for sub_idx, sub in enumerate(subs):
            if sub_idx >= len(slide_indices):
                break

            slide_num = slide_indices[sub_idx]
            subtitle = sub.get("subtitle", heading)
            points = sub.get("points", [])

            if not points:
                continue

            chart_type = infer_chart_type(points, subtitle or "")

            # 只为需要视觉增强的类型生成 SVG
            if chart_type not in ("pyramid", "flow", "comparison", "timeline", "composition"):
                print(f"  slide_{slide_num}: {(subtitle or heading or '')[:20]} → {chart_type} (跳过，用 shapes)")
                continue

            try:
                title = subtitle or heading or ""
                if chart_type == "pyramid":
                    svg = gen_pyramid_svg(points, title=title)
                elif chart_type == "flow":
                    svg = gen_flow_svg(points, title=title)
                elif chart_type == "comparison":
                    svg = gen_comparison_svg(points, title=title)
                elif chart_type == "timeline":
                    svg = gen_timeline_svg(points, title=title)
                elif chart_type == "composition":
                    svg = gen_kpi_svg(points, title=title)
                else:
                    continue

                png = svg_to_white_bg_png(svg)
                out_file = output_path / f"slide_{slide_num}.png"
                out_file.write_bytes(png)
                print(f"  slide_{slide_num}: {title[:20]} → {chart_type} (SVG ✓)")
                generated += 1

            except Exception as e:
                print(f"  slide_{slide_num}: {title[:20]} → 失败: {e}")

    print(f"\n共生成 {generated} 个视觉文件到 {output_dir}/")
    return generated


def main():
    import argparse
    parser = argparse.ArgumentParser(description="基于 ppt-master SVG 模板生成视觉内容")
    parser.add_argument("--outline", required=True, help="大纲 Markdown 文件")
    parser.add_argument("--output", required=True, help="输出目录（存放 slide_N.png）")
    args = parser.parse_args()

    if not os.path.exists(args.outline):
        print(f"错误: 大纲文件不存在: {args.outline}")
        sys.exit(1)

    generate_visuals(args.outline, args.output)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
公共工具模块单元测试

测试 utils.py 中的工具函数。
"""

import os
import sys
import tempfile
import json
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

from utils import (
    strip_markdown_bold,
    fix_chinese_font,
    load_config,
    get_stage_config,
    STAGE_SLIDE_CONFIG,
)


def test_strip_markdown_bold():
    """测试 Markdown 粗体移除"""
    assert strip_markdown_bold("**粗体文本**") == "粗体文本"
    assert strip_markdown_bold("普通文本") == "普通文本"
    assert strip_markdown_bold("**A**和**B**") == "A和B"
    assert strip_markdown_bold("") == ""


def test_fix_chinese_font():
    """测试 SVG 字体修复"""
    svg = '<text font-family="Arial">Hello</text>'
    result = fix_chinese_font(svg)
    assert "Microsoft YaHei" in result
    assert "Arial" not in result


def test_fix_chinese_font_preserves_structure():
    """测试字体修复保留 SVG 结构"""
    svg = '<svg><text font-family="SimSun">中文</text><rect fill="red"/></svg>'
    result = fix_chinese_font(svg)
    assert "<svg>" in result
    assert "</svg>" in result
    assert 'fill="red"' in result


def test_get_stage_config():
    """测试阶段配置获取"""
    charter_cfg = get_stage_config("charter")
    assert charter_cfg["cover_slide"] == 0
    assert charter_cfg["summary_slide"] == 18
    assert charter_cfg["value_page"] == 9

    pdcp_cfg = get_stage_config("pdcp")
    assert pdcp_cfg["cover_slide"] == 0
    assert pdcp_cfg["summary_slide"] == 15

    # 未知阶段回退到 charter
    unknown_cfg = get_stage_config("unknown")
    assert unknown_cfg == STAGE_SLIDE_CONFIG["charter"]


def test_all_stages_have_required_keys():
    """测试所有阶段配置包含必需的键"""
    required_keys = {"fixed_slides", "cover_slide", "summary_slide",
                     "scoring_slide", "value_page", "output_page", "content_range"}
    for stage, cfg in STAGE_SLIDE_CONFIG.items():
        for key in required_keys:
            assert key in cfg, f"阶段 {stage} 缺少键 {key}"


if __name__ == "__main__":
    test_strip_markdown_bold()
    test_fix_chinese_font()
    test_fix_chinese_font_preserves_structure()
    test_get_stage_config()
    test_all_stages_have_required_keys()
    print("✅ 所有测试通过！")

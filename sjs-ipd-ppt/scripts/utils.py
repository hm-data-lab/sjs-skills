#!/usr/bin/env python3
"""
公共工具模块 — 配置加载、字体修复、路径解析

所有脚本共享的基础设施，避免重复实现。
"""

import json
import re
from pathlib import Path

# ── 路径常量 ──

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
ASSETS_DIR = SKILL_DIR / "assets"
CONFIG_PATH = SKILL_DIR / "config.json"


# ── 配置加载 ──

def load_config() -> dict:
    """加载 config.json，不存在则返回空字典"""
    if CONFIG_PATH.exists():
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def get_template_dir() -> Path:
    """获取 ppt-master SVG 模板目录，优先从 config 读取"""
    config = load_config()
    ppt_master_dir = config.get("ppt_master_dir")
    if ppt_master_dir:
        return Path(ppt_master_dir) / "skills" / "ppt-master" / "templates" / "charts"
    # 回退到默认路径
    return Path.home() / "work_space" / "ppt-master" / "skills" / "ppt-master" / "templates" / "charts"


# ── 字体修复 ──

_CHINESE_FONT = "'Microsoft YaHei', 'PingFang SC', 'Noto Sans CJK SC', sans-serif"


def fix_chinese_font(svg: str) -> str:
    """修复 SVG 中文字体 — 只替换 font-family 属性中的字体"""
    return re.sub(
        r'font-family="[^"]*"',
        f'font-family="{_CHINESE_FONT}"',
        svg,
    )


# ── Markdown 工具 ──

def strip_markdown_bold(text: str) -> str:
    """移除 Markdown 粗体标记 **text** → text"""
    return re.sub(r"\*\*(.+?)\*\*", r"\1", text)


# ── 按阶段配置 slide 映射 ──

STAGE_SLIDE_CONFIG = {
    "charter": {
        "fixed_slides": {1, 19, 20},
        "cover_slide": 0,
        "summary_slide": 18,
        "scoring_slide": 17,
        "value_page": 9,
        "output_page": 15,
        "content_range": (2, 9),
    },
    "pdcp": {
        "fixed_slides": {1, 16, 17},
        "cover_slide": 0,
        "summary_slide": 15,
        "scoring_slide": 14,
        "value_page": None,
        "output_page": None,
        "content_range": (2, 7),
    },
    "adcp": {
        "fixed_slides": {1, 14, 15},
        "cover_slide": 0,
        "summary_slide": 13,
        "scoring_slide": None,
        "value_page": None,
        "output_page": None,
        "content_range": (2, 5),
    },
    "transfer": {
        "fixed_slides": {1, 12, 13},
        "cover_slide": 0,
        "summary_slide": 11,
        "scoring_slide": None,
        "value_page": None,
        "output_page": None,
        "content_range": (2, 5),
    },
}


def get_stage_config(stage: str) -> dict:
    """获取阶段配置，未知阶段回退到 charter"""
    return STAGE_SLIDE_CONFIG.get(stage, STAGE_SLIDE_CONFIG["charter"])

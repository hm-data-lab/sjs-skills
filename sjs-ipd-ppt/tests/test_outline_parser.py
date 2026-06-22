#!/usr/bin/env python3
"""
大纲解析器单元测试

测试 outline_parser.py 的解析逻辑，确保各种格式的大纲都能正确解析。
"""

import os
import sys
import tempfile
from pathlib import Path

# 将 scripts 目录加入路径
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

from outline_parser import parse_outline, parse_outline_simple


def _write_temp(content: str) -> str:
    """写入临时文件并返回路径"""
    f = tempfile.NamedTemporaryFile(mode="w", suffix=".md", delete=False, encoding="utf-8")
    f.write(content)
    f.close()
    return f.name


def _cleanup(path: str):
    os.unlink(path)


# ══════════════════════════════════════════════════════
#  基础解析测试
# ══════════════════════════════════════════════════════

def test_parse_title():
    """测试标题解析"""
    md = """# 冰箱凝霜识别 — Charter 汇报材料

> 产品线：ml | 类型：基础研究 | 版本：v1

## 一、项目背景
- 要点1
"""
    path = _write_temp(md)
    try:
        result = parse_outline(path)
        assert result["title"] == "冰箱凝霜识别 — Charter 汇报材料"
        assert "ml" in result["meta"]
    finally:
        _cleanup(path)


def test_parse_meta():
    """测试元信息解析"""
    md = """# 测试项目

> 产品线：hm | 类型：应用研究 | 版本：v2

## 一、项目背景
- 要点
"""
    path = _write_temp(md)
    try:
        result = parse_outline(path)
        assert "hm" in result["meta"]
        assert "应用研究" in result["meta"]
    finally:
        _cleanup(path)


def test_parse_sections():
    """测试章节解析"""
    md = """# 测试

> 产品线：ml

## 一、项目背景
### 1.1 需求分析
- 需求1
- 需求2

### 1.2 竞品分析
- 竞品A
- 竞品B

## 二、研究内容
### 2.1 技术路线
- 阶段一
- 阶段二
"""
    path = _write_temp(md)
    try:
        result = parse_outline(path)
        assert len(result["sections"]) == 2
        assert result["sections"][0]["heading"] == "一、项目背景"
        assert result["sections"][1]["heading"] == "二、研究内容"
        assert len(result["sections"][0]["slides"]) == 2
        assert result["sections"][0]["slides"][0]["subtitle"] == "1.1 需求分析"
        assert result["sections"][0]["slides"][0]["points"] == ["需求1", "需求2"]
    finally:
        _cleanup(path)


def test_parse_table():
    """测试表格解析"""
    md = """# 测试

> 产品线：ml

## 一、Q指标
| 序号 | 指标要求 | 交付物 |
|------|---------|--------|
| 1 | 准确率≥90% | 算法模型 |
| 2 | 延迟<100ms | 部署文档 |
"""
    path = _write_temp(md)
    try:
        result = parse_outline(path)
        slides = result["sections"][0]["slides"]
        assert len(slides) == 1
        points = slides[0]["points"]
        # 表格包含表头行 + 数据行
        assert len(points) == 3
        assert "序号" in points[0]  # 表头
        assert "准确率≥90%" in points[1]
        assert "算法模型" in points[1]
    finally:
        _cleanup(path)


def test_parse_notes():
    """测试备注区解析"""
    md = """# 测试

> 产品线：ml

## 一、项目背景
### 1.1 需求
- 要点

## 备注区
## 一、项目背景
- 讲解要点：行业趋势
- 数据支撑：市场规模

## 二、研究内容
- 技术路线说明
"""
    path = _write_temp(md)
    try:
        result = parse_outline(path)
        assert "一、项目背景" in result["notes_map"]
        assert "行业趋势" in result["notes_map"]["一、项目背景"]
        assert "二、研究内容" in result["notes_map"]
        assert "技术路线说明" in result["notes_map"]["二、研究内容"]
    finally:
        _cleanup(path)


def test_parse_chart_hint():
    """测试图表标注解析"""
    md = """# 测试

> 产品线：ml

## 一、项目背景
<!-- chart: comparison -->
### 1.1 竞品对比
| 维度 | 我方 | 竞品 |
|------|------|------|
| 精度 | 95% | 80% |
"""
    path = _write_temp(md)
    try:
        result = parse_outline(path)
        slide = result["sections"][0]["slides"][0]
        assert slide["chart"] == "comparison"
    finally:
        _cleanup(path)


def test_parse_plain_text():
    """测试普通文本行解析 — 在列表项之后的文本会被追加"""
    md = """# 测试

> 产品线：ml

## 一、项目背景
### 1.1 说明
- 列表项1
这是一段补充文本
- 列表项2
"""
    path = _write_temp(md)
    try:
        result = parse_outline(path)
        points = result["sections"][0]["slides"][0]["points"]
        assert "列表项1" in points
        assert "这是一段补充文本" in points
        assert "列表项2" in points
    finally:
        _cleanup(path)


# ══════════════════════════════════════════════════════
#  简化版解析测试
# ══════════════════════════════════════════════════════

def test_parse_simple():
    """测试简化版解析"""
    md = """# 测试

> 产品线：ml

## 一、项目背景
### 1.1 需求
- 要点1

## 备注区
## 一、项目背景
- 讲解要点
"""
    path = _write_temp(md)
    try:
        result = parse_outline_simple(path)
        assert "一、项目背景" in result
        assert "备注区" not in result  # 备注区应被跳过
        assert len(result["一、项目背景"]) == 1
        assert result["一、项目背景"][0]["subtitle"] == "1.1 需求"
    finally:
        _cleanup(path)


# ══════════════════════════════════════════════════════
#  边界情况测试
# ══════════════════════════════════════════════════════

def test_empty_file():
    """测试空文件"""
    path = _write_temp("")
    try:
        result = parse_outline(path)
        assert result["title"] == ""
        assert result["sections"] == []
    finally:
        _cleanup(path)


def test_no_meta():
    """测试无元信息"""
    md = """# 测试项目

## 一、项目背景
- 要点
"""
    path = _write_temp(md)
    try:
        result = parse_outline(path)
        assert result["title"] == "测试项目"
        assert result["meta"] == ""
    finally:
        _cleanup(path)


def test_section_without_subtitle():
    """测试无子标题的章节"""
    md = """# 测试

> 产品线：ml

## 一、产出和收益
- 亮点1：年节省300万
- 亮点2：覆盖100万台
"""
    path = _write_temp(md)
    try:
        result = parse_outline(path)
        slides = result["sections"][0]["slides"]
        assert len(slides) == 1
        assert slides[0]["subtitle"] is None
        assert len(slides[0]["points"]) == 2
    finally:
        _cleanup(path)


if __name__ == "__main__":
    test_parse_title()
    test_parse_meta()
    test_parse_sections()
    test_parse_table()
    test_parse_notes()
    test_parse_chart_hint()
    test_parse_plain_text()
    test_parse_simple()
    test_empty_file()
    test_no_meta()
    test_section_without_subtitle()
    print("✅ 所有测试通过！")

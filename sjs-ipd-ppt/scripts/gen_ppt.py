#!/usr/bin/env python3
"""
IPD 项目评审 PPT 生成脚本

使用方法：
    python gen_ppt.py --stage charter --project "项目名" --template /path/to/template.pptx
    python gen_ppt.py --stage pdcp --project "项目名" --template /path/to/template.pptx --output /output/dir/

参数：
    --stage       阶段：charter / pdcp / adcp / transfer
    --project     项目名称
    --type        项目类型：basic（基础研究）/ applied（应用研究），默认 basic
    --template    PPT 模板文件路径（必需，除非 config.json 中已配置）
    --output      输出目录，默认当前目录
    --cost-per-day 人力成本基准（元/人天），默认从 config.json 读取，未配置则 970

配置文件：
    脚本会尝试读取 Skill 目录下的 config.json，如存在则使用其中的配置。
    命令行参数优先于 config.json。
"""

import argparse
import json
import os
import sys
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN


def load_config() -> dict:
    """尝试加载 config.json"""
    # 在脚本所在目录的上级目录查找 config.json（即 skill 根目录）
    script_dir = Path(__file__).resolve().parent
    skill_dir = script_dir.parent
    config_path = skill_dir / "config.json"

    if config_path.exists():
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def create_presentation(
    project_name: str,
    stage: str,
    project_type: str,
    template_path: str,
    cost_per_day: int = 970,
) -> Presentation:
    """创建评审 PPT"""
    if not os.path.exists(template_path):
        print(f"错误：模板文件不存在: {template_path}", file=sys.stderr)
        sys.exit(1)

    prs = Presentation(template_path)

    add_cover_slide(prs, project_name, stage)
    add_toc_slide(prs, stage)

    if stage == "charter":
        add_charter_slides(prs, project_name, project_type, cost_per_day)
    elif stage == "pdcp":
        add_pdcp_slides(prs, project_name, project_type, cost_per_day)
    elif stage == "adcp":
        add_adcp_slides(prs, project_name, project_type, cost_per_day)
    elif stage == "transfer":
        add_transfer_slides(prs, project_name, project_type, cost_per_day)

    add_ending_slide(prs)
    return prs


def add_cover_slide(prs: Presentation, project_name: str, stage: str) -> None:
    """添加封面"""
    stage_names = {
        "charter": "立项（Charter）汇报材料",
        "pdcp": "总体设计方案评审（PDCP）汇报材料",
        "adcp": "可获得性评审（ADCP）汇报材料",
        "transfer": "转移阶段汇报材料",
    }
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.placeholders[0].text = project_name
    slide.placeholders[1].text = stage_names.get(stage, "汇报材料")


def add_toc_slide(prs: Presentation, stage: str) -> None:
    """添加目录页"""
    toc_items = {
        "charter": [
            "一、项目背景",
            "二、研究内容",
            "三、项目价值",
            "四、项目目标（Q/C/D）",
            "五、关键资源计划",
            "六、风险及问题管理",
            "七、财务计划",
            "八、结论",
        ],
        "pdcp": [
            "一、项目简介",
            "二、项目技术目标",
            "三、项目研究方法",
            "四、技术方案及关键点",
            "五、关键资源计划",
            "六、项目风险管理",
            "七、知识产权方案",
            "八、环境及职业健康影响",
            "九、项目成员构成",
        ],
        "adcp": [
            "一、项目简介",
            "二、项目目标完成情况",
            "三、项目转移计划",
            "四、关键资源计划",
            "五、项目转移风险管理",
            "六、项目技术创新总结",
            "七、经济效益和应用前景",
            "八、项目过程运作",
        ],
        "transfer": [
            "一、项目简介",
            "二、项目目标转移情况",
            "三、项目技术成果",
            "四、技术创新点推广应用",
            "五、转移过程运作",
        ],
    }
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.placeholders[0].text = "目录"
    body = slide.placeholders[1]
    for i, item in enumerate(toc_items.get(stage, [])):
        if i == 0:
            body.text_frame.paragraphs[0].text = item
        else:
            p = body.text_frame.add_paragraph()
            p.text = item


def add_content_slide(
    prs: Presentation, title: str, points: list[str], notes: str = ""
) -> None:
    """添加内容页"""
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.placeholders[0].text = title
    body = slide.placeholders[1]
    for i, point in enumerate(points):
        if i == 0:
            body.text_frame.paragraphs[0].text = point
        else:
            p = body.text_frame.add_paragraph()
            p.text = point
            p.level = 0
    if notes:
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = notes


def add_charter_slides(
    prs: Presentation, project_name: str, project_type: str, cost_per_day: int
) -> None:
    """添加 Charter 阶段内容页"""
    add_content_slide(
        prs,
        "一、项目背景",
        [
            "1. 行业趋势：[待填写]",
            "2. 技术现状：[待填写]",
            "3. 痛点问题：[待填写]",
            "4. 需求来源：[待填写]",
        ],
        "讲解要点：说明为什么是现在做这件事，用数据支撑行业趋势",
    )

    add_content_slide(
        prs,
        "二、研究内容",
        [
            "1. 核心研究方向：[待填写]",
            "2. 技术路线概述：[待填写]",
            "3. 关键方法论：[待填写]",
        ],
        "讲解要点：概述技术路线，不需要展开到实现细节",
    )

    add_content_slide(
        prs,
        "三、项目价值",
        [
            "技术价值：[待填写]",
            "  - 突破了什么技术瓶颈",
            "  - 建立了什么能力",
            "经济价值：[待填写]",
            "  - 预期经济效益",
            "  - 投入产出比",
            "应用价值：[待填写]",
            "  - 可应用的产品/场景",
            "  - 市场前景",
        ],
        "讲解要点：这是最重要的部分！用数据和案例支撑价值论证",
    )

    add_content_slide(
        prs,
        "四、项目目标（Q/C/D）",
        [
            "Q 指标（技术规格）：[待填写]",
            "C 指标（财务）：总预算 [待填写] 万元",
            "D 指标（进度）：[待填写] 立项 → [待填写] 验收",
        ],
        "讲解要点：明确、可量化的目标",
    )

    add_content_slide(
        prs,
        "五、关键资源计划",
        [
            f"人力：[待填写] 人天 × {cost_per_day} = [自动计算] 万元",
            "设备：[待填写]",
            "外部资源：[待填写]",
        ],
    )

    add_content_slide(
        prs,
        "六、风险及问题管理",
        [
            "风险1：[待填写]（高/中/低）",
            "  应对措施：[待填写]",
            "风险2：[待填写]（高/中/低）",
            "  应对措施：[待填写]",
        ],
    )

    add_content_slide(
        prs,
        "七、财务计划",
        [
            "人力成本：[待填写] 万元",
            "服务器：[待填写] 万元",
            "试制费：[待填写] 万元",
            "其他：[待填写] 万元",
            "合计：[自动计算] 万元",
        ],
    )

    add_content_slide(
        prs,
        "八、结论",
        [
            "1. 项目必要性：[一句话总结]",
            "2. 预期成果：[一句话总结]",
            "3. 下一步行动：[一句话总结]",
        ],
        "讲解要点：用 2-3 句话有力地总结，给评审委员留下深刻印象",
    )


def add_pdcp_slides(
    prs: Presentation, project_name: str, project_type: str, cost_per_day: int
) -> None:
    """添加 PDCP 阶段内容页"""
    slides_content = [
        ("一、项目简介", ["项目背景：[待填写]", "项目目标：[待填写]", "项目范围：[待填写]"]),
        ("二、项目技术目标", ["Q 指标细化：[待填写]", "C 指标各阶段分解：[待填写]", "D 指标里程碑：[待填写]"]),
        ("三、项目研究方法", ["技术路线图：[待填写]", "分步实施计划：[待填写]", "各阶段产出：[待填写]"]),
        ("四、技术方案及关键点", ["系统架构：[待填写]", "核心算法/模型：[待填写]", "关键技术难点：[待填写]", "可行性分析：[待填写]"]),
        ("五、关键资源计划", [f"人力：[待填写] 人天 × {cost_per_day}", "设备：[待填写]", "关键技术可获得性：[待填写]"]),
        ("六、项目风险管理", ["风险1：[待填写]", "风险2：[待填写]"]),
        ("七、知识产权方案", ["专利检索：[待填写]", "专利规划：[待填写]"]),
        ("八、环境及职业健康影响", ["影响分析：[待填写]"]),
        ("九、项目成员构成", ["项目经理：[待填写]", "核心成员：[待填写]"]),
    ]
    for title, points in slides_content:
        add_content_slide(prs, title, points)


def add_adcp_slides(
    prs: Presentation, project_name: str, project_type: str, cost_per_day: int
) -> None:
    """添加 ADCP 阶段内容页"""
    slides_content = [
        ("一、项目简介", ["项目概述：[待填写]"]),
        ("二、项目目标完成情况", [
            "Q 指标偏差分析：",
            "  指标1：目标 [待填写] → 实际 [待填写]",
            "C 指标偏差分析：",
            "  计划 [待填写] 万 → 实际 [待填写] 万",
            "D 指标偏差分析：",
            "  计划 [待填写] → 实际 [待填写]",
        ]),
        ("三、项目转移计划", ["应用产品/平台：[待填写]", "转移条件：[待填写]", "完成时间：[待填写]"]),
        ("四、关键资源计划", ["转移阶段资源：[待填写]"]),
        ("五、项目转移风险管理", ["风险：[待填写]", "应对：[待填写]"]),
        ("六、项目技术创新总结", ["创新点1：[待填写]", "创新点2：[待填写]", "专利/论文：[待填写]"]),
        ("七、经济效益和应用前景", ["已实现效益：[待填写]", "市场规模：[待填写]", "推广策略：[待填写]"]),
        ("八、项目过程运作", ["亮点：[待填写]", "不足：[待填写]", "经验教训：[待填写]"]),
    ]
    for title, points in slides_content:
        add_content_slide(prs, title, points)


def add_transfer_slides(
    prs: Presentation, project_name: str, project_type: str, cost_per_day: int
) -> None:
    """添加转移阶段内容页"""
    slides_content = [
        ("一、项目简介", ["项目基本信息：[待填写]"]),
        ("二、项目目标转移情况", [
            "Q 指标转移：",
            "  正式样机测试数据：[待填写]",
            "  正式样机评审：[待填写]",
            "  小批试制：[待填写]",
            "  试生产评审：[待填写]",
            "C 指标转移：[待填写]",
            "D 指标转移：[待填写]",
        ]),
        ("三、项目技术成果", ["方法/规范/标准：[待填写]", "专利：[待填写]", "论文：[待填写]"]),
        ("四、技术创新点推广应用", ["推广产品/平台：[待填写]", "推广效果：[待填写]", "后续计划：[待填写]"]),
        ("五、转移过程运作", ["亮点：[待填写]", "不足：[待填写]", "经验教训：[待填写]"]),
    ]
    for title, points in slides_content:
        add_content_slide(prs, title, points)


def add_ending_slide(prs: Presentation) -> None:
    """添加结尾页"""
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.placeholders[0].text = "谢谢！"
    slide.placeholders[1].text = "请各位领导和专家指导"


def main() -> None:
    parser = argparse.ArgumentParser(description="IPD 项目评审 PPT 生成")
    parser.add_argument("--stage", required=True, choices=["charter", "pdcp", "adcp", "transfer"])
    parser.add_argument("--project", required=True, help="项目名称")
    parser.add_argument("--type", default="basic", choices=["basic", "applied"])
    parser.add_argument("--template", default=None, help="PPT 模板文件路径")
    parser.add_argument("--output", default=".", help="输出目录")
    parser.add_argument("--cost-per-day", type=int, default=None, help="人力成本基准（元/人天）")
    args = parser.parse_args()

    # 加载配置
    config = load_config()

    # 命令行参数优先于 config.json
    template_path = args.template or config.get("template_path")
    if not template_path:
        print("错误：未指定模板路径。请通过 --template 参数指定，或在 config.json 中配置 template_path", file=sys.stderr)
        sys.exit(1)

    cost_per_day = args.cost_per_day or config.get("cost_per_day", 970)

    prs = create_presentation(args.project, args.stage, args.type, template_path, cost_per_day)

    # 确保输出目录存在
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    filename = f"{args.project}-{args.stage.upper()}-汇报材料.pptx"
    output_path = output_dir / filename
    prs.save(str(output_path))
    print(f"PPT 已生成: {output_path}")


if __name__ == "__main__":
    main()

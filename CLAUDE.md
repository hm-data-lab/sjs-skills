# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Overview

虹美数据技术所（hm-data-lab）AI 辅助工具 Skill 集合。每个 Skill 通过 `SKILL.md` 定义，可安装到多种 AI 编码工具（Claude Code、Copilot CLI、Codex、OpenCode、MimoCode、Qoder）。

**Naming convention:** 所有 Skill 目录必须以 `sjs-` 为前缀。

## Skill 结构规范

```
sjs-{name}/
├── SKILL.md              # 必需：YAML frontmatter + 指令内容
├── README.md             # 可选：说明文档
├── references/           # 可选：参考文档（引导流程、评审指南等）
├── scripts/              # 可选：辅助脚本（如 Python PPT 生成器）
└── assets/               # 可选：资源文件（模板、图片等）
```

### SKILL.md 格式

```yaml
---
name: sjs-{name}
description: 简要描述功能和触发条件（含口语化触发词）
---
# Skill 标题
... 指令内容 ...
```

`description` 字段同时作为触发条件，应包含用户可能使用的口语化表达。

## Adding a New Skill

1. 创建 `sjs-{功能名}/` 目录
2. 编写 `SKILL.md`（必须含 YAML frontmatter）
3. 更新根目录 `README.md` 的 Skills 列表
4. 运行 `bash install.sh sjs-{name}` 测试安装

## Install Script

`install.sh` 检测已安装的 AI 工具目录（`~/.claude/skills/`、`~/.copilot/skills/` 等），将 Skill 拷贝到对应工具路径。支持：
- `bash install.sh` — 安装全部 Skills
- `bash install.sh sjs-ipd-ppt` — 安装单个 Skill

## Existing Skills

### sjs-ipd-ppt

IPD 项目评审 PPT 生成（Charter / PDCP / ADCP / 转移）。

- **生成器:** [scripts/gen_ppt.py](sjs-ipd-ppt/scripts/gen_ppt.py) — 基于 `python-pptx` 操作 PPT 模板
- **模板:** [assets/template.pptx](sjs-ipd-ppt/assets/template.pptx)
- **首次使用:** 需配置 `config.json`（模板路径、项目目录、人力成本、产品线），参见 [references/config-guide.md](sjs-ipd-ppt/references/config-guide.md)
- **评审阶段指南:** [references/charter-guide.md](sjs-ipd-ppt/references/charter-guide.md)、[pdcp-guide.md](sjs-ipd-ppt/references/pdcp-guide.md)、[adcp-guide.md](sjs-ipd-ppt/references/adcp-guide.md)、[transfer-guide.md](sjs-ipd-ppt/references/transfer-guide.md)

## Key Conventions

- Skill 指令用中文编写，代码注释可用中英文
- 配置驱动：避免硬编码路径，通过 `config.json` 管理
- `references/` 下的指南文档是 Skill 执行时的参考依据，修改时需与 `SKILL.md` 中的流程保持一致

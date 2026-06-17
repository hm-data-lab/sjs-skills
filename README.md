# sjs-ipd-ppt

虹美数据技术所 IPD 项目评审 PPT 生成 Skill。

基于公司 IPD 流程，为 AI 辅助生成立项（Charter）、PDCP、ADCP、转移四个阶段的评审汇报材料。

## 功能

- **Charter（立项）**：引导式提炼项目必要性、重要性、投入产出价值
- **PDCP（总体设计方案评审）**：整理技术方案、研究方法、架构设计
- **ADCP（可获得性评审）**：Q/C/D 指标偏差分析、费用对比、成果总结
- **转移阶段**：技术成果转移、推广应用、经验总结

### 核心特点

- **引导式为主**：大部分人知道自己做什么，但讲不出价值。Skill 通过逐层提问帮助提炼
- **价值框架**：内置技术价值/经济价值/应用价值/战略价值提炼模板
- **费用自动计算**：人力成本基准 970 元/人天
- **双类型支持**：基础研究（技术驱动）和应用研究（需求驱动）

## 安装

### 一键安装（推荐）

```bash
# 克隆仓库
git clone https://github.com/hm-data-lab/sjs-ipd-ppt.git
cd sjs-ipd-ppt

# 自动安装到已安装的 AI 工具
bash install.sh
```

### 手动安装

将 `SKILL.md` 拷贝到对应工具的 skills 目录：

| 工具 | 安装路径 |
|------|---------|
| Claude Code | `~/.claude/skills/sjs-ipd-ppt/SKILL.md` |
| Copilot CLI | `~/.copilot/skills/sjs-ipd-ppt/SKILL.md` |
| Codex | `~/.codex/skills/sjs-ipd-ppt/SKILL.md` |
| OpenCode | `~/.opencode/skills/sjs-ipd-ppt/SKILL.md` |
| MimoCode | `~/.mimocode/skills/sjs-ipd-ppt/SKILL.md` |
| Qoder | `~/.qoder/skills/sjs-ipd-ppt/SKILL.md` |
| Cursor | 项目根目录 `.cursorrules` 中引用 SKILL.md 内容 |

## 使用示例

### 立项（Charter）

```
用户：我要给冰箱凝霜项目立项
Skill：请提供以下信息：
1. 项目全称？
2. 空调还是美菱产品线？
3. 基础研究还是应用研究？

用户：冰箱凝霜程度识别技术研究，美菱，基础研究
Skill：好的，请简单描述一下这个项目是做什么的？
...
```

### PDCP

```
用户：帮我做凝霜项目的PDCP汇报材料
Skill：好的，请描述一下总体技术方案...
```

### ADCP

```
用户：凝霜项目要结项了，帮我准备ADCP材料
Skill：好的，我们逐项核对指标完成情况。立项时的Q指标有哪些？
...
```

## 目录结构

```
sjs-ipd-ppt/
├── SKILL.md              # Skill 主文件
├── README.md             # 本文件
├── install.sh            # 一键安装脚本
├── LICENSE               # MIT License
├── references/           # 阶段详细指南
│   ├── charter-guide.md
│   ├── pdcp-guide.md
│   ├── adcp-guide.md
│   └── transfer-guide.md
├── scripts/              # 辅助脚本
│   └── gen_ppt.py        # PPT 生成模板
└── assets/               # 资源文件
    └── template-info.md  # PPT 模板说明
```

## 贡献

欢迎数据所团队成员贡献和完善！

1. Fork 本仓库
2. 创建特性分支：`git checkout -b feature/your-feature`
3. 提交更改
4. 创建 Pull Request

### 可改进方向

- 补充更多引导式提问模板
- 增加行业特定的价值论证案例
- 完善 PPT 生成脚本
- 添加更多可视化图表模板

## 依赖

- Python 3.x（用于 PPT 生成）
- python-pptx（`pip install python-pptx`）

## License

MIT License

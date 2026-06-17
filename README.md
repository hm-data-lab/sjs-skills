# sjs-skills

虹美数据技术所（hm-data-lab）AI 辅助工具 Skill 集合。

## Skills 列表

| Skill | 说明 | 状态 |
|-------|------|------|
| [sjs-ipd-ppt](sjs-ipd-ppt/) | IPD 项目评审 PPT 生成（Charter/PDCP/ADCP/转移） | 可用 |

> 所有数据所 Skill 统一以 `sjs-` 为前缀。

## 安装

### 安装所有 Skills

```bash
git clone https://github.com/hm-data-lab/sjs-skills.git
cd sjs-skills
bash install.sh
```

### 安装单个 Skill

```bash
bash install.sh sjs-ipd-ppt
```

### 支持的 AI 工具

| 工具 | 安装路径 |
|------|---------|
| Claude Code | `~/.claude/skills/{skill-name}/SKILL.md` |
| Copilot CLI | `~/.copilot/skills/{skill-name}/SKILL.md` |
| Codex | `~/.codex/skills/{skill-name}/SKILL.md` |
| OpenCode | `~/.opencode/skills/{skill-name}/SKILL.md` |
| MimoCode | `~/.mimocode/skills/{skill-name}/SKILL.md` |
| Qoder | `~/.qoder/skills/{skill-name}/SKILL.md` |

## 添加新 Skill

1. 在仓库根目录创建新目录，命名规范：`sjs-{功能名}`
2. 目录内必须包含 `SKILL.md`（YAML frontmatter + 指令内容）
3. 可选：`references/`、`scripts/`、`assets/`
4. 更新本 README 的 Skills 列表
5. 提交 PR

### Skill 目录结构

```
sjs-{name}/
├── SKILL.md              # 必需：Skill 主文件
├── README.md             # 可选：Skill 说明
├── references/           # 可选：参考文档
├── scripts/              # 可选：辅助脚本
└── assets/               # 可选：资源文件
```

### SKILL.md 规范

```yaml
---
name: sjs-{name}
description: 简要描述功能和触发条件
---
```

## 贡献

1. Fork 本仓库
2. 创建特性分支：`git checkout -b sjs-your-skill`
3. 添加你的 Skill 目录
4. 更新 README
5. 创建 Pull Request

## License

MIT License

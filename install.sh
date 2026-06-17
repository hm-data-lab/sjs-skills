#!/bin/bash
# sjs-ipd-ppt 一键安装脚本
# 自动检测已安装的 AI 工具并安装 Skill

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
SKILL_FILE="$SCRIPT_DIR/SKILL.md"

if [ ! -f "$SKILL_FILE" ]; then
    echo "错误: 未找到 SKILL.md，请在 sjs-ipd-ppt 目录下运行此脚本"
    exit 1
fi

INSTALLED=0

# 定义工具及其 skills 路径
declare -A TOOLS=(
    ["Claude Code"]="$HOME/.claude/skills/sjs-ipd-ppt"
    ["Copilot CLI"]="$HOME/.copilot/skills/sjs-ipd-ppt"
    ["Codex"]="$HOME/.codex/skills/sjs-ipd-ppt"
    ["OpenCode"]="$HOME/.opencode/skills/sjs-ipd-ppt"
    ["MimoCode"]="$HOME/.mimocode/skills/sjs-ipd-ppt"
    ["Qoder"]="$HOME/.qoder/skills/sjs-ipd-ppt"
)

echo "=== sjs-ipd-ppt Skill 安装程序 ==="
echo ""

for tool_name in "${!TOOLS[@]}"; do
    target_dir="${TOOLS[$tool_name]}"
    parent_dir="$(dirname "$target_dir")"

    if [ -d "$parent_dir" ]; then
        mkdir -p "$target_dir"
        cp "$SKILL_FILE" "$target_dir/"
        echo "已安装到 $tool_name: $target_dir"
        INSTALLED=$((INSTALLED + 1))
    else
        echo "跳过 $tool_name (未检测到安装目录: $parent_dir)"
    fi
done

echo ""

# Cursor / Windsurf 支持
CURSOR_DIR="${TOOLS["Claude Code"]%/*/*}"
if [ -d "$HOME/.cursor" ] || [ -d "$HOME/.windsurf" ]; then
    echo "检测到 Cursor/Windsurf，请手动将 SKILL.md 内容添加到项目 .cursorrules 文件"
fi

if [ $INSTALLED -eq 0 ]; then
    echo "未检测到已安装的 AI 工具"
    echo "请手动拷贝 SKILL.md 到对应工具的 skills 目录"
    echo ""
    echo "手动安装方法："
    echo "  mkdir -p <工具skills路径>/sjs-ipd-ppt"
    echo "  cp SKILL.md <工具skills路径>/sjs-ipd-ppt/"
else
    echo ""
    echo "安装完成！共安装到 $INSTALLED 个工具"
    echo "重启对应的 AI 工具即可使用"
fi

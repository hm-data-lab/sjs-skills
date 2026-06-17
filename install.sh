#!/bin/bash
# sjs-skills 一键安装脚本
# 安装所有或指定的 skill 到已安装的 AI 工具

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TARGET_SKILL="${1:-}"  # 可选：指定安装某个 skill

# 定义工具及其 skills 路径
declare -A TOOL_PATHS=(
    ["Claude Code"]="$HOME/.claude/skills"
    ["Copilot CLI"]="$HOME/.copilot/skills"
    ["Codex"]="$HOME/.codex/skills"
    ["OpenCode"]="$HOME/.opencode/skills"
    ["MimoCode"]="$HOME/.mimocode/skills"
    ["Qoder"]="$HOME/.qoder/skills"
)

echo "=== sjs-skills 安装程序 ==="
echo ""

# 收集已安装的工具
AVAILABLE_TOOLS=()
for tool_name in "${!TOOL_PATHS[@]}"; do
    parent_dir="${TOOL_PATHS[$tool_name]}"
    if [ -d "$parent_dir" ]; then
        AVAILABLE_TOOLS+=("$tool_name")
    fi
done

if [ ${#AVAILABLE_TOOLS[@]} -eq 0 ]; then
    echo "未检测到已安装的 AI 工具"
    echo "支持的工具：Claude Code, Copilot CLI, Codex, OpenCode, MimoCode, Qoder"
    echo "请先安装至少一个 AI 工具，或手动拷贝 SKILL.md"
    exit 1
fi

echo "检测到工具：${AVAILABLE_TOOLS[*]}"
echo ""

# 确定要安装的 skills
if [ -n "$TARGET_SKILL" ]; then
    if [ ! -d "$SCRIPT_DIR/$TARGET_SKILL" ]; then
        echo "错误：未找到 skill '$TARGET_SKILL'"
        echo "可用的 skills："
        for dir in "$SCRIPT_DIR"/sjs-*/; do
            [ -d "$dir" ] && echo "  $(basename "$dir")"
        done
        exit 1
    fi
    SKILLS=("$TARGET_SKILL")
else
    SKILLS=()
    for dir in "$SCRIPT_DIR"/sjs-*/; do
        [ -d "$dir" ] && SKILLS+=("$(basename "$dir")")
    done
fi

if [ ${#SKILLS[@]} -eq 0 ]; then
    echo "未找到任何 skill（以 sjs- 开头的目录）"
    exit 1
fi

echo "准备安装 ${#SKILLS[@]} 个 skill：${SKILLS[*]}"
echo ""

# 安装
INSTALLED=0
for skill in "${SKILLS[@]}"; do
    skill_file="$SCRIPT_DIR/$skill/SKILL.md"
    if [ ! -f "$skill_file" ]; then
        echo "跳过 $skill（未找到 SKILL.md）"
        continue
    fi

    for tool_name in "${AVAILABLE_TOOLS[@]}"; do
        target_dir="${TOOL_PATHS[$tool_name]}/$skill"
        mkdir -p "$target_dir"

        # 拷贝整个 skill 目录
        cp -r "$SCRIPT_DIR/$skill/"* "$target_dir/" 2>/dev/null

        echo "已安装 $skill → $tool_name"
        INSTALLED=$((INSTALLED + 1))
    done
done

echo ""
echo "安装完成！共安装 $INSTALLED 个 skill"
echo "重启对应的 AI 工具即可使用"

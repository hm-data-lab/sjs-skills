#!/bin/bash
# 企业微信智能表格读取 Skill 安装脚本（跨平台）

set -e

echo "=== 安装企业微信智能表格读取 Skill ==="

# 检测操作系统
OS="$(uname -s)"
case "${OS}" in
    Linux*)     MACHINE=Linux;;
    Darwin*)    MACHINE=Mac;;
    CYGWIN*)    MACHINE=Cygwin;;
    MINGW*)     MACHINE=MinGw;;
    MSYS*)      MACHINE=MsYS;;
    *)          MACHINE="UNKNOWN:${OS}"
esac

echo "检测到操作系统: ${MACHINE}"

# 检查 Python 版本
echo "检查 Python 版本..."
python3 --version || python --version

# 安装依赖
echo "安装依赖..."
pip install playwright || pip3 install playwright

# 检测浏览器
echo "检测浏览器..."

# 定义浏览器路径（跨平台）
detect_browser() {
    if [[ "${MACHINE}" == "Mac" ]]; then
        # macOS
        if [ -d "/Applications/Google Chrome.app" ]; then
            echo "Chrome: /Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
        fi
        if [ -d "/Applications/Microsoft Edge.app" ]; then
            echo "Edge: /Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"
        fi
        if [ -d "/Applications/Firefox.app" ]; then
            echo "Firefox: /Applications/Firefox.app/Contents/MacOS/firefox"
        fi
    elif [[ "${MACHINE}" == "Linux" ]]; then
        # Linux
        if [ -f "/usr/bin/google-chrome" ] || [ -f "/usr/bin/google-chrome-stable" ]; then
            echo "Chrome: $(which google-chrome || which google-chrome-stable)"
        fi
        if [ -f "/usr/bin/microsoft-edge" ] || [ -f "/usr/bin/microsoft-edge-stable" ]; then
            echo "Edge: $(which microsoft-edge || which microsoft-edge-stable)"
        fi
        if [ -f "/usr/bin/firefox" ]; then
            echo "Firefox: $(which firefox)"
        fi
    elif [[ "${MACHINE}" == "Cygwin" ]] || [[ "${MACHINE}" == "MinGw" ]] || [[ "${MACHINE}" == "MsYS" ]]; then
        # Windows (Git Bash / MSYS2)
        if [ -f "$PROGRAMFILES/Google/Chrome/Application/chrome.exe" ]; then
            echo "Chrome: $PROGRAMFILES/Google/Chrome/Application/chrome.exe"
        fi
        if [ -f "$PROGRAMFILES/Microsoft/Edge/Application/msedge.exe" ]; then
            echo "Edge: $PROGRAMFILES/Microsoft/Edge/Application/msedge.exe"
        fi
        if [ -f "$PROGRAMFILES/Mozilla Firefox/firefox.exe" ]; then
            echo "Firefox: $PROGRAMFILES/Mozilla Firefox/firefox.exe"
        fi
    fi
}

echo "检测到的浏览器："
detect_browser || echo "未检测到浏览器"

# 创建输出目录
echo "创建输出目录..."
mkdir -p output

# 设置脚本权限
echo "设置脚本权限..."
chmod +x scripts/*.py

echo ""
echo "=== 安装完成 ==="
echo ""
echo "使用方法："
echo ""
echo "1. 运行跨平台读取工具："
echo "   cd scripts"
echo "   python3 wecom_cross_platform.py"
echo ""
echo "2. 或者手动启动浏览器："
echo "   # macOS/Linux:"
echo '   "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge" --remote-debugging-port=9222 &'
echo ""
echo "   # Windows (CMD):"
echo '   "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222'
echo ""
echo "   # Windows (PowerShell):"
echo '   & "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222'
echo ""
echo "3. 在浏览器中登录企业微信"
echo ""
echo "4. 运行读取脚本："
echo "   python3 wecom_cross_platform.py"
echo ""

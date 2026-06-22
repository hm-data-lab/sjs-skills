@echo off
REM 企业微信智能表格读取 Skill 安装脚本（Windows）

echo === 安装企业微信智能表格读取 Skill ===

REM 检查 Python 版本
echo 检查 Python 版本...
python --version
if errorlevel 1 (
    echo Python 未安装或不在 PATH 中
    echo 请先安装 Python: https://www.python.org/downloads/
    pause
    exit /b 1
)

REM 安装依赖
echo 安装依赖...
pip install playwright
if errorlevel 1 (
    pip3 install playwright
)

REM 检测浏览器
echo 检测浏览器...

if exist "%PROGRAMFILES%\Google\Chrome\Application\chrome.exe" (
    echo Chrome: %PROGRAMFILES%\Google\Chrome\Application\chrome.exe
)
if exist "%PROGRAMFILES(x86)%\Google\Chrome\Application\chrome.exe" (
    echo Chrome: %PROGRAMFILES(x86)%\Google\Chrome\Application\chrome.exe
)
if exist "%PROGRAMFILES%\Microsoft\Edge\Application\msedge.exe" (
    echo Edge: %PROGRAMFILES%\Microsoft\Edge\Application\msedge.exe
)
if exist "%PROGRAMFILES(x86)%\Microsoft\Edge\Application\msedge.exe" (
    echo Edge: %PROGRAMFILES(x86)%\Microsoft\Edge\Application\msedge.exe
)
if exist "%PROGRAMFILES%\Mozilla Firefox\firefox.exe" (
    echo Firefox: %PROGRAMFILES%\Mozilla Firefox\firefox.exe
)

REM 创建输出目录
echo 创建输出目录...
if not exist output mkdir output

echo.
echo === 安装完成 ===
echo.
echo 使用方法：
echo.
echo 1. 运行跨平台读取工具：
echo    cd scripts
echo    python wecom_cross_platform.py
echo.
echo 2. 或者手动启动浏览器：
echo    "%PROGRAMFILES%\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222
echo.
echo 3. 在浏览器中登录企业微信
echo.
echo 4. 运行读取脚本：
echo    python wecom_cross_platform.py
echo.
pause

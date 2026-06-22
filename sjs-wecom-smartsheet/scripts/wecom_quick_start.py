#!/usr/bin/env python3
"""
快速启动脚本（增强版）
自动检测默认浏览器，支持选择浏览器
"""

import asyncio
import json
import os
import sys
import socket
from datetime import datetime

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from browser_utils import BrowserDetector, DebugPortChecker


def show_browser_info():
    """显示浏览器信息"""
    detector = BrowserDetector()
    info = detector.get_browser_info()

    print("="*60)
    print("🌐 浏览器信息")
    print("="*60)
    print(f"\n操作系统: {info['system']}")
    print(f"默认浏览器: {info['default_browser'] or '未知'}")
    print(f"推荐浏览器: {info['recommended'] or '未知'}")

    print(f"\n可用浏览器:")
    for name, path in info['available_browsers'].items():
        is_default = " ⭐ (默认)" if name == info['default_browser'] else ""
        print(f"  - {name}: {path}{is_default}")

    return detector


def show_launch_command(detector: BrowserDetector, browser_name: Optional[str] = None):
    """显示启动命令"""
    browser_path = detector.get_browser(browser_name)

    if not browser_path:
        print("\n❌ 未找到可用的浏览器")
        print("请先安装 Chrome、Edge 或 Firefox")
        return None

    display_name = detector.get_browser_name(browser_path)
    port = DebugPortChecker.get_available_port()

    print("\n" + "="*60)
    print("🚀 启动命令")
    print("="*60)

    print(f"\n浏览器: {display_name}")
    print(f"调试端口: {port}")

    print(f"\n请在终端中执行以下命令：\n")

    if detector.system == "windows":
        print(f'"{browser_path}" --remote-debugging-port={port}')
    else:
        print(f'"{browser_path}" --remote-debugging-port={port} &')

    print(f"\n" + "="*60)
    print("📋 后续步骤")
    print("="*60)
    print("\n1. 执行上述命令启动浏览器")
    print("2. 在浏览器中打开 https://doc.weixin.qq.com/")
    print("3. 登录企业微信")
    print("4. 打开你要读取的智能表格")
    print("5. 重新运行此脚本，选择'连接并读取数据'")

    return port


def check_browser_running(port: int) -> bool:
    """检查浏览器是否在运行"""
    return DebugPortChecker.check_port(port)


async def connect_and_read(port: int):
    """连接并读取数据"""
    try:
        from playwright.async_api import async_playwright

        # 使用 workspace 目录存放输出
        workspace_dir = os.path.expanduser(
            "~/work_space/ai/skill/sjs-skills/sjs-wecom-smartsheet/workspace"
        )
        os.makedirs(workspace_dir, exist_ok=True)

        # 按日期创建子目录
        from datetime import datetime
        date_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.join(workspace_dir, date_str)
        os.makedirs(output_dir, exist_ok=True)

        print("\n" + "="*60)
        print("📊 正在读取数据")
        print("="*60)

        async with async_playwright() as p:
            # 连接到浏览器
            print(f"\n连接到浏览器 (端口: {port})...")
            browser = await p.chromium.connect_over_cdp(
                f"http://localhost:{port}",
                timeout=10000
            )
            print("✅ 连接成功")

            # 查找智能表格页面
            page = None
            for context in browser.contexts:
                for p in context.pages:
                    if "doc.weixin.qq.com" in p.url or "smartsheet" in p.url:
                        page = p
                        break
                if page:
                    break

            if not page:
                print("❌ 未找到企业微信页面")
                print("请确保在浏览器中打开了企业微信智能表格")
                return

            print(f"✅ 找到页面: {page.url}")

            # 截取页面
            screenshot_path = os.path.join(output_dir, "current_page.png")
            await page.screenshot(path=screenshot_path)
            print(f"✅ 截图已保存: {screenshot_path}")

            # 获取页面内容
            body_text = await page.inner_text("body")
            text_file = os.path.join(output_dir, "page_content.txt")
            with open(text_file, 'w', encoding='utf-8') as f:
                f.write(body_text)
            print(f"✅ 内容已保存: {text_file}")

            # 尝试截取 Canvas
            try:
                canvas_data = await page.evaluate('''() => {
                    const canvas = document.getElementById('jumper-canvas');
                    if (!canvas) return null;
                    return canvas.toDataURL('image/png');
                }''')

                if canvas_data and canvas_data.startswith('data:image'):
                    import base64
                    canvas_path = os.path.join(output_dir, "canvas_screenshot.png")
                    base64_data = canvas_data.split(',')[1]
                    with open(canvas_path, 'wb') as f:
                        f.write(base64.b64decode(base64_data))
                    print(f"✅ Canvas 截图已保存: {canvas_path}")
            except Exception as e:
                print(f"⚠️  Canvas 截取失败: {e}")

            # 保存原始数据
            data = {
                "url": page.url,
                "title": await page.title(),
                "timestamp": datetime.now().isoformat(),
                "text_content": body_text[:5000]
            }

            data_file = os.path.join(output_dir, "raw_data.json")
            with open(data_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            print(f"✅ 数据已保存: {data_file}")

            print("\n" + "="*60)
            print("✅ 数据读取完成！")
            print("="*60)
            print(f"\n输出文件：")
            print(f"  - 截图: {screenshot_path}")
            print(f"  - 内容: {text_file}")
            print(f"  - 数据: {data_file}")

    except Exception as e:
        print(f"\n❌ 连接失败: {e}")
        print("\n请确保：")
        print("1. 浏览器已启动并登录企业微信")
        print("2. 已打开智能表格页面")
        print("3. 使用了正确的调试端口")


def main():
    """主函数"""
    # 显示浏览器信息
    detector = show_browser_info()

    # 检查默认端口是否被占用
    default_port = 9222
    port_in_use = check_browser_running(default_port)

    print("\n" + "="*60)
    print("📋 选择操作")
    print("="*60)

    if port_in_use:
        print(f"\n检测到端口 {default_port} 已被占用，可能浏览器已在运行")
        print("\n请选择：")
        print("1. 连接并读取数据（浏览器已在运行）")
        print("2. 显示启动命令（使用其他浏览器）")
        print("3. 退出")
    else:
        print("\n请选择：")
        print("1. 显示启动命令（使用默认浏览器）")
        print("2. 显示启动命令（选择其他浏览器）")
        print("3. 连接并读取数据（如果浏览器已在运行）")
        print("4. 退出")

    choice = input("\n请输入选择 (1/2/3/4): ").strip()

    if port_in_use:
        if choice == "1":
            asyncio.run(connect_and_read(default_port))
        elif choice == "2":
            # 选择浏览器
            browsers = detector.list_browsers()
            print(f"\n可用浏览器：")
            for i, name in enumerate(browsers, 1):
                is_default = " (默认)" if name == detector.default_browser else ""
                print(f"  {i}. {name}{is_default}")

            browser_choice = input("\n请选择浏览器编号: ").strip()
            if browser_choice.isdigit():
                idx = int(browser_choice) - 1
                if 0 <= idx < len(browsers):
                    show_launch_command(detector, browsers[idx])
    else:
        if choice == "1":
            show_launch_command(detector)
        elif choice == "2":
            # 选择浏览器
            browsers = detector.list_browsers()
            print(f"\n可用浏览器：")
            for i, name in enumerate(browsers, 1):
                is_default = " (默认)" if name == detector.default_browser else ""
                print(f"  {i}. {name}{is_default}")

            browser_choice = input("\n请选择浏览器编号: ").strip()
            if browser_choice.isdigit():
                idx = int(browser_choice) - 1
                if 0 <= idx < len(browsers):
                    show_launch_command(detector, browsers[idx])
        elif choice == "3":
            port = int(input(f"\n请输入调试端口 (默认 {default_port}): ").strip() or default_port)
            if check_browser_running(port):
                asyncio.run(connect_and_read(port))
            else:
                print(f"\n❌ 端口 {port} 未被占用，浏览器可能未启动")
                print("请先启动浏览器")


if __name__ == "__main__":
    main()

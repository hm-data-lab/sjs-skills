#!/usr/bin/env python3
"""
跨平台浏览器工具（增强版）
自动检测操作系统和本地默认浏览器
支持 Windows、macOS、Linux
"""

import os
import sys
import platform
import subprocess
import shutil
import json
from typing import Optional, List, Dict


class BrowserDetector:
    """浏览器检测器"""

    def __init__(self):
        self.system = platform.system().lower()  # windows, darwin, linux
        self.browsers = self._detect_browsers()
        self.default_browser = self._detect_default_browser()

    def _detect_browsers(self) -> Dict[str, str]:
        """检测已安装的浏览器"""
        browsers = {}

        # 定义各平台的浏览器路径
        browser_paths = {
            "windows": {
                "chrome": [
                    os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
                    os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
                    os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
                ],
                "edge": [
                    os.path.expandvars(r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"),
                    os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"),
                ],
                "firefox": [
                    os.path.expandvars(r"%ProgramFiles%\Mozilla Firefox\firefox.exe"),
                    os.path.expandvars(r"%ProgramFiles(x86)%\Mozilla Firefox\firefox.exe"),
                ],
                "ie": [
                    os.path.expandvars(r"%ProgramFiles%\Internet Explorer\iexplore.exe"),
                ],
            },
            "darwin": {  # macOS
                "chrome": [
                    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                ],
                "edge": [
                    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
                ],
                "firefox": [
                    "/Applications/Firefox.app/Contents/MacOS/firefox",
                ],
                "safari": [
                    "/Applications/Safari.app/Contents/MacOS/Safari",
                ],
                "opera": [
                    "/Applications/Opera.app/Contents/MacOS/Opera",
                ],
                "brave": [
                    "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser",
                ],
            },
            "linux": {
                "chrome": [
                    "/usr/bin/google-chrome",
                    "/usr/bin/google-chrome-stable",
                    "/usr/bin/chromium-browser",
                    "/usr/bin/chromium",
                    "/snap/bin/chromium",
                ],
                "edge": [
                    "/usr/bin/microsoft-edge",
                    "/usr/bin/microsoft-edge-stable",
                ],
                "firefox": [
                    "/usr/bin/firefox",
                    "/usr/bin/firefox-esr",
                    "/snap/bin/firefox",
                ],
                "opera": [
                    "/usr/bin/opera",
                    "/snap/bin/opera",
                ],
            },
        }

        # 检测当前平台的浏览器
        if self.system in browser_paths:
            for browser_name, paths in browser_paths[self.system].items():
                for path in paths:
                    if os.path.exists(path):
                        browsers[browser_name] = path
                        break

        return browsers

    def _detect_default_browser(self) -> Optional[str]:
        """检测默认浏览器"""
        try:
            if self.system == "darwin":
                # macOS: 使用 defaults 命令读取默认浏览器
                result = subprocess.run(
                    ["defaults", "read", "com.apple.LaunchServices/com.apple.launchservices.secure", "LSHandlers"],
                    capture_output=True,
                    text=True,
                    timeout=5
                )
                if result.returncode == 0:
                    # 解析输出，查找默认浏览器
                    lines = result.stdout.split('\n')
                    for i, line in enumerate(lines):
                        if 'https' in line:
                            # 下一行通常是浏览器标识
                            if i + 1 < len(lines):
                                next_line = lines[i + 1]
                                if 'com.google.chrome' in next_line:
                                    return 'chrome'
                                elif 'com.microsoft.edgemac' in next_line:
                                    return 'edge'
                                elif 'org.mozilla.firefox' in next_line:
                                    return 'firefox'
                                elif 'com.apple.Safari' in next_line:
                                    return 'safari'

            elif self.system == "windows":
                # Windows: 读取注册表
                try:
                    import winreg
                    key = winreg.OpenKey(
                        winreg.HKEY_CURRENT_USER,
                        r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\https\UserChoice"
                    )
                    prog_id = winreg.QueryValueEx(key, "ProgId")[0]
                    winreg.CloseKey(key)

                    if 'Chrome' in prog_id:
                        return 'chrome'
                    elif 'Edge' in prog_id:
                        return 'edge'
                    elif 'Firefox' in prog_id:
                        return 'firefox'
                except:
                    pass

            elif self.system == "linux":
                # Linux: 使用 xdg-settings
                result = subprocess.run(
                    ["xdg-settings", "get", "default-web-browser"],
                    capture_output=True,
                    text=True,
                    timeout=5
                )
                if result.returncode == 0:
                    browser = result.stdout.strip()
                    if 'chrome' in browser.lower():
                        return 'chrome'
                    elif 'edge' in browser.lower():
                        return 'edge'
                    elif 'firefox' in browser.lower():
                        return 'firefox'

        except Exception as e:
            pass

        # 如果无法检测默认浏览器，返回第一个可用的浏览器
        if self.browsers:
            return list(self.browsers.keys())[0]

        return None

    def get_browser(self, preferred: Optional[str] = None) -> Optional[str]:
        """获取可用的浏览器"""
        # 如果指定了首选浏览器，优先使用
        if preferred and preferred in self.browsers:
            return self.browsers[preferred]

        # 使用默认浏览器
        if self.default_browser and self.default_browser in self.browsers:
            return self.browsers[self.default_browser]

        # 按优先级选择浏览器
        priority = ["chrome", "edge", "firefox", "safari", "opera", "brave"]
        for browser in priority:
            if browser in self.browsers:
                return self.browsers[browser]

        return None

    def get_browser_name(self, browser_path: str) -> str:
        """根据路径获取浏览器名称"""
        for name, path in self.browsers.items():
            if path == browser_path:
                return name
        return "unknown"

    def list_browsers(self) -> List[str]:
        """列出所有可用的浏览器"""
        return list(self.browsers.keys())

    def get_browser_info(self) -> Dict:
        """获取浏览器信息"""
        return {
            "system": self.system,
            "default_browser": self.default_browser,
            "available_browsers": self.browsers,
            "recommended": self.get_browser()
        }


class BrowserLauncher:
    """浏览器启动器"""

    def __init__(self, debug_port: int = 9222):
        self.debug_port = debug_port
        self.detector = BrowserDetector()
        self.process = None

    def launch(self, browser_name: Optional[str] = None) -> bool:
        """启动浏览器"""
        browser_path = self.detector.get_browser(browser_name)

        if not browser_path:
            print(f"❌ 未找到可用的浏览器")
            print(f"可用的浏览器: {self.detector.list_browsers()}")
            return False

        browser_display_name = self.detector.get_browser_name(browser_path)
        print(f"正在启动浏览器: {browser_display_name}")
        print(f"路径: {browser_path}")
        print(f"调试端口: {self.debug_port}")

        # 构建启动命令
        cmd = [browser_path, f"--remote-debugging-port={self.debug_port}"]

        try:
            # 启动浏览器（非阻塞）
            if self.detector.system == "windows":
                self.process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=subprocess.CREATE_NEW_PROCESS_GROUP
                )
            else:
                self.process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    preexec_fn=os.setsid
                )

            print(f"✅ 浏览器已启动 (PID: {self.process.pid})")
            return True

        except Exception as e:
            print(f"❌ 启动失败: {e}")
            return False

    def is_running(self) -> bool:
        """检查浏览器是否运行"""
        if self.process:
            return self.process.poll() is None
        return False

    def terminate(self):
        """终止浏览器"""
        if self.process and self.is_running():
            self.process.terminate()
            print("浏览器已关闭")


class DebugPortChecker:
    """调试端口检查器"""

    @staticmethod
    def check_port(port: int) -> bool:
        """检查端口是否被占用"""
        import socket
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            result = sock.connect_ex(('localhost', port))
            return result == 0

    @staticmethod
    def get_available_port(start: int = 9222, end: int = 9322) -> int:
        """获取可用的端口"""
        import socket
        for port in range(start, end):
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
                result = sock.connect_ex(('localhost', port))
                if result != 0:
                    return port
        return start


def print_system_info():
    """打印系统信息"""
    print("="*60)
    print("系统信息")
    print("="*60)
    print(f"操作系统: {platform.system()} {platform.release()}")
    print(f"架构: {platform.machine()}")
    print(f"Python 版本: {platform.python_version()}")

    detector = BrowserDetector()
    info = detector.get_browser_info()

    print(f"\n默认浏览器: {info['default_browser'] or '未知'}")
    print(f"推荐浏览器: {info['recommended'] or '未知'}")

    browsers = detector.list_browsers()
    if browsers:
        print(f"可用浏览器: {', '.join(browsers)}")
    else:
        print("未检测到浏览器")


def main():
    """测试函数"""
    print_system_info()

    detector = BrowserDetector()
    print(f"\n检测到的浏览器:")
    for name, path in detector.browsers.items():
        is_default = " (默认)" if name == detector.default_browser else ""
        print(f"  {name}: {path}{is_default}")

    # 测试端口检查
    checker = DebugPortChecker()
    port = checker.get_available_port()
    print(f"\n可用调试端口: {port}")


if __name__ == "__main__":
    main()

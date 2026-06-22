#!/usr/bin/env python3
"""
企业微信智能表格自动导出脚本

简化流程：
1. --login 模式：启动可见浏览器，用户扫码登录，检测成功后关闭
2. 导出模式：用 headless 浏览器 + 已保存的登录态，自动导出 Excel
"""

import argparse
import asyncio
import json
import os
import sys
import glob
from datetime import datetime
from typing import Optional

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from browser_utils import BrowserDetector

# ============================================================
# 配置
# ============================================================

WEIXIN_DOC_URL = "https://doc.weixin.qq.com/"
LOGIN_WAIT_TIMEOUT = 120
LOGIN_CHECK_INTERVAL = 3


def load_config(skill_dir: str) -> dict:
    config_path = os.path.join(skill_dir, "config.json")
    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def ensure_workspace(workspace_dir: str) -> str:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = os.path.join(workspace_dir, ts)
    os.makedirs(output_dir, exist_ok=True)
    return output_dir


def get_storage_state_path() -> str:
    """获取 storage state 文件路径（保存 cookies 和 localStorage）"""
    skill_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    state_dir = os.path.join(skill_dir, ".browser-profile")
    os.makedirs(state_dir, exist_ok=True)
    return os.path.join(state_dir, "storage_state.json")


def get_user_data_dir() -> str:
    """获取专用 profile 目录"""
    skill_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    profile_dir = os.path.join(skill_dir, ".browser-profile")
    os.makedirs(profile_dir, exist_ok=True)
    return profile_dir


def print_step(step, total, msg):
    print(f"\n[{step}/{total}] {msg}")

def print_success(msg):
    print(f"  ✅ {msg}")

def print_info(msg):
    print(f"  ℹ️  {msg}")

def print_error(msg):
    print(f"  ❌ {msg}")

def print_waiting(msg):
    print(f"  ⏳ {msg}")


# ============================================================
# 登录检测
# ============================================================

async def check_login_status(page) -> bool:
    try:
        url = page.url
        if "login" in url.lower() or "passport" in url.lower():
            return False

        body_text = await page.inner_text("body")
        login_indicators = ["登录", "扫码", "请输入账号"]
        if any(i in body_text for i in login_indicators):
            doc_indicators = ["我的文档", "最近访问", "新建", "导入", "智能表格"]
            if any(i in body_text for i in doc_indicators):
                return True
            return False

        if body_text.strip():
            return True
        return False
    except Exception:
        return False


async def wait_for_login(page, timeout):
    elapsed = 0
    while elapsed < timeout:
        await asyncio.sleep(LOGIN_CHECK_INTERVAL)
        elapsed += LOGIN_CHECK_INTERVAL
        if await check_login_status(page):
            return True
        if elapsed % 15 == 0:
            print_waiting(f"已等待 {elapsed} 秒...")
    return False


# ============================================================
# --login 模式：启动可见浏览器让用户登录
# ============================================================

async def do_login(skill_dir: str):
    """启动可见浏览器，用户扫码登录，检测成功后关闭"""
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        print_error("需要安装 playwright: pip install playwright")
        sys.exit(1)

    user_data_dir = get_user_data_dir()

    print_info(f"使用浏览器 profile: {user_data_dir}")

    async with async_playwright() as p:
        # 用持久化上下文启动可见浏览器
        context = await p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            headless=False,
            channel="msedge",
            args=["--disable-blink-features=AutomationControlled"],
        )

        page = context.pages[0] if context.pages else await context.new_page()

        # 导航到企微文档
        await page.goto(WEIXIN_DOC_URL, wait_until="domcontentloaded", timeout=30000)
        await asyncio.sleep(3)

        print(f"\n  当前页面: {page.url}")

        # 检测登录状态
        is_logged_in = await check_login_status(page)

        if is_logged_in:
            print_success("已检测到登录状态，无需重新登录")
        else:
            print_waiting("请在浏览器中扫码或登录企业微信")
            print_waiting(f"等待登录完成（最多 {LOGIN_WAIT_TIMEOUT} 秒）...")
            is_logged_in = await wait_for_login(page, LOGIN_WAIT_TIMEOUT)

        if is_logged_in:
            # 保存 storage state（cookies + localStorage）
            storage_path = get_storage_state_path()
            state = await context.storage_state()
            with open(storage_path, "w", encoding="utf-8") as f:
                json.dump(state, f, ensure_ascii=False, indent=2)

            print_success("登录成功！")
            print_info(f"登录态已保存: {storage_path}")
            print_info("浏览器将在 3 秒后关闭")
            await asyncio.sleep(3)
        else:
            print_error("登录超时")
            print_info("浏览器保持打开，你可以稍后重试")

        await context.close()


# ============================================================
# 验证检测
# ============================================================

async def _check_verification(page) -> bool:
    """检测是否弹出验证（滑块、二维码等）"""
    try:
        url = page.url
        # 企微验证页面的 URL 特征
        if "verify" in url.lower() or "captcha" in url.lower() or "security" in url.lower():
            return True

        # 检查页面内容是否有验证关键词
        body = await page.inner_text("body")
        verify_keywords = ["滑动验证", "请完成验证", "安全验证", "请拖动", "slide to verify",
                           "verification", "请验证", "身份验证", "操作频繁"]
        if any(kw in body for kw in verify_keywords):
            return True

        # 检查是否有验证相关的 DOM 元素
        has_verify_el = await page.evaluate("""() => {
            const selectors = [
                '[class*="verify"]', '[class*="captcha"]',
                '[class*="slide"]', '[id*="verify"]',
                '[class*="tcaptcha"]',  # 腾讯验证码
                'iframe[src*="verify"]', 'iframe[src*="captcha"]',
            ];
            for (const sel of selectors) {
                const el = document.querySelector(sel);
                if (el) {
                    const r = el.getBoundingClientRect();
                    if (r.width > 50 && r.height > 50) return true;
                }
            }
            return false;
        }""")

        return has_verify_el

    except Exception:
        return False


async def _wait_for_verification(page, timeout=120):
    """等待用户完成验证"""
    print_waiting("请在浏览器中完成验证（滑块/扫码等）...")
    elapsed = 0
    while elapsed < timeout:
        await asyncio.sleep(3)
        elapsed += 3

        # 检查验证是否已消失
        still_verify = await _check_verification(page)
        if not still_verify:
            print_success("验证已完成")
            return True

        if elapsed % 15 == 0:
            print_waiting(f"已等待 {elapsed} 秒，仍在等待验证...")

    print_error("验证等待超时")
    return False


# ============================================================
# 导出模式：headless 浏览器自动导出
# ============================================================

async def do_export(
    output_dir: str,
    table_name: Optional[str] = None,
    skill_dir: str = "",
):
    """用浏览器导出 Excel（headless 优先，验证时自动切换可见浏览器）"""
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        print_error("需要安装 playwright: pip install playwright")
        sys.exit(1)

    storage_path = get_storage_state_path()
    total_steps = 3

    async with async_playwright() as p:

        async def _try_export(headless=True):
            """尝试导出，返回 (excel_path, need_fallback)"""
            mode = "headless" if headless else "可见"
            print_info(f"使用 {mode} 浏览器...")

            launch_args = ["--disable-blink-features=AutomationControlled"]
            browser = await p.chromium.launch(
                headless=headless,
                channel="msedge",
                args=launch_args,
            )

            context_kwargs = {"accept_downloads": True}
            if os.path.exists(storage_path):
                context_kwargs["storage_state"] = storage_path

            context = await browser.new_context(**context_kwargs)
            page = await context.new_page()

            try:
                # 导航
                await page.goto(WEIXIN_DOC_URL, wait_until="domcontentloaded", timeout=30000)
                await asyncio.sleep(3)

                url = page.url
                print_info(f"页面: {url}")

                # 检测是否需要登录或验证
                if "login" in url.lower():
                    if headless:
                        print_info("需要登录，切换到可见浏览器...")
                        await context.close()
                        await browser.close()
                        return None, True  # need_fallback
                    else:
                        print_waiting("请在浏览器中登录...")
                        is_logged_in = await wait_for_login(page, LOGIN_WAIT_TIMEOUT)
                        if not is_logged_in:
                            print_error("登录超时")
                            await context.close()
                            await browser.close()
                            return None, False
                        # 登录成功，保存状态
                        state = await context.storage_state()
                        with open(storage_path, "w", encoding="utf-8") as f:
                            json.dump(state, f, ensure_ascii=False, indent=2)
                        print_success("登录态已保存")

                # 检测是否弹出验证（滑块验证码等）
                has_verification = await _check_verification(page)
                if has_verification:
                    if headless:
                        print_info("检测到验证，切换到可见浏览器...")
                        await context.close()
                        await browser.close()
                        return None, True  # need_fallback
                    else:
                        print_waiting("请在浏览器中完成验证...")
                        await _wait_for_verification(page)
                        # 验证通过后保存状态
                        state = await context.storage_state()
                        with open(storage_path, "w", encoding="utf-8") as f:
                            json.dump(state, f, ensure_ascii=False, indent=2)
                        print_success("验证完成，登录态已保存")

                # 检测登录状态
                is_logged_in = await check_login_status(page)
                if not is_logged_in:
                    if headless:
                        print_info("未登录，切换到可见浏览器...")
                        await context.close()
                        await browser.close()
                        return None, True
                    else:
                        print_error("未登录")
                        await context.close()
                        await browser.close()
                        return None, False

                print_success("已登录企业微信")

                # 查找表格
                if table_name:
                    print_info(f"搜索表格: {table_name}")
                    target = await search_and_open(page, table_name, context)
                else:
                    print_info("展示文档列表...")
                    target = await show_list_and_select(page, context)

                if not target:
                    print_error("未找到目标智能表格")
                    await context.close()
                    await browser.close()
                    return None, False

                page = target
                print_success(f"已打开智能表格: {await page.title()}")

                # 导出
                excel_path = await export_excel(page, output_dir)

                # 导出后也检查是否触发了验证
                if not excel_path:
                    has_verification = await _check_verification(page)
                    if has_verification and headless:
                        print_info("导出时触发验证，切换到可见浏览器...")
                        await context.close()
                        await browser.close()
                        return None, True

                await context.close()
                await browser.close()
                return excel_path, False

            except Exception as e:
                print_error(f"出错: {e}")
                try:
                    await context.close()
                    await browser.close()
                except Exception:
                    pass
                return None, headless  # headless 时出错则 fallback

        # --------------------------------------------------
        # 主流程：先 headless，失败则切换可见浏览器
        # --------------------------------------------------
        print_step(1, total_steps, "启动浏览器")
        excel_path, need_fallback = await _try_export(headless=True)

        if need_fallback:
            print_info("切换到可见浏览器，请在弹出的窗口中操作...")
            excel_path, _ = await _try_export(headless=False)

        if excel_path:
            print_step(3, total_steps, "分析数据")
            print_success(f"Excel 已保存: {excel_path}")

            analyze_script = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                "wecom_analyze_excel.py",
            )
            if os.path.exists(analyze_script):
                import subprocess
                result = subprocess.run(
                    [sys.executable, analyze_script, excel_path],
                    capture_output=True, text=True,
                )
                if result.returncode == 0:
                    print_success("分析完成")
                else:
                    print_error(f"分析失败: {result.stderr[:200]}")
        else:
            print_error("导出失败")


# ============================================================
# 文档查找
# ============================================================

async def search_and_open(page, table_name, context):
    """搜索并打开指定名称的表格"""
    # 按名称搜索
    search_input = await page.query_selector('input[type="search"], input[placeholder*="搜索"]')
    if search_input:
        await search_input.click()
        await search_input.fill(table_name)
        await asyncio.sleep(1.5)

    # 找到匹配的文档名
    name_el = await page.query_selector(f'.fileList_item_name:has-text("{table_name}")')
    if not name_el:
        items = await page.query_selector_all('.fileList_item_name')
        for item in items:
            text = (await item.inner_text()).strip()
            if table_name in text:
                name_el = item
                break

    if name_el:
        return await open_doc(name_el, context)
    return None


async def show_list_and_select(page, context):
    """展示文档列表让用户选择"""
    items = await page.query_selector_all('.fileList_item')
    if not items:
        return None

    names = []
    for item in items:
        name_el = await item.query_selector('.fileList_item_name')
        if name_el:
            names.append((await name_el.inner_text()).strip())

    for i, name in enumerate(names[:15]):
        print(f"  {i + 1}. {name}")

    try:
        choice = input("\n请输入文档编号: ").strip()
        if choice.isdigit():
            idx = int(choice) - 1
            if 0 <= idx < len(items):
                name_el = await items[idx].query_selector('.fileList_item_name')
                if name_el:
                    return await open_doc(name_el, context)
    except (EOFError, KeyboardInterrupt):
        pass
    return None


async def open_doc(name_el, context):
    """单击文档名打开（新 tab）"""
    old_urls = set(pg.url for pg in context.pages)
    old_count = len(context.pages)

    await name_el.click()
    await asyncio.sleep(4)

    # 检查是否有新页面，或者当前页面 URL 变了
    if len(context.pages) > old_count:
        # 新 tab 打开了
        for pg in context.pages:
            if pg.url not in old_urls:
                return pg

    # 没有新 tab，检查第一个页面是否导航了
    if context.pages:
        current = context.pages[0]
        if current.url not in old_urls and "/sheet/" in current.url:
            return current

    # 兜底：返回最后一个页面
    if context.pages:
        return context.pages[-1]

    return None


# ============================================================
# 导出 Excel
# ============================================================

async def export_excel(page, output_dir):
    """导出 Excel：菜单 → 导出 → hover → .xlsx"""
    try:
        # 关闭可能的弹窗
        await page.keyboard.press("Escape")
        await asyncio.sleep(0.3)
        await page.mouse.click(600, 400)
        await asyncio.sleep(0.3)

        # 点击文件菜单
        menu_btn = await page.query_selector('.menu-button-file')
        if not menu_btn:
            btn_info = await page.evaluate("""() => {
                const el = document.querySelector('[class*="menu-button-file"]');
                if (el) {
                    const r = el.getBoundingClientRect();
                    return {x: r.x + r.width/2, y: r.y + r.height/2};
                }
                return null;
            }""")
            if btn_info:
                await page.mouse.click(btn_info['x'], btn_info['y'])
            else:
                print_error("未找到文件菜单按钮")
                await page.screenshot(path=os.path.join(output_dir, "debug_no_menu.png"))
                return None
        else:
            await menu_btn.click()

        await asyncio.sleep(1.5)

        # 找导出项
        export_info = await page.evaluate("""() => {
            const allEls = document.querySelectorAll('*');
            let best = null, bestArea = Infinity;
            for (const el of allEls) {
                const text = el.innerText?.trim();
                if (text === '导出') {
                    const r = el.getBoundingClientRect();
                    const a = r.width * r.height;
                    if (a > 50 && a < 20000 && r.height > 15 && r.height < 60 && a < bestArea) {
                        bestArea = a;
                        best = {x: r.x + r.width/2, y: r.y + r.height/2};
                    }
                }
            }
            return best;
        }""")

        if not export_info:
            print_error("未找到导出菜单项")
            await page.screenshot(path=os.path.join(output_dir, "debug_no_export.png"))
            return None

        # 触发子菜单展开
        # 方式1：JS dispatchEvent（兼容 headless 模式）
        await page.evaluate("""() => {
            const exportItem = document.querySelector('.mainmenu-submenu-exportAs');
            if (exportItem) {
                for (const evt of ['mouseenter', 'mouseover', 'pointerenter']) {
                    exportItem.dispatchEvent(new MouseEvent(evt, {
                        bubbles: true, cancelable: true, view: window,
                    }));
                }
            }
        }""")
        await asyncio.sleep(1)

        # 检查子菜单是否可见，如果不可见用 mouse.move 兜底
        vis = await page.evaluate("""() => {
            const el = document.querySelector('.mainmenu-item-export-local');
            return el ? window.getComputedStyle(el).visibility : null;
        }""")
        if vis != "visible":
            await page.mouse.move(export_info['x'], export_info['y'])
            await asyncio.sleep(0.5)
            for step in range(10):
                await page.mouse.move(export_info['x'] - step * 30, export_info['y'])
                await asyncio.sleep(0.1)
            await asyncio.sleep(1)

        # 找 .xlsx（优先用 class 精确匹配）
        xlsx_info = await page.evaluate("""() => {
            // 优先1：class 精确匹配 mainmenu-item-export-local
            const exportLocal = document.querySelector('.mainmenu-item-export-local');
            if (exportLocal) {
                const r = exportLocal.getBoundingClientRect();
                if (r.width > 10 && r.height > 10) {
                    return {x: r.x + r.width/2, y: r.y + r.height/2, text: exportLocal.innerText?.trim()};
                }
            }
            // 优先2：在子菜单 wrapper 中查找
            const wrappers = document.querySelectorAll('.dui-menu-submenu-wrapper');
            for (const w of wrappers) {
                const r = w.getBoundingClientRect();
                if (r.width > 100 && r.height > 30 && r.y > 100) {
                    for (const el of w.querySelectorAll('*')) {
                        const text = el.innerText?.trim();
                        if (text && text.includes('.xlsx') && text.length < 40) {
                            const er = el.getBoundingClientRect();
                            if (er.width > 10 && er.height > 10) {
                                return {x: er.x + er.width/2, y: er.y + er.height/2, text};
                            }
                        }
                    }
                }
            }
            // 兜底：文本匹配
            let best = null, bestArea = Infinity;
            for (const el of document.querySelectorAll('*')) {
                const text = el.innerText?.trim();
                if (text && text.includes('Excel') && text.includes('.xlsx') && text.length < 40) {
                    const r = el.getBoundingClientRect();
                    const a = r.width * r.height;
                    if (a > 10 && a < 10000 && r.height > 10 && r.height < 50 && a < bestArea) {
                        bestArea = a;
                        best = {x: r.x + r.width/2, y: r.y + r.height/2, text};
                    }
                }
            }
            return best;
        }""")

        if not xlsx_info:
            print_error("未找到 xlsx 导出选项")
            await page.screenshot(path=os.path.join(output_dir, "debug_no_xlsx.png"))
            return None

        print_info(f"导出: {xlsx_info['text']}")

        async with page.expect_download(timeout=30000) as dl:
            await page.mouse.click(xlsx_info['x'], xlsx_info['y'])

        download = await dl.value
        filename = download.suggested_filename or f"export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        save_path = os.path.join(output_dir, filename)
        await download.save_as(save_path)
        return save_path

    except Exception as e:
        print_error(f"导出失败: {e}")
        try:
            await page.screenshot(path=os.path.join(output_dir, "debug_export_error.png"))
        except Exception:
            pass
        return None


# ============================================================
# 主入口
# ============================================================

def main():
    parser = argparse.ArgumentParser(description="企业微信智能表格自动导出")
    parser.add_argument("--workspace", help="输出目录路径")
    parser.add_argument("--table", help="目标表格名称（可选）")
    parser.add_argument("--skill-dir", help="Skill 安装目录")
    parser.add_argument("--login", action="store_true", help="启动浏览器进行登录")
    args = parser.parse_args()

    skill_dir = args.skill_dir or os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    config = load_config(skill_dir)
    workspace_dir = args.workspace or config.get("workspace_dir", os.path.join(skill_dir, "workspace"))
    output_dir = ensure_workspace(workspace_dir)

    print("=" * 60)
    print("🚀 企业微信智能表格自动导出")
    print("=" * 60)

    if args.login:
        asyncio.run(do_login(skill_dir))
    else:
        asyncio.run(do_export(
            output_dir=output_dir,
            table_name=args.table,
            skill_dir=skill_dir,
        ))

    print("\n" + "=" * 60)
    print("✅ 完成")
    print("=" * 60)


if __name__ == "__main__":
    main()

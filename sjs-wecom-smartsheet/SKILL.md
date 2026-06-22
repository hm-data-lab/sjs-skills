---
name: sjs-wecom-smartsheet
description: 读取和分析企业微信智能表格数据。当用户提到要读取企业微信表格、分析智能表格、导出企微数据、查看企微表格内容时触发。即使用户只说"帮我看看企微表格"、"读一下那个智能表格"、"企微数据导出来分析一下"等口语化表达，只要涉及企业微信智能表格的读取或分析，也应触发此 Skill。
---

# 企业微信智能表格读取分析

**重要：执行时直接使用 `scripts/` 下已有的脚本，不要重新开发新脚本。**

---

## 整体流程

```
用户："帮我看看企微表格"
  │
  ├─ 检查 config.json → 不存在则引导配置 workspace_dir
  │
  ├─ 检查 storage_state.json（登录态）
  │   ├─ 存在 → 直接走方式一（自动导出）
  │   └─ 不存在 → 走方式一的登录流程
  │
  ├─ 方式一：自动导出 Excel（推荐）
  │   ├─ 首次：--login 启动可见浏览器 → 用户登录 → 保存登录态
  │   └─ 后续：headless 浏览器 + 登录态 → 自动导出 → 分析 → 报告
  │
  ├─ 方式二：手动导出 Excel（备选）
  │   └─ 用户手动导出 → 提供文件路径 → 分析 → 报告
  │
  └─ 方式三：在线读取（最后备选）
      └─ 连接已有浏览器 → 截图 + 提取页面内容
```

---

## 三种方式对比

| 优先级 | 方式 | 命令 | 用户操作 | 适用场景 |
|--------|------|------|---------|---------|
| ⭐⭐⭐ | **自动导出** | `wecom_auto_export.py` | 仅首次登录 | 日常使用 |
| ⭐⭐ | **手动导出** | `wecom_analyze_excel.py` | 手动导出+给路径 | 自动导出失败时 |
| ⭐ | **在线读取** | `wecom_quick_start.py` | 启动浏览器+登录 | 需要实时页面数据 |

---

## 方式一：自动导出 Excel

### 首次使用：登录

```bash
python3 scripts/wecom_auto_export.py --login
```

流程：
1. 启动可见浏览器，打开 `https://doc.weixin.qq.com/`
2. 用户在浏览器中登录（扫码或快捷登录）
3. 脚本检测登录成功 → 保存登录态到 `.browser-profile/storage_state.json`
4. 浏览器自动关闭

### 后续使用：导出

```bash
# 指定表格名
python3 scripts/wecom_auto_export.py --table "数据所专利统计"

# 不指定表格名（展示列表让用户选）
python3 scripts/wecom_auto_export.py
```

流程：
1. headless 浏览器加载 storage_state（自动登录）
2. 导航到企微文档 → 检测登录状态
3. 搜索/选择目标智能表格 → 单击打开（新 tab）
4. 点击文件菜单 `.menu-button-file`
5. JS 触发 `mouseenter` 展开"导出"子菜单
6. 点击 `mainmenu-item-export-local`（本地 Excel 表格 .xlsx）
7. 下载 Excel 到 `{workspace_dir}/{时间戳}/`
8. 调用 `wecom_analyze_excel.py` 分析数据
9. 输出 Markdown 报告

### 智能降级机制

脚本自动处理登录和验证，**大部分情况无需用户参与**：

```
headless 浏览器启动
  ├─ 登录态有效 → 自动完成导出
  ├─ 需要登录 → 自动切换可见浏览器 → 用户登录 → 保存登录态 → 继续
  ├─ 触发验证 → 自动切换可见浏览器 → 用户完成验证 → 保存状态 → 继续
  └─ 导出时触发验证 → 自动切换可见浏览器 → 用户完成验证 → 重试导出
```

只有在需要登录或验证时，才会弹出浏览器窗口让用户操作。

---

## 方式二：手动导出 Excel

用户手动导出后，提供文件路径：

```bash
python3 scripts/wecom_analyze_excel.py /path/to/导出的文件.xlsx
```

分析报告包含：Sheet 数、行列数、表头、前 10 行数据样本。

---

## 方式三：在线读取

连接已运行的浏览器（需带调试端口），截图并提取页面内容：

```bash
python3 scripts/wecom_quick_start.py
```

---

## 配置

首次使用时检查 `config.json`，不存在则引导配置：

| 配置项 | 说明 | 默认值 |
|--------|------|--------|
| `workspace_dir` | 输出目录 | `{skill目录}/workspace` |

---

## 脚本清单

| 脚本 | 功能 |
|------|------|
| `wecom_auto_export.py` | 自动导出 Excel + 分析 |
| `wecom_analyze_excel.py` | 分析 Excel 文件 |
| `wecom_quick_start.py` | 浏览器在线读取 |
| `browser_utils.py` | 跨平台浏览器检测 |

---

## 依赖

```bash
pip install openpyxl playwright
```

---

## 工作目录

输出按时间戳组织：`{workspace_dir}/{YYYYMMDD_HHMMSS}/`

```
├── *.xlsx                    # 下载的 Excel
├── analysis_report_*.md     # 分析报告
└── excel_data.json          # 原始数据
```

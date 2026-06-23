# sjs-wecom-smartsheet

企业微信智能表格读取分析 Skill

## 触发条件

当你需要读取或分析企业微信智能表格时，可以使用此 Skill。例如：
- "帮我看看企微表格"
- "读一下那个智能表格"
- "企微数据导出来分析一下"
- "分析一下 XX 表格的数据"

## 用户需要提供什么

### 方式一：自动导出（推荐）

**首次使用需要：**
1. 企业微信账号（扫码或快捷登录）

**后续使用：**
- 表格名称（可选，如不提供会展示列表让你选）

示例对话：
```
你：帮我看看企微表格
AI：检测到未配置，请先登录...
    [启动浏览器，你扫码登录]
    AI：登录成功！请告诉我表格名称，或者查看所有可用文档？
你：数据所专利统计
AI：正在导出...分析完成，报告已生成
```

### 方式二：手动导出

**需要提供：**
- 已导出的 Excel 文件路径

示例：
```
你：分析一下这个表格 ~/Downloads/数据所专利统计.xlsx
AI：正在分析...报告已生成
```

### 方式三：在线读取

**需要提供：**
- 已打开的智能表格页面（浏览器需开启调试端口）

## 输出内容

每次运行会在 `workspace/{时间戳}/` 生成：

| 文件 | 说明 |
|------|------|
| `*.xlsx` | 下载的 Excel 文件（自动导出时） |
| `analysis_report_*.md` | Markdown 分析报告 |
| `excel_data.json` | 原始 JSON 数据 |
| `current_page.png` | 页面截图（在线读取时） |

## 快速开始

### 1. 安装依赖

```bash
cd sjs-wecom-smartsheet
bash scripts/install.sh  # macOS/Linux
# 或
scripts/install.bat       # Windows
```

### 2. 首次登录

```bash
python3 scripts/wecom_auto_export.py --login
```

按提示在浏览器中登录企业微信，登录成功后自动保存登录态。

### 3. 导出并分析

```bash
# 指定表格名
python3 scripts/wecom_auto_export.py --table "数据所专利统计"

# 或不指定，从列表中选择
python3 scripts/wecom_auto_export.py
```

### 4. 查看报告

报告保存在 `workspace/{时间戳}/analysis_report_*.md`

## 目录结构

```
sjs-wecom-smartsheet/
├── SKILL.md                       # Skill 定义文件
├── README.md                      # 本文件
├── scripts/
│   ├── wecom_auto_export.py       # 自动导出 + 分析（推荐）
│   ├── wecom_analyze_excel.py     # 分析已有 Excel 文件
│   ├── wecom_quick_start.py       # 浏览器在线读取
│   ├── browser_utils.py           # 跨平台浏览器检测
│   ├── install.sh                 # macOS/Linux 安装脚本
│   └── install.bat                # Windows 安装脚本
├── workspace/                     # 输出目录（按时间戳组织）
├── .browser-profile/              # 浏览器登录态存储
└── tests/                         # 测试文件
```

## 三种方式对比

| 方式 | 用户操作 | 适用场景 |
|------|---------|---------|
| **自动导出** | 仅首次登录 | 日常使用（推荐） |
| **手动导出** | 手动导出 Excel + 提供路径 | 自动导出失败时 |
| **在线读取** | 启动浏览器 + 登录 | 需要实时页面数据 |

## 常见问题

**Q: 登录态过期了怎么办？**
重新运行 `python3 scripts/wecom_auto_export.py --login` 即可。

**Q: 找不到目标表格？**
运行 `python3 scripts/wecom_auto_export.py --list` 查看所有可用文档。

**Q: 分析报告包含哪些内容？**
包含 Sheet 数、行列数、表头信息、前 10 行数据样本、数据概览统计。

## 依赖

- Python 3.8+
- openpyxl（Excel 分析）
- playwright（浏览器自动化）

```bash
pip install openpyxl playwright
playwright install chromium
```

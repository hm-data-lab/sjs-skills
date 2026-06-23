# sjs-wecom-smartsheet

企业微信智能表格读取分析 Skill

## 触发条件

当你需要读取或分析企业微信智能表格时，可以使用此 Skill。

**口语化触发词：**
- "帮我看看企微表格"
- "读一下那个智能表格"
- "企微数据导出来分析一下"
- "分析一下 XX 表格的数据"

## 整体流程

```
你："帮我看看企微表格"
  │
  ├─ 检查登录态
  │   ├─ 已登录 → 直接进入自动导出
  │   └─ 未登录 → 启动浏览器让你登录
  │
  ├─ 方式一：自动导出 Excel（推荐）
  │   ├─ 首次：启动可见浏览器 → 你登录 → 保存登录态
  │   └─ 后续：自动导出 → 分析 → 生成报告
  │
  ├─ 方式二：手动导出 Excel（备选）
  │   └─ 你手动导出 → 提供文件路径 → 分析 → 报告
  │
  └─ 方式三：在线读取（最后备选）
      └─ 连接浏览器 → 截图 + 提取内容
```

## 三种方式

| 优先级 | 方式 | 你需要做的 | 适用场景 |
|--------|------|-----------|---------|
| ⭐⭐⭐ | **自动导出** | 仅首次登录 | 日常使用 |
| ⭐⭐ | **列出文档** | 登录后查看 | 查看可用文档 |
| ⭐⭐ | **手动导出** | 手动导出+给路径 | 自动导出失败时 |
| ⭐ | **在线读取** | 启动浏览器+登录 | 需要实时页面数据 |

---

## 方式一：自动导出 Excel

### 首次使用

1. **登录**
   ```bash
   python3 scripts/wecom_auto_export.py --login
   ```
   - 启动可见浏览器，打开企微文档
   - 你在浏览器中登录（扫码或快捷登录）
   - 登录成功后自动保存登录态

2. **导出**
   ```bash
   # 指定表格名
   python3 scripts/wecom_auto_export.py --table "数据所专利统计"

   # 或从列表中选择
   python3 scripts/wecom_auto_export.py
   ```

### 后续使用

直接运行导出命令，脚本会自动使用保存的登录态。

### 智能降级

脚本自动处理各种情况，**大部分情况无需你参与**：
- 登录态有效 → 自动完成
- 需要登录 → 弹出浏览器让你登录
- 触发验证 → 弹出浏览器让你完成验证

---

## 方式二：手动导出 Excel

如果你已经手动导出了 Excel 文件，提供文件路径即可：

```bash
python3 scripts/wecom_analyze_excel.py /path/to/导出的文件.xlsx
```

---

## 方式三：在线读取

连接已运行的浏览器（需带调试端口），截图并提取页面内容：

```bash
python3 scripts/wecom_quick_start.py
```

---

## 输出内容

每次运行会在 `workspace/{时间戳}/` 生成：

```
workspace/20260623_114751/
├── *.xlsx                    # 下载的 Excel（自动导出时）
├── analysis_report_*.md      # 分析报告
├── excel_data.json           # 原始数据
├── current_page.png          # 页面截图（在线读取时）
└── page_content.txt          # 页面文本（在线读取时）
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

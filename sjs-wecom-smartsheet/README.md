# sjs-wecom-smartsheet

企业微信智能表格读取分析 Skill

## 功能

- 自动检测本地默认浏览器（Chrome/Edge/Firefox）
- 读取企业微信智能表格数据（浏览器自动化 / Excel 导出）
- 生成 Markdown 分析报告
- 支持 Windows、macOS、Linux

## 快速开始

### 方式 1：Excel 导出（推荐）

```bash
python3 scripts/wecom_analyze_excel.py ~/Downloads/你的文件.xlsx
```

### 方式 2：浏览器自动化

```bash
python3 scripts/wecom_quick_start.py
```

1. 按提示启动浏览器（带调试端口）
2. 登录企业微信，打开智能表格
3. 重新运行脚本读取数据

## 目录结构

```
sjs-wecom-smartsheet/
├── SKILL.md                       # Skill 定义文件
├── README.md                      # 说明文档
├── scripts/
│   ├── wecom_quick_start.py       # 浏览器自动化快速启动
│   ├── wecom_analyze_excel.py     # Excel 文件分析
│   ├── browser_utils.py           # 跨平台浏览器检测工具
│   ├── install.sh                 # macOS/Linux 依赖安装
│   └── install.bat                # Windows 依赖安装
├── tests/                         # 测试文件
├── assets/                        # 资源文件
└── references/                    # 参考文档
```

## 工作目录

输出目录可通过 `config.json` 配置（`workspace_dir` 字段），默认为 `{skill目录}/workspace`。按时间戳组织：

```
{workspace_dir}/
├── 20260622_103059/
│   ├── analysis_report_*.md     # 分析报告
│   ├── current_page.png         # 页面截图
│   ├── page_content.txt         # 页面文本
│   └── raw_data.json            # 原始数据
└── README.md
```

## 依赖

- Python 3.8+
- openpyxl（Excel 分析必需）
- playwright（浏览器自动化必需）

```bash
pip install openpyxl playwright
```

## 许可证

MIT License

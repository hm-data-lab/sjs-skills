# Changelog

All notable changes to the sjs-skills repository will be documented in this file.

Format based on [Keep a Changelog](https://keepachangelog.com/).

---

## [Unreleased]

### sjs-wecom-smartsheet

#### Added

- 企业微信智能表格读取分析 Skill
- 自动检测本地默认浏览器（Chrome/Edge/Firefox）
- 跨平台支持（Windows/macOS/Linux）
- 快速启动脚本 `wecom_quick_start.py`
- Excel 分析工具 `wecom_analyze_excel.py`
- 浏览器检测工具 `browser_utils.py`

### sjs-ipd-ppt

#### Added

- PPT 输出文件名保留大纲版本号，实现大纲与 PPT 版本一一对应
  - 之前：`{项目名}-CHARTER-汇报材料.pptx`（每次生成覆盖同一文件）
  - 之后：`{项目名}-CHARTER-v{N}-汇报材料.pptx`（每个大纲版本对应独立 PPT 文件）

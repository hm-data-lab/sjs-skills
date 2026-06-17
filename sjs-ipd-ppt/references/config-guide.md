# 配置指南

本 Skill 需要配置一些路径和参数才能正常工作。首次使用时会自动引导配置。

## 配置文件位置

`config.json`，位于 Skill 根目录下（与 `SKILL.md` 同级）。

## 配置项说明

```json
{
  "template_path": "",
  "project_base_dir": "",
  "cost_per_day": 970,
  "product_lines": {}
}
```

### template_path（必需）

公司统一 PPT 模板文件的完整路径。

- 类型：字符串
- 格式：绝对路径或相对于当前工作目录的路径
- 示例：`"/home/user/company/模板/company-template.pptx"`

**如何获取：** 向团队管理员询问公司 PPT 模板文件位置，或在公司共享盘搜索 `.pptx` 模板文件。

### project_base_dir（推荐）

项目资料存放的根目录。用于建议输出路径。

- 类型：字符串
- 格式：目录路径
- 示例：`"/home/user/projects/"`

**如何获取：** 确认团队统一的项目资料存放位置。

### cost_per_day（可选）

人力成本基准，单位：元/人天。

- 类型：整数
- 默认值：`970`
- 说明：用于自动计算人力费用（人天数 × cost_per_day）

### product_lines（可选）

产品线配置，每个产品线可指定独立的模板。

- 类型：对象
- 格式：

```json
{
  "product_lines": {
    "产品线名称": {
      "template": "该产品线的模板路径"
    }
  }
}
```

- 示例：

```json
{
  "product_lines": {
    "空调": {
      "template": "/path/to/ac-template.pptx"
    },
    "美菱": {
      "template": "/path/to/meiling-template.pptx"
    }
  }
}
```

**说明：** 如果产品线没有独立模板，可以不配置，脚本会使用 `template_path` 作为默认模板。

## 首次配置流程

当 Skill 检测到 `config.json` 不存在时，会引导用户完成以下步骤：

1. **确认模板路径** — 询问公司 PPT 模板文件位置
2. **确认项目目录** — 询问项目资料存放位置
3. **确认人力成本** — 确认人天单价（默认 970）
4. **确认产品线** — 询问有哪些产品线及对应模板
5. **保存配置** — 写入 `config.json`

## 跳过配置

如果用户只需要输出结构化要点（不需要生成 PPT 文件），可以跳过模板配置。Skill 仍然可以提供完整的引导式提问和内容框架。

## 更新配置

编辑 `config.json` 即可。常见更新场景：

- 公司模板更新了 → 修改 `template_path`
- 新增产品线 → 在 `product_lines` 中添加
- 人力成本调整 → 修改 `cost_per_day`

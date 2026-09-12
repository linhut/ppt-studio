<!--
  (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
  https://github.com/linhut/ppt-studio
  Licensed under the MIT License. See the LICENSE file for details.
-->

# Changelog

本仓库采用 [Semantic Versioning](https://semver.org/lang/zh-CN/)（主.次.修订）。

## v1.0.0 (2026-09-12)

### Added

- **初始同步（仓库建立）**：ppt-studio 以独立 GitHub 仓库 `linhut/ppt-studio` 公开
  - 双引擎：PPTD 路径（YAML 清单 + 逐页 .page，55173 编辑器联动）与 JSON 数据驱动路径（复合组件 + $theme 变量）
  - 核心脚本：`export_pptx.py` / `inline_bg.py` / `json2pptx.py` / `presets.py` / `quality_check.py` / `wizard.py`
  - 11 个 references 设计文档、模板（templates/）与示例（examples/，仅代码/数据文本，不含预览图与二进制）
  - 20 套配色（`templates/themes/all-themes.yaml`）+ 18 个复合组件布局库
- **统一命令行入口 `ppt`**（`scripts/ppt.py` + `ppt.bat` / `ppt.sh`）：build / export / check / render / presets / doctor / new / wizard / md2deck / validate / pptx2json / version，一条命令覆盖全流程
- **工程化工具链**：
  - `ppt doctor` 环境预检（依赖 / PowerPoint / LibreOffice 渲染通道）
  - `ppt new` 新 deck 脚手架（--theme / --pages）
  - `ppt wizard` 交互式向导（6 问 + --defaults / --answers 非交互模式）
  - `ppt md2deck` Markdown 大纲 → deck.json（数字型条目自动 KPI 布局，`<!-- layout: -->` 提示）
  - `ppt validate` deck.json schema 校验（JSONPath + 行号定位）
  - `ppt pptx2json` PPTX → deck.json 反向导入（矢量结构回填，图片占位不落二进制）
- **渲染与质量工程化**：
  - `quality_check --html` / `--html-out` 可视化审查报告（逐页卡片 + 缩略图 + 渲染检测结果）
  - `render_check` LibreOffice headless 兜底渲染（soffice → PDF → PNG，无 PowerPoint 也可出报告）
- **JSON 引擎扩展**：
  - `theme.fonts` 字体配置（heading / body / stat / table / tagline）
  - `chartType` 图表类型：bar / line / pie / radar / stacked / scatter，支持多系列（单位列自动跳过）
  - canvas 可配置（4:3 / 16:10 等任意尺寸）
- **工程质量**：`requirements.txt` 依赖清单 / pytest 黄金结构单测（tests/，21 项）/ GitHub Actions CI（.github/workflows/ci.yml）
- **发布纪律**：
  - 仅同步代码/文档/模板文本；图片、PPTX 产物、预览截图、本地导出脚本（`_export.ps1`）、`__pycache__` 等一律不提交（见 `.gitignore` 与 `RELEASING.md`）
  - 干净 git 历史，主分支 `main`，tag 格式 `vX.Y.Z`，GitHub Release 编号与 tag 一致
- **作者信息**：所有代码/文档文件标注 `(c) 2026 Jose AI (https://github.com/linhut/ppt-studio)`，仓库统一按 MIT License 授权（第三方编译内核 `scripts/LICENSE.np-ppt` 随附保留）

### Notes

- 首个正式版本（仓库此前无发布 tag）；本次为功能完整首发，后续功能迭代按 SemVer 递增（1.0.x / 1.1.x）
- 渲染级检查在未安装 PowerPoint / LibreOffice 的机器上自动跳过（仅警告不失败）
<!--
  (c) 2026 Jose AI (https://www.linhut.cn)
  https://github.com/linhut/ppt-studio
  Licensed under the MIT License. See the LICENSE file for details.
-->

# Changelog

本仓库采用 [Semantic Versioning](https://semver.org/lang/zh-CN/)（主.次.修订）。

## Unreleased

### Added

- **初始同步（仓库建立，不做版本发布）**：ppt-studio 以独立 GitHub 仓库 `linhut/ppt-studio` 公开
  - 双引擎：PPTD 路径（YAML 清单 + 逐页 .page，55173 编辑器联动）与 JSON 数据驱动路径（复合组件 + $theme 变量）
  - 6 个核心脚本：`export_pptx.py` / `inline_bg.py` / `json2pptx.py` / `presets.py` / `quality_check.py` / `wizard.py`
  - 11 个 references 设计文档、模板（templates/）与示例（examples/，仅代码/数据文本，不含预览图与二进制）
  - 20 套配色（`templates/themes/all-themes.yaml`）+ 18 个复合组件布局库
- **发布纪律**：
  - 仅同步代码/文档/模板文本；图片、PPTX 产物、预览截图、本地导出脚本（`_export.ps1`）、`__pycache__` 等一律不提交（见 `.gitignore` 与 `RELEASING.md`）
  - 干净 git 历史，主分支 `main`，tag 格式 `vX.Y.Z`，GitHub Release 编号与 tag 一致
- **作者信息**：所有代码/文档文件标注 `(c) 2026 Jose AI (https://www.linhut.cn)`，仓库统一按 MIT License 授权（第三方编译内核 `scripts/LICENSE.np-ppt` 随附保留）

### Notes

- 本次为仓库初始同步（含作者标注等微小调整），**不打 tag、不建 Release、不做版本发布**；版本发布流程见 `RELEASING.md`
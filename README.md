<!--
  (c) 2026 Jose AI (https://www.linhut.cn)
  https://github.com/linhut/ppt-studio
  Licensed under the MIT License. See the LICENSE file for details.
-->
# ppt-studio — PPT 全能工坊 (DSH Skill)

> `PPT` · `PPTX` · `Presentation` · `Deck` · `DSH Skill` · `Python` · `PPTD`

一站式生成**原生 .pptx 演示文稿**的 DSH 技能：融合多套设计体系的设计思路，
把「设计蓝图」与「生成引擎」合二为一，产出真正的 PowerPoint 文件（不是网页）。

## 双引擎

| | PPTD 经典路径 | JSON 数据驱动路径 |
|---|---|---|
| 数据 | deck.deck.pptd + pages/*.page | 单文件 deck.json |
| 元素 | shape/line/text/table/chart/image | 上述全部 + **复合组件**（卡片网格/统计卡/目录列表/总结条/大数字） |
| 背景 | 必须内联（inline_bg.py 辅助） | 可直接用 $theme 变量 + 支持渐变 |
| 适用 | 55173 编辑器联动的精细 deck | 数据驱动、批量、Agent 自动生成 |
| 编译 | `scripts/export_pptx.py` | `scripts/json2pptx.py` |
| 审查 | 文档流程 | `scripts/quality_check.py` 自动打分 |

## 能力一览

- 🎨 **18 套成品配色 + 4 大设计预设**（academic/consultant/business/tech），见 `references/themes.md`、`scripts/presets.py`
- 📐 **31 类页面布局配方 + 12 个 JSON 布局模板**（封面/目录/KPI/时间轴/对比/图表/致谢…）
- 🧩 **PPTD DSL**：YAML 清单 + 模块化 .page，一页一文件，16:9（960×540pt）画布
- ⚡ **高保真编译**：微软雅黑、noAutofit、富文本、表格、图表、淡入淡出
- 🔀 **元素路由器**：text/image/chart/table/shape + cards_2x3/cards_1x4_info/card_list_wide/tagline_bar/num_big
- 🧪 **质量审查**：5 维度自动打分 + 硬约束校验（越界/间距/字数/视觉占比）
- 🗣️ **演讲者逐字稿**：每页 150–300 字口语化讲稿（notes 字段）

## 快速开始（PPTD 路径）

```bash
# 1. 复制示例工程开始
cp -r examples/tech-share my-deck
# 2. 改 deck.deck.pptd 里的 theme（配色从 templates/themes/all-themes.yaml 挑）
# 3. 逐页编辑 pages/*.page（布局配方见 references/layouts.md）
# 4. 内联背景色（用 $bg/$ink 可读写法后必跑）
py -3 scripts/inline_bg.py my-deck/deck.deck.pptd
# 5. 编译输出原生 PPTX
py -3 scripts/export_pptx.py my-deck/deck.deck.pptd --output my-deck/deck.pptx --force
```

## 快速开始（JSON 路径）

```bash
# 1. 复制 JSON 示例（单文件，Agent 友好）
cp -r examples/json-demo my-deck
# 2. 编辑 my-deck/deck.json：换 theme.colors / slides / elements
# 3. 编译（背景/文字直接写 $theme 变量即可，无需内联）
py -3 scripts/json2pptx.py my-deck/deck.json -o my-deck/deck.pptx
# 4. 质量审查（自动打分 + 硬约束校验）
py -3 scripts/quality_check.py my-deck/deck.json --preset tech
# 5. 查可用预设
py -3 scripts/presets.py list
```

> 依赖：Python 3 + `pyyaml` + `python-pptx`（编译器首次运行自动 pip 安装）。

## 目录结构

```
ppt-studio/
├── SKILL.md                    # 技能主入口（工作流）
├── README.md                   # 项目说明（本文件）
├── LICENSE                     # MIT License（(c) 2026 Jose AI）
├── CHANGELOG.md                # 版本变更记录（SemVer）
├── RELEASING.md                # 版本发布流程（Git 仓库 + GitHub Release，无 npm/pip）
├── references/                 # 设计文档（11 个，均含作者标注）
│   ├── palettes.md             # 20 套配色 + 色阶 + 透明度 + 字体规范
│   ├── themes.md               # 配色 → theme 变量映射 + 风格配方
│   ├── layouts.md              # 31 类页面布局配方
│   ├── animations.md           # 动画思路 → PPTX 方案
│   ├── pptd.md                 # PPTD DSL 语法参考
│   ├── json-engine.md          # JSON 引擎参考（元素/复合组件/预设/审查）
│   ├── layout-catalog.md       # 18 个复合组件速查 + 布局选择指南
│   ├── tone-guide.md           # 7 种语言风格 + 写作约束
│   ├── consulting-style.md     # 咨询/学术级内容纪律
│   ├── design-system.md        # 多源融合设计规范（含容量精算）
│   └── presenter-mode.md       # 演讲者逐字稿指南
├── templates/
│   ├── deck.deck.pptd          # PPTD 清单模板
│   ├── pages/                  # PPTD 页面模板（5 种）
│   ├── themes/all-themes.yaml  # 20 套配色合集（机器可读）
│   └── json/
│       ├── deck.json           # JSON 引擎模板骨架
│       └── layout-templates.json  # 12 个布局模板（占位符式）
├── scripts/
│   ├── export_pptx.py          # PPTD → PPTX 编译内核（第三方内核，随附 LICENSE.np-ppt）
│   ├── inline_bg.py            # $theme 背景色自动内联
│   ├── json2pptx.py            # JSON → PPTX 引擎（复合组件 + 渐变 + fade）
│   ├── presets.py              # 设计预设（4 大 + 18 配色加载）
│   ├── quality_check.py        # 质量审查（5 维度 + 硬约束）
│   ├── wizard.py               # answers.json → deck.json 骨架生成
│   └── LICENSE.np-ppt          # 编译内核许可证（第三方）
└── examples/                   # 示例（仅代码/数据文本；预览图与 PPTX 产物不公开）
    ├── tech-share/             # PPTD 完整示例（7 页）
    ├── pptd-compare/           # PPTD 对比演示（6 页）
    ├── json-demo/              # JSON 引擎示例（4 页）
    ├── engine-test/            # JSON 引擎回归测试（9 页）
    ├── layout-showcase/        # 15 个复合组件布局展示（12 页）
    ├── gov-report/             # 政务风格年度总结（9 页）
    ├── academic-demo/          # 学术演示（4 页）
    └── csbm-research/          # 常山北明行情研究（12 页）
```

## 公开范围（只同步代码文件）

本仓库**只提交代码 / 文档 / 模板文本**，以下内容不公开：

- 图片与二进制：`*.png` / `*.gif` / `*.PNG` / `*.pptx`（含 examples 预览截图、palette-sample 等）
- 本地导出辅助脚本：`examples/*/_export.ps1`（含本机绝对路径，仅本地预览用）
- 参考素材副本：`references/sources-gorden/`（来源项目副本，仅本地参考）
- 缓存与构建产物：`__pycache__/`、`*.pyc`、`dist/`、`build/` 等（见 `.gitignore`）

## 版本发布

遵循仓库内 [RELEASING.md](RELEASING.md)：SemVer 版本规范 + `CHANGELOG.md` 条目 +
注解 tag（`vX.Y.Z`）+ GitHub Release。**不涉及 npm / pip 发布**；
作者信息类微调（版权行/作者标注）直接提交推送，不做版本发布。

## 安装到 DSH（可选，安全且不影响现有环境）

1. **方式 A（推荐）**：整个目录复制/链接到 DSH 技能目录，纯新增、不动现有文件：
   ```bash
   cp -r ppt-studio ~/.dsh/skills/ppt-studio
   ```
2. **方式 B**：不安装，工作目录直接使用（在任何 DSH 会话里读取 `SKILL.md` 即生效）。

> 加载后新会话的技能目录会出现 `ppt-studio`。对 DSH 无任何副作用——不修改
> 现有技能、不碰配置文件、不新增运行服务。

## 许可

- 仓库整体：**MIT License**，Copyright (c) 2026 **Jose AI**（https://www.linhut.cn），见 [LICENSE](LICENSE)
- 作者信息标注：所有代码/文档文件头部标注 `(c) 2026 Jose AI (https://www.linhut.cn)`；
  `SKILL.md` 通过 frontmatter `metadata` 声明；数据/模板文件（`.json`/`.page`/`.yaml`/`.pptd`）
  为保证解析安全不加注释，版权由本 LICENSE 统一声明
- 第三方组件许可证随附保留：编译内核 `scripts/export_pptx.py`（MIT），
  见 [scripts/LICENSE.np-ppt](scripts/LICENSE.np-ppt)

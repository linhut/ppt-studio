---
name: ppt-studio
description: "PPT 全能工坊 — 一站式生成原生 .pptx 演示文稿。当用户要求做PPT、演示文稿、幻灯片、汇报、演讲稿、路演、pitch deck 或配色排版时使用。内置 PPTD DSL 与 JSON 双引擎、20 套配色、18 个复合组件、Python-PPTX 高保真编译内核。产出原生 PPTX（形状/文本框/图表/表格/淡入淡出），不是网页。Triggers: PPT, pptx, 演示, 幻灯片, 演讲稿, 汇报, 路演, 分享, deck, slides, presentation, 配色方案, 排版, keynotes."
whenToUse: "用户要求制作/优化演示文稿，或需要配色、排版、演讲逐字稿、原生 PPTX 交付时使用。仅当用户明确要求 HTML 网页演示或非 PPTX 格式时不用。"
disable-model-invocation: false
user-invocable: true
metadata:
  author: "Jose AI (https://github.com/linhut/ppt-studio)"
  repository: "https://github.com/linhut/ppt-studio"
  license: "MIT"
---

# ppt-studio — PPT 全能工坊

**目标**：把「设计思路」与「生成引擎」合二为一，一次对话产出**原生 .pptx 文件**。

**产出物**：`deck/deck.pptx` —— 原生 PowerPoint 文件（16:9，960×540pt 画布，微软雅黑，noAutofit，纯色背景，可选淡入淡出过渡）。不是 HTML，不是网页。

**双引擎**：
- **PPTD 路径**（默认）：YAML 清单 + 逐页 `.page`，配合 55173 编辑器精细控制 → [references/pptd.md](references/pptd.md)
- **JSON 路径**（3B 节）：单文件 deck.json，数据驱动、批量、Agent 友好 → [references/json-engine.md](references/json-engine.md)

---

## 何时使用

- 用户要求「做PPT / 演示文稿 / 幻灯片 / 演讲稿 / 汇报 / 路演 / pitch deck / 分享材料」
- 用户要求配色方案、排版优化、演讲者逐字稿、PPT 质量检查
- 需要交付**原生 .pptx**（可在 PowerPoint/WPS 打开编辑），而不是 HTML 网页

## 核心原则（所有页面的底层纪律）

1. **每页只做一件事**，相邻页不重复同一种布局
2. **颜色只用主题变量**，一页 ≤1 主色 + 1 强调色 + 黑白灰；禁 5-6 色混用
3. **容量克制**：每页 ≤80 字 + ≥2 种视觉元素（GordenPPTSkill 容量精算：CJK=1.0 / ASCII=0.5 视觉宽度 + 20% 容差，防溢出遮挡）
4. **层级清晰**：字号按 封面 54 / 页标题 40 / 小节 24 / 卡片 18 / 正文 16 / 弱化 13（pt）；标题加粗，正文不加粗
5. **留白优先**：页边距 ≥40pt、卡片间距 24pt、卡片圆角 12pt（Soft & Balanced）、标题-内容间距 ≥24pt
6. **字体固定**：中文微软雅黑 / 英文 Arial（可选 Georgia/Calibri/Cambria/Trebuchet MS）
7. **最终交付原生 .pptx**，不是网页

---

## 工作流程（必须按顺序）

### 0. 向导式需求确认（必做，先问再做）

> 动手生成 PPT 之前，用 `ask_user_question` 一次发 6 问；用户已明确给出答案的可跳过对应问题，信息不足的问题**必须提问**，默认值标注「推荐」并放选项第一位。

| # | 问题 | 选项（推荐放第一） | 映射到 |
|---|---|---|---|
| 1 | **主题内容** | 能力展示 / 产品介绍 / 汇报 / 技术分享 / 自定义 | 文案与页面结构 |
| 2 | **页数结构** | 8-9 页完整（推荐）/ 5-6 页精简 / 12+ 大 deck | 页面模板组合 |
| 3 | **配色方案** | 科技→纯净科技蓝·科技与夜景·活力与科技；汇报→商务与权威·铂金白金·冷白调研；品牌→轻奢与神秘·黑金实验·深蓝杂志 | theme.colors（[palettes.md](references/palettes.md)） |
| 4 | **风格配方** | Soft & Balanced（推荐）/ Sharp & Compact / Rounded & Spacious / Pill & Airy | 圆角/间距档位（[themes.md](references/themes.md)） |
| 5 | **深浅模式** | 浅底深字（推荐，正文页）/ 深底浅字（发布会/高端）/ 混排 | bg/text 明暗组合 |
| 6 | **语言风格** | 商务正式（推荐）/ 轻松口语 / 科技极客 / 叙事故事 / 说服销售 / 学术报告 / 政务公文体 | 文案腔调（[tone-guide.md](references/tone-guide.md)） |

**交互后**：把 6 个答案整理为 answers.json（结构见 [templates/wizard/answers.json](templates/wizard/answers.json)），生成骨架并填充内容：

```bash
py -3 scripts/wizard.py examples/my-deck/answers.json --output examples/my-deck/deck.json
```

**进阶选项**：answers.json 可加 `content_layouts`（缺省 cards → kpi → process → timeline → compare，相邻页互不重复）与 `tone`（缺省 corporate，保证全文文案风格一致）。可选布局与 7 种腔调见 [references/layout-catalog.md](references/layout-catalog.md)、[references/tone-guide.md](references/tone-guide.md)。

### 1. 选配色 → 生成 theme

从 [references/palettes.md](references/palettes.md) 选定 1 套配色（或按 [references/themes.md](references/themes.md) 自定义），映射为 theme 变量：

```yaml
theme:
  colors:
    $bg: "#0A0A0F"       # 页面背景（纯色，禁渐变）
    $paper: "#F5F0E8"    # 浅色表面 / 卡片
    $primary: "#D4AF37"  # 主色：标题、强调
    $accent: "#0070F3"   # 强调色：数据、焦点
    $text: "#1F2937"     # 正文色
    $muted: "#6B7280"    # 弱化文字
    $line: "#E5E7EB"     # 线条 / 边框
    $ink: "#111827"      # 表头 / 深色块
    $green: "#10B981"    # 图表正色
    $brass: "#D4AF37"    # 装饰辅助色
  textStyles:
    $title: { fontSize: 36, bold: true, color: "$primary" }
    $h2:    { fontSize: 22, bold: true, color: "$text" }
    $body:  { fontSize: 14, bold: false, color: "$text" }
    $muted: { fontSize: 12, bold: false, color: "$muted" }
```

**字体规范（强制）**：中文微软雅黑，英文 Arial；中英混排时中文雅黑+英文所选字体。正文不加粗，标题才加粗。

### 2. 页面编排

把内容拆成标准页面类型（31 类布局配方见 [references/layouts.md](references/layouts.md)）：

- **开场**：cover（封面）、toc（目录 2×3 卡片）
- **过渡**：section-divider（章节页，大编号）
- **内容页**（选子类型，相邻页避免重复）：
  - 文本：bullets、two-column、three-column、big-quote
  - 数据：stat-highlight、kpi-grid、table、chart-bar/line/pie/radar
  - 代码：code、diff、terminal
  - 图示：flow-diagram、arch-diagram、process-steps、timeline、roadmap、gantt、comparison、pros-cons
  - 视觉：image-hero、image-grid
- **收尾**：cta、thanks

**编排规则**：每页只做一件事；正文左对齐，只有标题居中；标题 36pt+，正文 14–16pt，注释 10–12pt；页面边距 ≥0.3"，块间距 0.3–0.5"，留白优先。

### 3. 生成 PPTD 工程

```
deck/
├── deck.pptd              # YAML 清单：theme + pages 列表
└── pages/
    ├── 01_cover.page      # JSON/YAML，每页一个文件
    ├── 02_toc.page
    ├── 03_section.page
    ├── 04_content.page
    └── 99_final.page
```

**页面结构**（完整语法见 [references/pptd.md](references/pptd.md)）：

```yaml
id: "01_cover"
pageType: "cover"
background:
  color: "#0A0A0F"          # 必须纯色！禁 gradient/stops/angle
elements:
  - elementType: "shape"     # rect | roundRect | oval | donut
    shapeName: "roundRect"
    bounds: [60, 60, 840, 420]     # 960×540pt 画布
    fill: { type: "solid", color: "$paper" }
    border: { style: "solid", color: "$line", width: 1 }
  - elementType: "text"
    bounds: [120, 140, 720, 100]
    content:
      text: "<p>2026 年度汇报</p>"
      align: ["center", "middle"]
      style: "$title"
  - elementType: "line"      # 箭头：arrow: [0, "arrow"]
    bounds: [100, 300, 760, 2]
    border: { color: "$accent", width: 2 }
```

**元素类型**：`shape` / `line` / `text` / `table` / `chart` / `image`。
**富文本**：`content.text` 支持 `<p>`、`<span style="...">`、`<strong>`、`<em>`、`<u>`、`<br/>`；**颜色必须内联为 #HEX / rgb()，禁止 $theme 变量**。
**图表**：`elementType: "chart"`，`data.rows` 为 [[分类, 值, 值], ...]，编译为柱状图 + 数据标签。

### 4. 内联背景色（推荐）→ 编译 → 原生 PPTX

页面里用 `color: "$bg"` 可读写法没问题，但 55173 一致性校验要求 `background.color` 是内联色。先内联再编译：

```bash
py -3 scripts/inline_bg.py deck/deck.deck.pptd          # 把 pages/*.page 的 $theme 色内联为实际色值
py -3 scripts/export_pptx.py deck/deck.pptd --output deck/deck.pptx --force
```

编译器自动安装 `pyyaml` / `python-pptx` 依赖（如缺失），输出微软雅黑、noAutofit、淡入淡出过渡的高保真 PPTX（13.333×7.5 英寸，16:9，960×540pt）。

### 3B. JSON 数据驱动快速路径（可选）

数据驱动 / 批量 / Agent 自动生成时，跳过 PPTD 工程，用单文件 JSON：

1. 从 [templates/json/deck.json](templates/json/deck.json) 复制骨架，或从 [templates/json/layout-templates.json](templates/json/layout-templates.json) 挑布局模板粘贴 elements
2. 配色直接引用预设名（4 大预设 + 18 套配色），背景/填充/文字可写 `$primary` 等变量（本引擎自行解析，无内联限制），背景还支持渐变；字体可用 `theme.fonts` 配置（heading/body/stat/table/tagline），图表可用 `chartType`（bar/line/pie/radar/stacked/scatter 多系列）
3. 用复合组件表达数据（18 个，速查见 [references/layout-catalog.md](references/layout-catalog.md)）：`table` / `chart` / `cards_2x3` / `cards_1x4_info` / `card_list_wide` / `kpi_row` / `num_big` / `tagline_bar` / `section_divider` / `comparison_2col` / `pros_cons` / `timeline_h` / `process_steps` / `roadmap_4col` / `checklist` / `big_quote` / `cover_asym` / `figure_text` / `breadcrumb` / `references`
4. 编译 + 质量审查（推荐走统一 CLI `ppt`，脚本直调亦可）：

```bash
ppt build deck.json -o deck.pptx --strict                  # 编辑（淡入淡出）
ppt check deck.json --preset tech --auto-only              # 自动打分
ppt check deck.json --preset tech --auto-only --html       # HTML 审查报告（deck.quality.html）
ppt validate deck.json                                     # schema 校验（JSONPath + 行号）
ppt presets list                                           # 查预设
```

JSON 结构与全部元素类型见 [references/json-engine.md](references/json-engine.md)。

### 5. QA（必做）

1. `python -m markitdown deck/deck.pptx` 提取全部文本，核对内容/顺序/错别字
2. 检查占位符：`... | grep -iE "xxxx|lorem|ipsum|placeholder|TODO"`
3. 逐页核对：配色是否严格来自主题变量、文字是否越界/重叠、图表数据是否正确
4. **渲染级检查（推荐）**：把 deck 真实渲染成图做像素级检测，能抓静态估算看不见的问题——文字被实体遮挡、元素超出画布被裁剪、页面过密/过空（PowerPoint COM 优先，无 PowerPoint 时自动用 LibreOffice headless 兜底）：
   ```bash
   ppt render my-deck/deck.json            # 编译+渲染+检测（完整报告）
   ppt render my-deck/deck.pptx            # 直接检测已有 pptx
   ppt render my-deck/deck.json --strict   # 有未达标页时退出码 1（CI 可用）
   ppt check my-deck/deck.json --preset tech --auto-only --render  # 静态+渲染合并审查
   ppt check my-deck/deck.json --preset tech --auto-only --html    # 可视化报告（含渲染缩略图）
   ```
5. `ppt validate my-deck/deck.json` 先做 schema 校验（JSONPath + 行号定位），再至少完成一轮「发现问题→修复→再验证」循环，再交付

---

## 演讲者逐字稿（可选）

用户要「演讲 / 分享 / 讲稿 / 逐字稿」时，为每页写 150–300 字口语化讲稿，放入页面的 `notes` 字段（编译器不渲染，保留在备注里；或交付时附 `speaker-notes.md`）。写稿三规则：不是讲稿是提示信号（加粗核心词+过渡句独立成段）；每页 150–300 字（2–3 分钟/页）；用口语不用书面语（"因此"→"所以"）。详见 [references/presenter-mode.md](references/presenter-mode.md)。

---

## HARD-GATE（违反 = 返工）

| 规则 | 说明 |
|---|---|
| 背景纯色 | `background.color` 只允许 #HEX/rgb()/rgba()，**禁止** gradient/stops/angle |
| 背景免$theme | `background.color` 不能写 `$bg` 等变量——写完后跑 `scripts/inline_bg.py` 自动内联，或直接写实际色值 |
| 颜色内联 | `content.text` 内颜色用 #HEX/rgb()，**禁止** $theme |
| 严格色板 | 只用主题变量里的颜色，不发明新色 |
| 字体 | 中文微软雅黑 / 英文 Arial；正文不加粗 |
| 画布 | 960×540pt（16:9），bounds 为 pt 值 |
| 色值格式 | 编译器接受 #RRGGBB 或 #RRGGBBAA（带 alpha） |
| 淡入淡出 | 默认 `--transition fade`，需要静态可加 `--transition none` |
| JSON 引擎例外 | JSON 路径背景/填充/文字可直接用 `$theme` 变量，支持渐变与复合组件；PPTD 路径仍须内联背景 |

**编译前自检**：跑 `py -3 scripts/quality_check.py deck.json --preset <配色>` 检查容量溢出、越界、字号层级、颜色数；有「容量溢出」警告时先精简措辞，不要用省略号截断。交付前若本机装有 PowerPoint，加跑 `py -3 scripts/render_check.py deck.json --strict` 做渲染级复核（抓文字被遮挡、元素超出画布被裁剪等静态看不见的问题）。

## 反模式

- ❌ 一页混用 5-6 种颜色，或发明色板外的新色
- ❌ PPTD 页面写渐变背景 / `background.color` 用 `$theme` 变量不内联
- ❌ 正文加粗、字号随意（破坏层级）
- ❌ 文字溢出卡片或画布仍交付（必须过 quality_check + markitdown 核对）
- ❌ 残留 `xxxx` / `lorem` / `placeholder` / `TODO` 占位符
- ❌ 不跑向导式需求确认就动手生成
- ❌ 交付 HTML 网页冒充 PPT（本 skill 只交付原生 .pptx）
- ❌ 相邻页重复同一种布局，或一页塞多个主题

## 注意事项

- 命令统一用 `py -3`（Windows）或 `python3`（macOS/Linux）；依赖缺失时编译器自动安装
- PPTD 路径可与 55173 所见即所得编辑器联动；JSON 路径为纯命令行
- 背景内联仅在 PPTD 路径需要；JSON 引擎可直接用 `$theme` 变量与渐变
- 渲染级检查 `render_check.py` 需要本机安装 Microsoft PowerPoint（COM 渲染）或 LibreOffice（soffice 兜底，需 `pip install pymupdf`）+ Pillow/numpy；两者都没有时自动跳过（仅警告不失败）
- 本 skill 不修改现有文件、不新增服务；复制到 `~/.dsh/skills/ppt-studio` 即部署
- 许可与第三方组件许可证见 [README.md](README.md) 与 [scripts/LICENSE.np-ppt](scripts/LICENSE.np-ppt)

## 参考文件索引

- [references/palettes.md](references/palettes.md) — 20 套配色 + 色阶 + 透明度 + 字体规范
- [references/themes.md](references/themes.md) — 配色 → theme 变量映射 + 风格配方（圆角/间距）
- [references/layouts.md](references/layouts.md) — 31 类页面布局配方
- [references/animations.md](references/animations.md) — 动画思路（27 类 CSS 动画 → PPTX 过渡/入场效果）
- [references/pptd.md](references/pptd.md) — PPTD DSL 完整语法参考
- [references/json-engine.md](references/json-engine.md) — JSON 引擎参考（元素/复合组件/预设/审查）
- [references/layout-catalog.md](references/layout-catalog.md) — 布局样式库：18 个复合组件速查 + 按页面角色选布局
- [references/tone-guide.md](references/tone-guide.md) — 语言风格指南：7 种腔调 + 写作约束
- [references/consulting-style.md](references/consulting-style.md) — 咨询/学术级内容纪律：行动标题 + Ghost deck test + 留白 15-20%
- [references/presenter-mode.md](references/presenter-mode.md) — 演讲者逐字稿指南

## 模板与示例

- `templates/deck.deck.pptd` — 清单模板；`templates/pages/*.page` — 页面类型模板；`templates/themes/*.yaml` — 20 套配色预映射
- `templates/json/deck.json` — JSON 引擎模板骨架；`templates/json/layout-templates.json` — 12 个布局模板
- `examples/tech-share/`、`examples/json-demo/` — 双引擎可直接编译示例
- `examples/engine-test/` — JSON 引擎回归测试（9 页，100 分）
- `examples/layout-showcase/` — 15 个复合组件布局展示（12 页）
- `examples/gov-report/` — 政务风格年度总结（9 页，中国红配色 + 公文体）
- `examples/academic-demo/` — 学术演示（行动标题 + 结果页 + 参考文献）
- `examples/csbm-research/` — 常山北明（000158）市场行情与研究分析（12 页，真实行情数据）

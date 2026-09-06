<!--
  (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
  https://github.com/linhut/ppt-studio
  Licensed under the MIT License. See the LICENSE file for details.
-->
# PPTD DSL 语法参考（dsh-np-ppt 编译内核）

> 来源：<https://github.com/z953218350/dsh-np-ppt>（MIT）。编译器：`scripts/export_pptx.py`。
> PPTD = PowerPoint Template DSL：YAML 清单（deck.pptd）+ 每页一个 .page 文件（YAML/JSON），
> 由 Python-PPTX 编译为原生 PPTX。**本 skill 的输出物即 PPTD 工程。**

## 一、工程结构

```
deck/
├── deck.deck.pptd      # 也可以叫 deck.pptd；清单文件，含 theme + pages
└── pages/
    ├── 01_cover.page
    ├── 02_toc.page
    └── 99_final.page
```

编译器找清单：输入传 `.pptd` 文件，或项目目录（目录下**只能有 1 个** `.pptd`）。

**命令**：
```bash
py -3 scripts/export_pptx.py deck/deck.deck.pptd --output deck/deck.pptx --force
# 选项：--transition fade|none（默认 fade 淡入淡出）· --force 覆盖已存在输出
```

## 二、清单 deck.deck.pptd

```yaml
# 可选顶层字段
transition: "fade"          # 编译时可覆盖
theme:
  colors:                   # 颜色变量，供 $xxx 引用（背景/填充/文字）
    $bg: "#0A0A0F"
    $paper: "#F5F0E8"
    $primary: "#D4AF37"
    $accent: "#0070F3"
    $text: "#1F2937"
    $muted: "#6B7280"
    $line: "#E5E7EB"
    $ink: "#111827"
    $green: "#10B981"
    $brass: "#D4AF37"
  textStyles:               # 文字样式变量，供 content.style 引用
    $title: { fontSize: 36, bold: true, color: "$primary" }
    $h2:    { fontSize: 22, bold: true, color: "$text" }
    $h3:    { fontSize: 18, bold: true, color: "$text" }
    $body:  { fontSize: 14, bold: false, color: "$text", lineHeight: 1.4 }
    $muted: { fontSize: 12, bold: false, color: "$muted" }
    $stat:  { fontSize: 44, bold: true, color: "$accent" }
pages:
  - "pages/01_cover.page"
  - "pages/02_toc.page"
  - "pages/03_section.page"
  - "pages/04_content.page"
  - "pages/99_final.page"
```

## 三、页面 .page

每页一个 dict。顶层字段：`id`、`pageType`、`background`、`elements`、`notes`（可选，演讲者备注）。

### 3.1 background（背景）

**必须纯色**（编辑器与导出一致性铁律）：

```yaml
background:
  color: "#0A0A0F"       # #HEX / rgb() / rgba() / 命名色（不可用 $theme 变量，见一致性规则）
```

禁止：`type: gradient` / `stops` / `angle`。编译器虽支持渐变，但 55173 编辑器渲染为白底 → 预览≠导出。**本 skill 一律纯色。**

> 💡 **偷懒写法**：写页面时背景可以先用 `color: "$bg"` 保持可读，编译前运行
> `py -3 scripts/inline_bg.py deck.deck.pptd` 一键内联为实际色值（该脚本读取
> manifest 的 theme.colors 批量替换页面中所有 `color: "$xxx"`）。

### 3.2 elements（元素数组）

坐标 `bounds: [x, y, w, h]`，pt 值（画布 960×540）。支持 6 种 elementType：

#### A. shape（几何形状）

```yaml
- elementType: "shape"
  shapeName: "roundRect"      # rect | roundRect | oval | donut
  bounds: [40, 140, 420, 300]
  fill: { type: "solid", color: "$paper" }   # 或 color: "#RRGGBB"
  border: { style: "solid", color: "$line", width: 1 }
  # donut 可选：adjustments: [18000]（内径比例）
```

#### B. line（线条/箭头）

```yaml
- elementType: "line"
  bounds: [100, 220, 760, 2]   # 起点 x,y；w=长度，h≈2
  border: { color: "$accent", width: 2 }
  arrow: [0, "arrow"]          # 可选：末端箭头
```

#### C. text（文本框 + HTML 富文本）

```yaml
- elementType: "text"
  bounds: [60, 160, 400, 60]
  content:
    text: "<p>标题文字</p>"
    align: ["left", "middle"]      # 对齐：[水平 left|center|right, 垂直 top|middle|bottom]
    style: "$title"                # 继承 textStyles 变量（可省略，用下面字段）
    fontSize: 36                   # 覆盖/独立设置
    bold: true
    color: "$primary"              # $theme 变量（编译器解析）或内联 #HEX/rgb()
    lineHeight: 1.4
    letterSpacing: 0
```

**富文本 HTML**（编译器内 RichTextHTMLParser 支持）：
- 块级：`<p style="...">`，style 支持 `font-size`、`font-weight`、`color`、`text-align`、`margin-top`、`margin-bottom`、`line-height`、`letter-spacing`
- 行内：`<strong>`/`<b>`（加粗）、`<em>`/`<i>`（斜体）、`<u>`（下划线）、`<span style="...">`、`<br/>`（换行）
- **颜色必须内联为 #HEX/rgb()，禁止 $theme**（否则编辑器预览丢失颜色，见一致性规则）
- 无 HTML 标签时为纯文本段落

#### D. table（表格）

```yaml
- elementType: "table"
  bounds: [40, 140, 880, 320]
  rows:
    - [{ text: "指标" }, { text: "2025" }, { text: "2026" }, { text: "同比" }]
    - [{ text: "营收" }, { text: "1.2亿" }, { text: "1.8亿" }, { text: "+50%" }]
  columnWidths: [0.4, 0.2, 0.2, 0.2]   # 各列宽比例（和=1）
  rowHeights: [0.25, 0.25, 0.25, 0.25] # 可选
```

首行自动为表头（深色底 `$ink`，白字，居中）。奇偶行交替底色。边框颜色取 `$line`。

#### E. chart（柱状图）

```yaml
- elementType: "chart"
  bounds: [60, 150, 600, 330]
  data:
    rows:
      - ["Q1", 0, 32]     # [分类, 占位, 数值]
      - ["Q2", 0, 45]
      - ["Q3", 0, 58]
      - ["Q4", 0, 72]
```

编译为簇状柱状图：无图例无标题，柱子色 `$green`，数据标签在柱外，坐标轴用 `$line`，网格线浅色。取 rows 每行第 1 列（分类）与第 3 列（数值）。

#### F. image（图片）

```yaml
- elementType: "image"
  bounds: [40, 140, 420, 300]
  path: "imgs/photo.png"      # 相对 deck.deck.pptd 所在目录
```

### 3.3 notes（演讲者备注，可选）

```yaml
notes: |
  这一页要讲：先说结论，再给数据支撑。口语化，150–300 字。
```

编译器不输出备注到 PPTX 的备注页（保留字段即可，交付时附 `speaker-notes.md` 更稳妥）。

---

## 四、一致性校验（强制，违反即报错）

编译器启动时执行 `validate_deck_consistency`，违规直接抛错：

1. 每页必须有 `background` 块，且含 `color` 字段，值为 #HEX/rgb()/rgba()/命名色——**不能用 $theme**
2. 背景禁止渐变（`type/stops/angle`）
3. `content.text` 富文本内禁止 `color: $xxx` 写法——必须内联 #HEX/rgb()

> 技巧：纯命令行编译时 `$theme` 在 fill/border/color 字段是允许的（编译器能解析）；但 HTML 文本内一律内联，避免 55173 编辑器预览不一致。

## 五、输出规格

- 幻灯片：13.333" × 7.5"（16:9，= 960×540pt）
- 字体：中文微软雅黑 / 英文 Microsoft YaHei，`noAutofit` 锁定（防缩放）
- 过渡：默认 fade（淡入淡出），`--transition none` 关闭
- 依赖：`pyyaml`、`python-pptx`（缺失时编译器自动 pip 安装到临时目录）
- Python 命令探测：`py -3` → `python3` → `python`

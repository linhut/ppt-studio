<!--
  (c) 2026 Jose AI (https://www.linhut.cn)
  https://github.com/linhut/ppt-studio
  Licensed under the MIT License. See the LICENSE file for details.
-->
# 页面布局配方库（html-ppt 31 类布局 → PPTD 实现）

> 借鉴 html-ppt 的布局分类设计，将每一类页面翻译成 PPTD elements 配方。
> 画布 **960×540pt**（16:9）；bounds = [x, y, w, h]，pt 值。
> 排版纪律：页边距 ≥ 40pt，块间距 20–40pt，标题 36pt+，正文 14–16pt，左对齐（标题可居中）。

## 画布坐标系速查

```
+---------------------------------- 960 ----------------------------------+
| 40                                                                   40 |
|  ┌──────────────────────────────────────────────────────────────────┐  |
|  │               内容区：x 40–920，y 60–500                         │  |
|  └──────────────────────────────────────────────────────────────────┘  |
| 40                                                                   40 |
+---------------------------------- 540 ----------------------------------+
```

- 封面大标题区：y 150–220；副标题：y 260–300；信息行：y 420–460
- 章节页大编号：左上大字；标题：中左
- 正文页标题：y 60–100；内容起点 y 130

---

## 一、开场与过渡页

### 1. cover（封面）
配方：深色底（或主色底）+ 大标题 + 副标题 + 信息行（日期/演讲者）+ 装饰线条。

```yaml
id: "01_cover"
pageType: "cover"
background: { color: "$ink" }
elements:
  - { elementType: "shape", shapeName: "roundRect", bounds: [80, 150, 800, 8],
      fill: { type: "solid", color: "$brass" } }
  - { elementType: "text", bounds: [80, 60, 800, 40],
      content: { text: "<p>2026 · ANNUAL REPORT</p>", align: ["left", "middle"], style: "$muted" } }
  - { elementType: "text", bounds: [80, 180, 800, 80],
      content: { text: "<p>年度战略汇报</p>", align: ["left", "middle"], style: "$cover-title" } }
  - { elementType: "text", bounds: [80, 270, 800, 60],
      content: { text: "<p><span style='color:#bbb'>副标题一句话概括本次汇报主题</span></p>", align: ["left", "middle"], style: "$h3" } }
  - { elementType: "text", bounds: [80, 440, 800, 40],
      content: { text: "<p>汇报人：姓名　|　2026-01-15　|　会议厅</p>", align: ["left", "middle"], style: "$muted" } }
```

> 深色封面上所有文字颜色需与 `$ink` 对比（用浅色 `$text` 或 #ffffff 系）。

### 2. toc（目录，2×3 卡片网格）
配方：标题 + 6 个编号卡片（roundRect + 数字 + 标题）。卡片宽 260、高 110，间距 20。

```yaml
id: "02_toc"
pageType: "toc"
background: { color: "$bg" }
elements:
  - { elementType: "text", bounds: [40, 50, 400, 60], content: { text: "<p>目录</p>", align: ["left", "middle"], style: "$title" } }
  # 卡片 1–6：列 = [40, 320, 600],  行 = [150, 280, 410]
```

卡片通用块（以第 1 张为例）：
```yaml
  - { elementType: "shape", shapeName: "roundRect", bounds: [40, 150, 260, 110],
      fill: { type: "solid", color: "$paper" }, border: { style: "solid", color: "$line", width: 1 } }
  - { elementType: "text", bounds: [60, 165, 80, 40], content: { text: "<p>01</p>", align: ["left", "middle"], style: "$stat" } }
  - { elementType: "text", bounds: [60, 205, 220, 40], content: { text: "<p>市场概况</p>", align: ["left", "middle"], style: "$h3" } }
```

### 3. section-divider（章节页）
配方：大编号 + 章节标题 + 一句引言 + 底部装饰线。

```yaml
id: "03_section"
pageType: "section"
background: { color: "$bg" }
elements:
  - { elementType: "shape", shapeName: "roundRect", bounds: [40, 120, 420, 260],
      fill: { type: "solid", color: "$primary" } }
  - { elementType: "text", bounds: [60, 150, 380, 200],
      content: { text: "<p style='text-align:center'>02</p>", align: ["center", "middle"],
                 fontSize: 90, bold: true, color: "#ffffff" } }
  - { elementType: "text", bounds: [500, 170, 420, 70],
      content: { text: "<p>市场概况</p>", align: ["left", "middle"], style: "$title" } }
  - { elementType: "text", bounds: [500, 250, 420, 50],
      content: { text: "<p>行业规模、增长趋势与关键玩家</p>", align: ["left", "middle"], style: "$h3" } }
```

---

## 二、文本内容页

### 4. bullets（要点列表）
配方：标题 + 卡片式条目（左侧色条 + 要点文字）。

```yaml
id: "04_bullets"
pageType: "content"
background: { color: "$bg" }
elements:
  - { elementType: "text", bounds: [40, 50, 600, 60], content: { text: "<p>核心要点</p>", align: ["left", "middle"], style: "$title" } }
  - { elementType: "text", bounds: [100, 150, 780, 300],
      content: { text: "<p><strong>要点一：</strong>一句话说明…</p><p style='margin-top:18px'><strong>要点二：</strong>一句话说明…</p><p style='margin-top:18px'><strong>要点三：</strong>一句话说明…</p>",
                 align: ["left", "top"], style: "$body" } }
```

### 5. two-column（双栏对比：概念 + 示例）
```yaml
id: "05_twocol"
pageType: "content"
background: { color: "$bg" }
elements:
  - { elementType: "text", bounds: [40, 50, 600, 60], content: { text: "<p>两种视角</p>", align: ["left", "middle"], style: "$title" } }
  - { elementType: "shape", shapeName: "roundRect", bounds: [40, 140, 420, 330],
      fill: { type: "solid", color: "$paper" }, border: { style: "solid", color: "$line", width: 1 } }
  - { elementType: "text", bounds: [65, 160, 380, 50], content: { text: "<p>概念</p>", align: ["left", "middle"], style: "$h3" } }
  - { elementType: "text", bounds: [65, 215, 380, 230], content: { text: "<p>左侧描述…</p>", align: ["left", "top"], style: "$body" } }
  - { elementType: "shape", shapeName: "roundRect", bounds: [500, 140, 420, 330],
      fill: { type: "solid", color: "$ink" } }
  - { elementType: "text", bounds: [525, 160, 380, 50], content: { text: "<p>示例</p>", align: ["left", "middle"], style: "$h3" } }
  - { elementType: "text", bounds: [525, 215, 380, 230], content: { text: "<p>右侧示例…</p>", align: ["left", "top"], style: "$muted" } }
```

### 6. three-column（三栏图标 + 标题 + 描述）
三栏 x 坐标：40 / 340 / 640，统一宽 280。每栏：圆图标（oval 80×80）+ h3 + body。

### 7. big-quote（大字引述）
```yaml
id: "07_quote"
pageType: "content"
background: { color: "$ink" }
elements:
  - { elementType: "text", bounds: [120, 150, 720, 180],
      content: { text: "<p style='text-align:center'><em>“伟大的产品来自对需求的深刻理解。”</em></p>",
                 align: ["center", "middle"], fontSize: 32, bold: false, color: "#f5f0e8" } }
  - { elementType: "text", bounds: [120, 350, 720, 40],
      content: { text: "<p style='text-align:center'>—— 产品总监</p>", align: ["center", "middle"], style: "$muted" } }
```

---

## 三、数据内容页

### 8. stat-highlight（大数字）
```yaml
  - { elementType: "text", bounds: [40, 140, 880, 120],
      content: { text: "<p style='text-align:center'>3.2x</p>", align: ["center", "middle"],
                 fontSize: 96, bold: true, color: "$accent" } }
  - { elementType: "text", bounds: [40, 280, 880, 40],
      content: { text: "<p style='text-align:center'>转化率提升</p>", align: ["center", "middle"], style: "$h3" } }
```

### 9. kpi-grid（4 KPI 一行）
4 卡片 x：40 / 265 / 490 / 715，宽 205，高 150，y 150。每卡：数值（stat）+ 标签（muted）+ 涨跌（green/red）。

### 10. table（数据表）

```yaml
id: "10_table"
pageType: "content"
background: { color: "$bg" }
elements:
  - { elementType: "text", bounds: [40, 50, 600, 60], content: { text: "<p>核心指标</p>", align: ["left", "middle"], style: "$title" } }
  - elementType: "table"
    bounds: [40, 140, 880, 320]
    rows:
      - [{ text: "指标" }, { text: "2025" }, { text: "2026" }, { text: "同比" }]
      - [{ text: "营收" }, { text: "1.2亿" }, { text: "1.8亿" }, { text: "+50%" }]
      - [{ text: "客户数" }, { text: "320" }, { text: "610" }, { text: "+91%" }]
      - [{ text: "NPS" }, { text: "38" }, { text: "52" }, { text: "+14" }]
    columnWidths: [0.4, 0.2, 0.2, 0.2]
```

### 11. chart-bar（柱状图）
```yaml
id: "11_chart"
pageType: "content"
background: { color: "$bg" }
elements:
  - { elementType: "text", bounds: [40, 50, 600, 60], content: { text: "<p>营收趋势</p>", align: ["left", "middle"], style: "$title" } }
  - elementType: "chart"
    bounds: [60, 140, 600, 340]
    data:
      rows:
        - ["Q1", 0, 32]
        - ["Q2", 0, 45]
        - ["Q3", 0, 58]
        - ["Q4", 0, 72]
```

> chart 的 rows 格式：`[分类, 占位, 数值]`（第二列为占位，编译器取第 1、3 列）。

---

## 四、代码 / 终端页

### 12. code（代码块）
配方：深色代码容器（roundRect $ink）+ 浅色等宽文字 + 行号感。

```yaml
id: "12_code"
pageType: "content"
background: { color: "$bg" }
elements:
  - { elementType: "text", bounds: [40, 50, 600, 60], content: { text: "<p>核心实现</p>", align: ["left", "middle"], style: "$title" } }
  - { elementType: "shape", shapeName: "roundRect", bounds: [40, 140, 880, 340],
      fill: { type: "solid", color: "$ink" } }
  - { elementType: "text", bounds: [70, 170, 820, 280],
      content: { text: "<p style='color:#9ecbff'>def suggest(query):</p><p style='color:#d4d4d4;margin-top:6px'>    return rank(candidates(query))</p>",
                 align: ["left", "top"], fontSize: 14, color: "#e8e8e8" } }
```

### 13. terminal（终端窗口）
深色容器 + 顶部三色圆点（3 个 oval）+ 命令行文字。

---

## 五、图示 / 流程图页

### 14. process-steps（4 步流程）
4 个编号圆 + 箭头连线（line + arrow）+ 步骤卡片。

```yaml
  # 步骤圆点 x: 120, 320, 520, 720（oval 64×64）
  # 箭头：line bounds [x1, y+32, len, 2], arrow: [0, "arrow"]
```

### 15. flow-diagram（流水线 5 节点）
标题 + 5 个等宽方块（w 150）+ 4 条箭头。

### 16. timeline（横向时间线）
基线 line（y 300，横贯 880）+ 5 个圆点 + 上下交替文本。

### 17. roadmap（NOW / NEXT / LATER / VISION 四列）
4 列卡片，x: 40/270/500/730，宽 200，标题 + 2–3 条要点。

### 18. gantt（甘特图）
用 table 模拟或一排 roundRect 色块 + 轴标签（12 周）。

### 19. comparison（前后对比两栏）
类似 two-column，但两栏都承载"Before/After"式对比卡片。

### 20. pros-cons（利弊两卡）
左卡（green 色系标题"优点"）+ 右卡（red 系标题"缺点"）。

### 21. arch-diagram（架构图 3 层）
上层 1 宽条 + 中层 3 块 + 下层 1 宽条，箭头连接。

---

## 六、视觉内容页

### 22. image-hero（全幅图片 + 覆盖标题）
图片铺满（image）+ 深色透明遮罩（rect + #00000080）+ 底部标题文字。注意：编译器 alpha 走 `#RRGGBBAA`。

### 23. image-grid（图片栅格）
4–7 个 image 元素 + 说明文字。

---

## 七、收尾页

### 24. cta（行动号召）
深色底（$ink）+ 居中大字 + 强调按钮感色块。

### 25. thanks（致谢）
居中"谢谢"/"Thanks"大字 + 一句感谢语 + 联系方式。

---

## 布局选择纪律（来自 ppt-orchestra）

- 每页只选 **1 种**布局类型；相邻页不要重复
- 数据页必须带「结论一句话」和「数据来源」；图表 + 1–3 条 takeaway
- 对比页：Before/After、优缺点、选项对比，左右结构
- 图片页：hero 大图或栅格，图注 10–12pt muted
- 收尾必须有 CTA/下一步/联系信息，不要光秃秃的"谢谢"
- 正文不居中、不是纯文字页——至少加色块/图标/装饰线
- 正文不居中、不是纯文字页——至少加色块/图标/装饰线

---

## JSON 布局模板库（可直接粘贴 elements）

JSON 路径有开箱即用的布局模板：`templates/json/layout-templates.json`（12 个）。

覆盖：cover / toc / section / bullets / two_column / three_column / kpi_grid /
timeline / comparison / chart_page / image_hero / thanks。

用法：把模板的 `elements` 数组复制进 slide.elements，把 `{title}`、`{items}`、
`{rows}` 等占位符替换为实际内容即可。占位符式模板让 AI 或脚本一键套版，
详见 `references/json-engine.md`。

<!--
  (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
  https://github.com/linhut/ppt-studio
  Licensed under the MIT License. See the LICENSE file for details.
-->
# json-engine — JSON 数据驱动引擎参考

> 借鉴 [harness-anything](https://github.com/yb2460/harness-anything)（MIT）的
> 「JSON 数据驱动 + 元素类型路由器」设计，用 python-pptx 编译为**原生 PPTX**
> （跨平台，无需 WPS/Office）。

与 PPTD 引擎（ \`scripts/export_pptx.py\` ）二选一：

| 维度 | PPTD 引擎 | JSON 引擎（本页） |
|---|---|---|
| 数据 | deck.deck.pptd 清单 + pages/*.page | 单个 deck.json |
| 元素 | shape/line/text/table/chart/image | 上述全部 + **复合组件** |
| 复合组件 | 无 | 卡片网格 / 统计卡 / 目录列表 / 总结条 / 大数字 |
| 背景 | 必须内联 #HEX（inline_bg.py 辅助） | 可直接用 **$theme 变量**与**渐变** |
| 组件 | 55173 编辑器联动 | 纯命令行（Agent 友好） |
| 适用 | 需要编辑器联动的精细 deck | 数据驱动、批量、Agent 自动生成 |

## 数据格式

\`\`\`json
{
  "canvas": {"w": 960, "h": 540},
  "theme": {
    "colors": {"$primary": "#03045e", ...},
    "textStyles": {}
  },
  "slides": [
    {"id": "01", "title": "封面", "background": {"color": "$bg"},
     "elements": [ ... ]}
  ]
}
\`\`\`

- \`canvas\`：画布 pt（默认 960×540 = 16:9）
- \`theme.colors\`：\`$xxx\` 变量，任何颜色字段可用（可省略，用内置默认）
- \`slides[].background\`：省略 = \`$bg\` 纯色；\`{"type":"gradient","stops":[{color:$bg},{color:$accent}],"angle":115}\` 渐变
- \`slides[].elements\`：元素列表，见下

## 元素类型（路由器）

| type | 说明 | 关键字段 |
|---|---|---|
| \`text\` | 文本框 | x,y,w,h; text（支持 \\n 多行）; role(title/body/caption); fontSize/fs; bold; color; align(left/center/right); valign(top/middle/bottom); font; lineHeight |
| \`shape\` | 形状 | shapeName(rect/roundRect/oval/donut); fill.color; border.color/border.width |
| \`rect\` / \`line\` / \`circle\` | 快捷形状 | 同 shape |
| \`image\` | 图片 | file（相对 deck.json 的路径或绝对路径） |
| \`table\` | 数据表格 | rows,cols; data[][]; header_color; th_fs/td_fs |
| \`chart\` | 原生柱状图 | data.rows = [[类别, 单位, 数值]]; color |
| \`cards_2x3\` | 2行×3列彩色卡片 | items[{title,desc,color?}]; gap; 顶部 5pt 色条 |
| \`cards_1x4_info\` | 4 列统计卡 | items[{num,label,color?,fs?}] |
| \`card_list_wide\` | 编号列表（目录/要点/时间轴） | items[{num,title,sub}]; start_y; item_h; color_a/color_b 交替 |
| \`tagline_bar\` | 底部品牌色总结条 | text; x,y,w,h（默认贴底） |
| \`num_big\` | 大数字+标签（三段式 40/10/50） | num; label; color; fs |

## 设计预设（DesignPreset）

\`scripts/presets.py\`：colors + fonts + spacing + rules 四元组结构。

- 4 大预设：\`academic\` 学术 / \`consultant\` 咨询 / \`business\` 商务 / \`tech\` 科技
- 18 套配色（templates/themes/all-themes.yaml）也可直接作为预设引用，
  例如 \`--preset 15-pure-tech-blue\`

\`\`\`bash
py -3 scripts/presets.py list          # 列出全部预设
py -3 scripts/presets.py show academic # 查看预设细节
\`\`\`

## 质量审查（5 维度）

\`scripts/quality_check.py\`：逐页自动打分 + 硬约束校验 + 人工复核清单。

| 维度 | 检查内容 | 阈值 |
|---|---|---|
| visual | 字体层级 / 颜色数 / 视觉占比 / 一页一主题 | 70 |
| hard | 元素越界 / 标题-内容间距 ≥24pt / 标题字数 | 60 |
| pedagogy | 叙事弧（人工复核） | 75 |
| proofreading | 拼写术语溢出（人工复核） | 80 |
| parity | PPTX/PDF 一致性（人工复核） | 85 |
| substance | 数据准确性（人工复核） | 90 |

\`\`\`bash
py -3 scripts/quality_check.py my-deck/deck.json --preset academic
py -3 scripts/quality_check.py my-deck/deck.json --preset tech --json
\`\`\`

## 布局模板库

\`templates/json/layout-templates.json\`：12 个布局模板
（cover/toc/section/bullets/two_column/three_column/kpi_grid/timeline/
comparison/chart_page/image_hero/thanks），元素里用 \`{占位符}\` 表示待填内容。
把模板 elements 复制进 slide.elements 并替换占位符即可。

## 硬约束（自动检查）

- 元素不出画布：x+w ≤ 960，y+h ≤ 540
- 标题-内容间距 ≥ 24pt
- 标题 ≤ 12 字（tech 16 字），正文页标题 36-44pt
- 每页中文 ≤ 80 字 + ≥2 种视觉元素（图表/表格/卡片/图片/形状）
- 一页一主题（≤1 个 role=title）

## 编译

\`\`\`bash
# 默认：淡入淡出过渡
py -3 scripts/json2pptx.py my-deck/deck.json -o my-deck/deck.pptx
# 关闭过渡 / 严格模式（警告即失败）
py -3 scripts/json2pptx.py my-deck/deck.json --no-fade --strict
\`\`\`

## 完整示例

\`examples/json-demo/\`：4 页 deck.json（封面/KPI 卡/路线图/致谢），
覆盖复合组件与淡入淡出，直接 \`py -3 scripts/json2pptx.py examples/json-demo/deck.json\` 即可编译。

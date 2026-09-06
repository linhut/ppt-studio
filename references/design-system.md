<!--
  (c) 2026 Jose AI (https://www.linhut.cn)
  https://github.com/linhut/ppt-studio
  Licensed under the MIT License. See the LICENSE file for details.
-->
# design-system — 多源融合设计规范（综合评判）

> 本文档回答一个问题：多个优秀 skill 各自的专长是什么，如何取长补短合成一套
> 既专业又可编译的 PPT 设计系统。数值来自各家源码的真实值，不是自创。
>
> 参考源（全部本地已克隆 / 可联网复核）：
> 1. **html-ppt**（lewislulu）— 视觉层级 / 31 布局 / token 设计系统
> 2. **Color Scheme - PPT**（MiniMax）— 18 套成品配色 / 字体规范
> 3. **dsh-np-ppt** — PPTD DSL + Python-PPTX 高保真编译内核
> 4. **harness-anything**（yb2460）— 品牌色 60-70% 视觉占比 / 图表铁律 / 真实软件 harness
> 5. **GordenPPTSkill**（GordenSun）— 容量精算（防溢出）/ 设计纪律（≤1主色+1强调色）/ 非破坏性编辑

## 一、综合评判：各源专长

| 维度 | html-ppt | Color Scheme - PPT (MiniMax) | dsh-np-ppt | harness-anything | GordenPPTSkill |
|---|---|---|---|---|---|
| **配色** | 36 主题 token | ⭐ **18 套成品配色** + Agent Design System 色阶 | 9 大风格气质 | ⭐ **品牌色克制哲学**（1 主色 + 黑白灰，60-70% 视觉占比）| ⭐ 模式 C：≤1 主色 + 1 强调色，纯白底+深色字 |
| **字体** | 字号层级 h1=72…px | ⭐ 中雅黑/英 Arial | 雅黑锁定 | ⭐ 大字号体系：H1 40-48 / Body 18-20 / Caption 14-15 | ⭐ 梯度 48-54 封面 / 32-40 章节 / 18-22 正文 / 12-14 脚注 |
| **布局** | ⭐ 31 类布局 | 无 | 960×540pt 坐标 | ⭐ 精确坐标（标题 y=14 / 内容 y≥90 / 底条 y=498）| 版式按内容类型（封面居中/3列/对比2列/时间线/大数字）|
| **间距/圆角** | ⭐ base.css 真实 token | 风格配方英寸值 | 页边距≥40pt | 顶/底品牌色条 5pt | 留白 ≥1.5cm、不堆装饰 |
| **生成** | HTML（非 PPTX）| PptxGenJS | ⭐ PPTD + python-pptx 内核 | ⭐ COM 直绘真实软件 | 非破坏性编辑（只改文本）|
| **QA** | 演讲逐字稿 | 5 页面类型检查 | 55173 一致性校验 | 5 维度审查 | ⭐ **容量精算**（CJK=1.0/ASCII=0.5 视觉宽度 + 20% 容差，防溢出遮挡）|

### 结论：取长补短矩阵

| 要做的事 | 采用谁 | 为什么 |
|---|---|---|
| 选配色 | **MiniMax 18 套成品** + **harness/Gorden 克制哲学** | 成品色板保证专业；但**每页只用 1 主色 + 1 强调色 + 黑白灰**，禁止 5-6 色混用 |
| 定字号层级 | **harness/Gorden 大字号体系**（对齐后）| 正文 16-18、标题 40、卡片 18、弱化 13 —— 解决"字体大小不协调" |
| 定页面结构 | **html-ppt 31 布局** + **harness 坐标系统** | 布局内容组织 + 精确坐标（标题 y=14 / 内容 y≥84 / 底条 y=498）|
| 防溢出 | **GordenPPTSkill 容量精算** | 视觉宽度单位（CJK 1.0 / ASCII 0.5）+ 20% 容差 → 编译前预警遮挡 |
| 生成引擎 | **np-ppt**（PPTD + python-pptx）| 高保真原生 PPTX 路线 |
| 复合组件/路由 | **harness-anything** | JSON 数据驱动 + 元素路由，Agent 友好 |
| 字体 | MiniMax 规范 + np-ppt 锁定 | 中雅黑/英 Arial，正文不加粗 |

## 二、融合字号层级（html-ppt px → 960×540pt 画布）

换算依据：html-ppt 全屏 1920px ≈ PPT 960pt 画布，视觉权重换算系数约 0.75。

> 字号已对齐 harness-anything / GordenPPTSkill 大字号体系（原 html-ppt px 映射偏小，
> 导致"字体大小不协调"。以下为定稿值，json2pptx.py FS_* 常量同步）。

| 层级 | 定稿值 | 用途 |
|---|---|---|
| 封面标题 | **54pt** | 封面主标题（harness H1 40-48 上限档）|
| 内容页标题 | **40pt** | 页面标题（harness H1 40-48）|
| 小节标题 | **24pt** | H2（harness H2 24-26）|
| 卡片标题 | **18pt** | 卡片内标题（harness H3 20-22 → 卡片取 18）|
| 正文 | **16pt** | 正文要点（harness Body 18-20 → 画布 960×540 取 16）|
| 弱化/注释 | **13pt** | 说明文字（harness Caption 14-15 → 取 13）|
| 眉题 | **12pt** | 页眉眉题 |
| 引导词 | **13pt** | kicker（品牌色）|
| 大数字 | **44pt** | 统计数字（粗体）|
| 卡片描述 | **14pt** | 卡片说明文字 |
| 表头 | **13pt** | 表格表头 |
| 单元格 | **12pt** | 表格数据 |
| 底部总结条 | **15pt** | tagline_bar（白字）|

## 三、融合间距与圆角（html-ppt base.css → pt）

| token | 原值 | PPT 映射 | 用途 |
|---|---|---|---|
| slide padding | 72px 96px | 上下 40pt / 左右 48pt | 页边距（≥40pt 规范）|
| stack gap | 14px | 14pt | 块内元素间距 |
| row/grid gap | 24px | 24pt | 卡片/列间距 |
| mt-s / mt-m / mt-l | 8/18/32px | 8/14/24pt | 上边距档位 |
| mb-s / mb-m / mb-l | 8/18/32px | 8/14/24pt | 下边距档位 |
| card padding | 26px 28px | 16pt 18pt | 卡片内边距 |
| card border | 1px | 1pt（浅色边框）| 卡片描边 |
| radius / sm / lg | 18/12/26px | 12/8/18pt | 卡片圆角 |
| divider-accent | 3px×72px | 3pt×48pt | 标题下装饰线 |
| card-accent 顶条 | 3px | 3pt | 卡片顶部色条 |

## 四、融合配色映射（MiniMax 18 套 → PPTD/JSON 变量）

18 套成品配色的 5 色与「建议」列一起映射到主题变量。规则：

1. 场景建议里的「标题色」→ \`$primary\`；「背景色」→ \`$bg\`
2. 剩余色按明度分配：最浅 → \`$paper\`/\`$line\`；最深 → \`$ink\`/\`$text\`
3. 每套色板的中间高亮色 → \`$accent\`（数据焦点）
4. 绿系/正向 → \`$green\`；剩余装饰 → \`$brass\`

完整 18 套映射见 \`references/themes.md\`（已按本规则维护）与
\`templates/themes/all-themes.yaml\`（机器可读）。

### 深色主题特例（第 9 套「科技与夜景」等）

- \`$bg\` 用最深色，\`$text\` 用最浅色，\`$primary\` 用高亮金——保证对比度
- 深色封面 + 浅色正文页不混用；一 deck 一主题

## 五、防遮挡与排版纪律（本规范强制）

1. **标题区**：内容页标题 y=14~54（高 40pt），标题下装饰线 y=60；第一个内容元素必须 y≥84（与标题底 ≥24pt 间距）
2. **卡片文字适配**：卡片标题 16pt、正文 12-13pt；desc 高度按 1.4 行高 × 行数预计算，超出即缩短文案（≤2 行）
3. **表格**：表头 12pt 白字深底，单元格 11-12pt；行高 = h/rows，最小 24pt；文字不换行（WordWrap=False）
4. **大数字组件**：数字 40-44pt 占 40%、标签 12pt 占 50%，中间 10% 间隙——三段式不重叠
5. **颜色纪律**：一页 ≤5 色（主题变量 + 黑白）；正文只用 \`$text\`/\`$muted\`，强调才用 \`$accent\`
6. **视觉占比**：每页 ≥2 种视觉元素（卡片/图表/表格/图片/色块/装饰线），纯文字页不合格
7. **元素边界**：x+w≤960、y+h≤540，任何元素不得越界（审查硬约束）

## 六、页面结构模板（html-ppt 布局 → JSON/PPTD）

| html-ppt 模板 | 结构 | 对应元素 |
|---|---|---|
| cover | deck-header(kicker) + h1 + lede + pills 行 + footer | text(kicker/title/lede) + pill 圆角标签（shape+text）|
| toc | eyebrow + h2 + grid g2 卡片（编号 h3 + h4 + dim）| cards_2x3（2 列卡片）|
| kpi-grid | kicker + h2 + grid g4 大数字卡（eyebrow + 44pt 数字 + dim 增量）| cards_1x4_info |
| bullets | kicker + h2 + lede + card-accent 列表（h4 + dim）| card_list_wide / 文本卡 |
| two-column | h2 + grid g2 卡片（h3 + dim + ul）| 双卡布局 |
| roadmap | kicker + h2 + 列式时间线（tag pill + h4 + ul）| 时间轴卡片 |

## 七、三源合成后的主题示例（「纯净科技蓝」#15）

\`\`\`yaml
theme:
  colors:
    $bg: "#caf0f8"      # 最浅蓝（背景，建议"从深海到天空"浅端）
    $paper: "#ffffff"   # 卡片表面
    $primary: "#03045e" # 深海蓝（标题）
    $accent: "#0077b6"  # 中蓝（数据焦点）
    $text: "#001f3d"    # 深蓝黑（正文）
    $muted: "#5c7a99"   # 灰蓝（弱化）
    $line: "#b8dce8"    # 浅蓝边
    $ink: "#03045e"     # 深色块
    $green: "#00b4d8"   # 亮青（正数据）
    $brass: "#90e0ef"   # 浅青（装饰）
  textStyles:
    $title: { fontSize: 36, bold: true, color: "$primary" }
    $h2:    { fontSize: 24, bold: true, color: "$text" }
    $body:  { fontSize: 14, bold: false, color: "$text", lineHeight: 1.4 }
    $muted: { fontSize: 12, bold: false, color: "$muted" }
    $kicker:{ fontSize: 12, bold: true, color: "$accent", letterSpacing: 1 }
\`\`\`
## 九、容量精算（借鉴 GordenPPTSkill，防溢出遮挡）

> 用户最常抱怨"错位遮挡"。根因之一是文字超过文本框容量。质量审查现已内置容量精算，
> 用"视觉宽度单位"估算每个文本框能放多少字，编译前预警。

- **视觉宽度单位 vw**：CJK/全角字符 = 1.0；空格 = 0.35；ASCII 拉丁/数字 = 0.5；其他 = 0.8。
- 每行 vw 容量 = 可用宽（pt）÷ 字号（pt）；可用宽 = 文本框宽 − 8pt（左右内边距合计）。
- 最大行数 = 可用高（pt）÷（字号 × 行高 1.0）；不换行 = 1 行。
- 容量预算 = 每行容量 × 最大行数 × **1.2**（20% 容差，防误报）。
- 实际内容：按 <br/> 与换行分行，每行 vw ÷ 每行容量向上取整得所需行数，求和。
- 所需行数 > 最大行数 → 警告「容量溢出」，扣分。

**使用纪律**：溢出警告是提示，不是硬阻断。优先用更精炼的措辞重写；实在放不下时
宁可轻微超框也不要用省略号截断半句话。

## 十、设计纪律（借鉴 GordenPPTSkill 模式 C + harness 品牌哲学）

**颜色**：
- 每页 ≤ 1 主色 + 1 强调色 + 黑白灰（禁止 5-6 色混用）。
- 纯白底 + 深色文字，或深色底 + 纯白文字；禁止浅底浅字（对比度不足）。
- 品牌色用于标题/色条/表头/强调色块，视觉占比 60-70%（靠留白与色块，不靠堆色）。

**版式**：
- 大量留白：元素距画布边 ≥ 1.5cm（≈ 43pt），页边距 ≥ 40pt。
- 一页最多 1 标题 + 1 副标 + 3-4 项核心内容；信息点超过 4 个改纵向列表。
- 不堆砌装饰：避免无意义的边框、阴影、渐变、纹理、icon 集群；不做"卡片套卡片"。
- 版式按内容类型：封面全屏居中 / 目录竖排 / 3 列要点 / 对比 2 列+分隔线 /
  时间线横向轴+节点 / 数据 1 个大数字占 50% + 解释 / 结束致谢居中。
- 标题 y=14、内容区 y≥84（标题下间距 ≥24pt）、底部总结条 y=498。

**字体**：
- 中文微软雅黑 / 思源黑体 CN；英文 Inter / Helvetica / Arial。
- 标题加粗，正文常规，不用斜体；不用装饰字体（华文琥珀/隶书/楷体）。
- 同级标题字号一致（type scale 唯一性）。

---

## 十一、布局样式库（Layout Catalog）

> 完整布局索引（按页面角色选样式、复合组件速查、选择指南）见
> **[layout-catalog.md](layout-catalog.md)**。本节只列引擎已内置的复合组件
> 与其设计纪律约束。

**复合组件（15 个，一行 JSON 排版一整页）**：

| type | 布局 | 纪律 |
|---|---|---|
| `section_divider` | 分隔页（center / leftaccent） | 一页只放数字+标题+一句 intro |
| `comparison_2col` | 双卡对比 | 每侧 ≤3 点，色条区分主/强调 |
| `pros_cons` | 利弊两卡 | 各 ≤4 条；✓ 绿 / ✗ 蓝 |
| `timeline_h` | 横向时间线 | ≤5 节点，每节点 ≤12 字说明 |
| `process_steps` | 流程步骤 | ≤5 步，箭头连接 |
| `kpi_row` | KPI 横排 | ≤4 个，带增量自动着色 |
| `big_quote` | 大引语 | 引文 ≤40 字 |
| `roadmap_4col` | NOW/NEXT/LATER/VISION | 每列 ≤3 项 |
| `checklist` | 清单 | ≤6 项 |
| `cover_asym` | 非对称封面 | 大色块 ≤1/3 画布宽 |

**约束**：相邻两页不重复同一种布局；新组件内部形状不计入"颜色数"检查
（颜色纪律仍以每页文本/色块为准）；容量精算适用于所有 text 元素。

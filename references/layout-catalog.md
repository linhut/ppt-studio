<!--
  (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
  https://github.com/linhut/ppt-studio
  Licensed under the MIT License. See the LICENSE file for details.
-->
# 布局样式库 Layout Catalog

> 本目录是 ppt-studio 的"排版样式词典"。生成 deck 时按内容类型挑选布局，
> 优先使用下方**复合组件**（一行 JSON 即可排版整页），复杂页面再用
> 基础元素（rect / text / table / chart）手工组合。
>
> 综合来源（详情见 design-system.md 五源矩阵）：
> - html-ppt：31 类布局（toc/section-divider/stat/kpi/timeline/roadmap/comparison/pros-cons/checklist/image-hero…）
> - MiniMax Color Scheme：页面生成器（cover/toc/section-divider/summary/content 六子类型）
> - GordenPPTSkill：模式 C 版式表（封面/目录/分隔/引用/3列/5+列表/对比/时间线/数据大数字/结尾）
> - harness-anything：元素路由器（HTML 布局 ≤8 个元素的克制哲学）
> - academic-pptx-skill：学术演示 11 模式（行动标题/结果图左文右/面包屑/参考文献/附录）
> - deckary（300+ 咨询演示）：5 大设计原则（对齐/对比/留白 15-20%/一致性/层级）

## 一、复合组件速查（JSON 引擎，一行排版一整页）

| type | 布局 | 来源 | 关键参数 |
|---|---|---|---|
| `section_divider` | 分隔页：大数字+标题+说明 | MiniMax / Gorden | num, title, intro, variant="center|leftaccent" |
| `comparison_2col` | 左右对比双卡片 | html-ppt / MiniMax | left{title,points}, right{...}, left_color, right_color |
| `pros_cons` | 利/弊两卡（✓绿 / ✗蓝） | html-ppt | pros[], cons[], pro_title, con_title |
| `timeline_h` | 横向时间线（轴线+节点圆） | html-ppt / Gorden | items[{time,title,desc}] |
| `process_steps` | 编号流程卡+箭头 | html-ppt / MiniMax | steps[{title,desc}] |
| `kpi_row` | KPI 横排（数字+标签+增量） | html-ppt kpi-grid | items[{num,label,delta}]（delta 含 +/- 自动着色） |
| `big_quote` | 大引语（左侧粗线） | html-ppt | quote, author |
| `roadmap_4col` | NOW/NEXT/LATER/VISION | html-ppt | columns[{phase,items}] |
| `checklist` | 清单（✓完成/□未做） | html-ppt | items[{text,done}] |
| `cover_asym` | 非对称封面（左文右色块） | MiniMax | kicker, title, subtitle, meta, block_side |
| `cards_2x3` | 2×3 卡片网格（顶部色条） | html-ppt toc | items[{title,desc}], card_color |
| `cards_1x4_info` | 4 列统计卡 | html-ppt kpi | items[{num,label}] |
| `card_list_wide` | 编号圆+标题+副标列表 | html-ppt / MiniMax | items[{num,title,sub}] |
| `tagline_bar` | 底部品牌色总结条 | harness | text, color |
| `num_big` | 大数字+标签（44pt） | Gorden | num, label, fs, color |
| `figure_text` | 结果页：图左文右+底部来源 | academic-pptx | figure_ratio, figure_title, figure_placeholder, heading, points[{lead,text}], source, point_color |
| `breadcrumb` | 顶部面包屑导航条（章节+当前高亮） | academic-pptx | sections[], active, height, color, bar_bg |
| `references` | 参考文献小字列表 | academic-pptx | items[] |

## 二、按页面角色选布局（html-ppt 31 布局索引）

### 开场（Openers）
| 布局 | 组件/元素 | 适用 |
|---|---|---|
| 标准封面 | text(kicker/title/lede) + 顶底色条 | 通用汇报 |
| 非对称封面 | `cover_asym` | 产品发布、留白美学 |
| 中心封面 | text 居中 + 大标题 | 演讲、活动 |
| 目录-竖排编号 | `card_list_wide`（01-04） | 3-5 章 |
| 目录-卡片网格 | `cards_2x3` | 4-6 章内容型 |
| 目录-侧边栏 | 左窄色条 + 右标题列表 | 现代商务 |

### 分隔页（Section Divider）
| 布局 | 参数 | 适用 |
|---|---|---|
| 居中大数字 | `section_divider` variant="center" | 极简现代 |
| 左侧色块 | variant="leftaccent" | 商务结构化 |
| 分屏背景 | rect 左深右浅 + 文字 | 高对比过渡 |

### 内容页（Content）
| 布局 | 组件/元素 | 适用 |
|---|---|---|
| 要点列表 | text bullets（≤6 条） | 通用 |
| 双栏概念+示例 | 左 text 右 rect/text | 解释说明 |
| 三等分图标柱 | 3×rect + 图标 + 文字 | 三支柱 |
| 流程步骤 | `process_steps` | 工序/管线 |
| 混排图文 | 左文右图（image） | 场景演示 |
| 时间线 | `timeline_h` | 里程碑/历史 |
| 代码/终端 | 深底 rect + 等宽字 | 技术分享 |
| 长汇报导航 | `breadcrumb`（顶部章节条，内容 y≥40） | >15 分钟长 deck 每页加 |

### 数据可视化（Numbers & Data）
| 布局 | 组件/元素 | 适用 |
|---|---|---|
| 单一大数字 | `num_big`（44pt） | 核心指标强调 |
| KPI 横排 | `kpi_row`（带增量） | 仪表盘 |
| 折线/柱状/饼/雷达 | chart + 数据源标注 | 趋势/占比/对比 |
| 表格 | table（表头品牌色） | 明细数据 |
| 结果页（一图一结论） | `figure_text`（图左文右 + 来源） | 研究/数据汇报：图上标注关键数字 |

### 计划类（Plans）
| 布局 | 组件/元素 | 适用 |
|---|---|---|
| 路线图四列 | `roadmap_4col` | NOW/NEXT/LATER |
| 甘特条 | rect 横向堆叠 | 12 周排期 |
| 里程碑 | `timeline_h` | 项目节点 |

### 对比类（Comparison）
| 布局 | 组件/元素 | 适用 |
|---|---|---|
| 双卡对比 | `comparison_2col` | 方案 A/B |
| 利弊两卡 | `pros_cons` | 决策权衡 |
| 前后对比 | `comparison_2col` 左"之前"右"之后" | 改造效果 |
| 清单 | `checklist` | 检查项 |

### 结尾（Closers）
| 布局 | 组件/元素 | 适用 |
|---|---|---|
| 要点收束 | 3-5 ✓ 要点列表 | 汇报总结 |
| CTA/下一步 | 编号行动项 + 联系信息 | 提案/销售 |
| 致谢 | 居中大标题 + 联系信息 | 演讲 |
| 大引语 | `big_quote` | 收束升华 |
| 参考文献 | `references`（小字列表） | 学术/研究 deck 必收尾 |

## 三、选择指南（按内容快速定位）

- 讲**结构** → 目录（竖排/网格）→ 分隔页 → 要点列表
- 讲**数字** → num_big（1 个大数）→ kpi_row（4 个）→ chart
- 讲**流程** → process_steps → timeline_h
- 讲**对比** → comparison_2col → pros_cons
- 讲**计划** → roadmap_4col → timeline_h
- 讲**结论** → big_quote → 要点收束 → 致谢
- 讲**证据/数据结果** → figure_text（一图一结论）→ 页内引来源
- 长 deck 导航 → breadcrumb 每页顶部；研究类收尾 → references 参考文献页

## 四、排版纪律（与 design-system.md 第十章联动）

1. **每页 ≤1 主色 + 1 强调色 + 黑白灰**；白底深字优先
2. **每页 ≤4 个信息点**；一页一个主题、一个视觉锚点
3. **字号梯度**：封面 54 / 标题 40 / H3 24 / 卡片标题 18 / 正文 16 / 弱化 13
4. **留白 ≥1.5cm**；不卡片套卡片、不堆装饰
5. **容量自测**：编译前跑 quality_check，>80 字/页文本需拆分（容量精算见 design-system 第九章）
6. 复合组件优先于手工拼元素——减少错位遮挡风险

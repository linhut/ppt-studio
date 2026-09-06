<!--
  (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
  https://github.com/linhut/ppt-studio
  Licensed under the MIT License. See the LICENSE file for details.
-->
# 动画思路（html-ppt 动画 → PPTX 实现）

> 借鉴 html-ppt 的 27 种 CSS 动画 + 20 种 canvas 特效的**分类思路**，
> 翻译为 PPTX 可行的动画/过渡方案。html-ppt 的动画是网页特效，PPTX 是静态文件，
> 因此分两层：**过渡动画**（编译器原生支持）与**元素入场动画**（XML 注入，高级可选）+ **静态动感设计**（推荐）。

## 一、过渡动画（默认支持）

编译器 `--transition fade` 为每页写入淡入淡出（Fade，fast）。这是 PPTX 最稳妥的转场：

```bash
py -3 scripts/export_pptx.py deck/deck.pptd --output deck/deck.pptx --force          # fade（默认）
py -3 scripts/export_pptx.py deck/deck.pptd --output deck/deck.pptx --transition none --force   # 无过渡
```

## 二、元素入场动画（高级：XML 注入）

python-pptx 未暴露动画 API，但可注入时序 XML（`<p:timing>`）。典型做法是给
目标 shape 的 `spPr` 相邻插入动画效果。**仅当用户明确要求逐元素动画时使用**，
并做好 PowerPoint 兼容性验证：

```python
# 思路（完整实现见各项目验证）：
# 1. 给元素 shape 命名（slide.shapes 遍历按 index 定位）
# 2. 构造 <p:timing> 树：bldLst → bldP (spid=shape_id, grpId=0) → animEffect/anim
# 3. 插入 slide._element（在 transition 之前）
# 效果建议：fade（淡入）、fly（飞入）、zoom（缩放）
```

> 风险提示：PowerPoint 对 timing 树较挑剔，注入后务必用 PowerPoint/WPS 打开验证；
> 若验证成本高，优先用下面的「静态动感设计」替代。

## 三、静态动感设计（推荐，零风险）

把 html-ppt 的动画**语义**转成静态视觉语言，所有 PPTX 都能稳定呈现：

| html-ppt 动画 | 意图 | PPTD 静态等价实现 |
|---|---|---|
| fade-up / fade-in | 元素逐层浮现 | 元素分层：卡片从深到浅堆叠，关键数字放大强调 |
| slide-in / fly-in | 内容滑入 | 左右分栏：左侧文字、右侧色块，视觉重心偏移 |
| zoom-in | 焦点放大 | 大数字（stat 96pt）+ 小标签，大小对比 |
| stagger-list | 列表依次出现 | 列表项用色条序号 01/02/03 引导视线逐条阅读 |
| pulse / glow | 强调呼吸感 | 强调元素用 accent 亮色 + 深底对比（如金色标题） |
| typewriter | 逐字呈现 | 大标题拆行，短句独立成段（留白引导节奏） |
| particle / confetti / firework | 庆祝、惊喜 | 结尾页用多色圆点（oval 阵列）点缀 |
| knowledge-graph / neural-net | 关系网络 | arch-diagram 或 flow-diagram 节点+连线 |
| gradient-blob / aurora | 氛围感 | 封面用大色块（$primary/$brass 圆角块）营造层次 |
| counter-explosion | 数字增长 | 大数字 + 涨跌箭头（▲ +50% 用 $green） |

**设计节奏原则**（对应 html-ppt 的动画节奏）：
- 每页 1 个视觉焦点，其余元素弱化（muted）
- 相邻页切换节奏：数据页 → 图示页 → 文本页交替，避免连续 3 页同构
- 演讲者逐字稿与页面节奏对应：1 页 = 2–3 分钟 = 150–300 字

## 四、动画分类速查（html-ppt 原始目录，供灵感）

- **入场**：fade-up、fade-in、slide-up、slide-in、zoom-in、flip-in、blur-in
- **列表**：stagger-list、stagger-grid
- **强调**：pulse、glow、shake、bounce
- **文字**：typewriter、word-cascade、letter-explode
- **背景/氛围**：gradient-blob、aurora、starfield、matrix-rain、constellation
- **庆祝**：confetti-cannon、firework、particle-burst、sparkle-trail
- **关系**：knowledge-graph、neural-net、orbit-ring、galaxy-swirl
- **数据**：counter-explosion、data-stream、magnetic-field

> 在 PPTX 中，这些动画的"气质"通过配色、构图、字体层级来传达——这是 html-ppt 设计思路对静态 PPT 最有价值的迁移。

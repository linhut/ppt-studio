<!--
  (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
  https://github.com/linhut/ppt-studio
  Licensed under the MIT License. See the LICENSE file for details.
-->
# 主题设计系统：18 套配色 → PPTD theme 映射

> 融合：Color Scheme - PPT 的 18 套配色 + html-ppt 的 token 设计系统（背景/表面/文字/强调分层）+ dsh-np-ppt 的 theme 变量。
> 每套配色都是**开箱即用**的 PPTD theme 代码块，直接粘进 `deck.deck.pptd` 的 `theme.colors` 即可。

## 一、变量语义（html-ppt token → PPTD 变量）

| PPTD 变量 | 对应 html-ppt token | 用途 |
|---|---|---|
| `$bg` | `--bg` | 页面背景（纯色） |
| `$paper` | `--surface` | 卡片/浅色表面 |
| `$primary` | `--accent` | 主色：标题强调、品牌色 |
| `$accent` | `--accent-2` | 强调色：数据焦点、高亮 |
| `$text` | `--text-1` | 正文主色 |
| `$muted` | `--text-2/3` | 弱化文字、注释 |
| `$line` | `--border` | 线条/边框/分隔 |
| `$ink` | `--text-1`（深色块） | 表头/深色块填充 |
| `$green` | `--good` | 图表正色、正向数据 |
| `$brass` | `--accent-3` | 装饰辅助色 |

**推导规则**：5 色板按"明度 + 饱和度"分配——最浅→背景，最鲜→强调，最深→主色/文字。深色系主题（bg 暗）时 `$text` 用最浅色、`$bg` 用最深色。

## 二、18 套配色完整映射

### 1. 现代与健康（清新治愈）
```yaml
colors:
  $bg: "#edf6f9"    # 最浅蓝白
  $paper: "#ffffff"
  $primary: "#006d77"   # 深青
  $accent: "#e29578"    # 珊瑚
  $text: "#00363a"      # 深青黑
  $muted: "#6b8f94"
  $line: "#cfe0e3"
  $ink: "#006d77"
  $green: "#83c5be"
  $brass: "#ffddd2"
```

### 2. 商务与权威（严谨经典）
```yaml
colors:
  $bg: "#edf2f4"    # 冷灰白
  $paper: "#ffffff"
  $primary: "#2b2d42"   # 深蓝黑
  $accent: "#ef233c"    # 亮红（数据强调）
  $text: "#2b2d42"
  $muted: "#8d99ae"
  $line: "#d7dde3"
  $ink: "#2b2d42"
  $green: "#4a7c59"
  $brass: "#8d99ae"
```

### 3. 自然与户外（沉稳大地）
```yaml
colors:
  $bg: "#fefae0"    # 米色
  $paper: "#ffffff"
  $primary: "#283618"   # 深绿
  $accent: "#bc6c25"    # 赭石
  $text: "#283618"
  $muted: "#8a7d5f"
  $line: "#e8e0c4"
  $ink: "#283618"
  $green: "#606c38"
  $brass: "#dda15e"
```

### 4. 复古与学院（经典书卷）
```yaml
colors:
  $bg: "#fdf0d5"    # 羊皮纸
  $paper: "#ffffff"
  $primary: "#780000"   # 深红
  $accent: "#669bbc"    # 蓝
  $text: "#2c2c2c"
  $muted: "#8a8a8a"
  $line: "#e5d5b0"
  $ink: "#780000"
  $green: "#3a5a40"
  $brass: "#c1121f"
```

### 5. 柔美与创意（梦幻糖果）
```yaml
colors:
  $bg: "#fdf7ff"    # 极浅紫
  $paper: "#ffffff"
  $primary: "#b5838d"   # 玫瑰
  $accent: "#a2d2ff"    # 浅蓝
  $text: "#3a3a4a"
  $muted: "#9a8f9f"
  $line: "#eadcf0"
  $ink: "#6d597a"
  $green: "#9ec5ab"
  $brass: "#ffc8dd"
```

### 6. 波西米亚（温柔低饱和）
```yaml
colors:
  $bg: "#fefae0"
  $paper: "#ffffff"
  $primary: "#6b705c"   # 橄榄绿
  $accent: "#d4a373"    # 陶土
  $text: "#4a4a3a"
  $muted: "#9a9484"
  $line: "#eae2c8"
  $ink: "#6b705c"
  $green: "#a3b18a"
  $brass: "#ccd5ae"
```

### 7. 活力与科技（高能量运动）
```yaml
colors:
  $bg: "#ffffff"
  $paper: "#f2f8fb"
  $primary: "#023047"   # 深蓝
  $accent: "#fb8500"    # 亮橙
  $text: "#023047"
  $muted: "#6c7f8d"
  $line: "#dce8ef"
  $ink: "#023047"
  $green: "#219ebc"
  $brass: "#ffb703"
```

### 8. 匠心与手作（质朴咖啡）
```yaml
colors:
  $bg: "#ede0d4"    # 奶油棕
  $paper: "#ffffff"
  $primary: "#414833"   # 深咖
  $accent: "#7f5539"    # 咖啡
  $text: "#33302b"
  $muted: "#8f8578"
  $line: "#dccfc0"
  $ink: "#414833"
  $green: "#656d4a"
  $brass: "#a68a64"
```

### 9. 科技与夜景（深邃高亮，深色模式）
```yaml
colors:
  $bg: "#000814"    # 最深黑
  $paper: "#001d3d" # 深蓝表面
  $primary: "#ffc300"   # 亮金
  $accent: "#ffd60a"    # 亮黄
  $text: "#e8edf5"      # 浅色正文（深色模式）
  $muted: "#7a8aa5"
  $line: "#123052"
  $ink: "#003566"
  $green: "#4cc9f0"
  $brass: "#ffc300"
```

### 10. 教育与图表（清晰逻辑）
```yaml
colors:
  $bg: "#ffffff"
  $paper: "#f4f9f8"
  $primary: "#264653"   # 深青蓝
  $accent: "#e76f51"    # 陶红
  $text: "#264653"
  $muted: "#6f8a93"
  $line: "#dce5e8"
  $ink: "#264653"
  $green: "#2a9d8f"
  $brass: "#e9c46a"
```

### 11. 森林与环保（单色森系）
```yaml
colors:
  $bg: "#f5f4ee"
  $paper: "#ffffff"
  $primary: "#344e41"   # 深森绿
  $accent: "#588157"    # 松绿
  $text: "#344e41"
  $muted: "#8a9384"
  $line: "#dde2d6"
  $ink: "#344e41"
  $green: "#588157"
  $brass: "#a3b18a"
```

### 12. 优雅与时尚（低饱和莫兰迪）
```yaml
colors:
  $bg: "#f7f1ed"
  $paper: "#ffffff"
  $primary: "#4a5759"   # 灰绿
  $accent: "#b0c4b1"    # 鼠尾草
  $text: "#3d4445"
  $muted: "#8b8f8e"
  $line: "#e5ded8"
  $ink: "#4a5759"
  $green: "#9db89f"
  $brass: "#edafb8"
```

### 13. 艺术与美食（浓郁画报）
```yaml
colors:
  $bg: "#fdf6e8"
  $paper: "#ffffff"
  $primary: "#540b0e"   # 深酒红
  $accent: "#e09f3e"    # 金橙
  $text: "#3a2020"
  $muted: "#8c7a6a"
  $line: "#e8dcc0"
  $ink: "#540b0e"
  $green: "#335c67"
  $brass: "#9e2a2b"
```

### 14. 轻奢与神秘（冷艳紫调）
```yaml
colors:
  $bg: "#f2e9e4"    # 暖灰
  $paper: "#ffffff"
  $primary: "#22223b"   # 深紫黑
  $accent: "#9a8c98"    # 紫灰
  $text: "#22223b"
  $muted: "#8b8490"
  $line: "#e0d7d0"
  $ink: "#22223b"
  $green: "#4a4e69"
  $brass: "#c9ada7"
```

### 15. 纯净科技蓝（未来感）
```yaml
colors:
  $bg: "#f0f8fc"
  $paper: "#ffffff"
  $primary: "#03045e"   # 深海蓝
  $accent: "#00b4d8"    # 亮青
  $text: "#0a2472"
  $muted: "#6c93b8"
  $line: "#d4e8f2"
  $ink: "#03045e"
  $green: "#00afb9"
  $brass: "#90e0ef"
```

### 16. 海岸珊瑚（清爽夏日）
```yaml
colors:
  $bg: "#fdfcdc"    # 淡米
  $paper: "#ffffff"
  $primary: "#0081a7"   # 海青
  $accent: "#f07167"    # 珊瑚
  $text: "#1d3557"
  $muted: "#7c9aa6"
  $line: "#e4ece0"
  $ink: "#0081a7"
  $green: "#00afb9"
  $brass: "#fed9b7"
```

### 17. 活力橙薄荷（明亮欢快）
```yaml
colors:
  $bg: "#ffffff"
  $paper: "#f5fbfa"
  $primary: "#ff9f1c"   # 亮橙
  $accent: "#2ec4b6"    # 薄荷绿
  $text: "#1f2937"
  $muted: "#94a3b8"
  $line: "#e5e7eb"
  $ink: "#ff9f1c"
  $green: "#2ec4b6"
  $brass: "#ffbf69"
```

### 18. 铂金白金（高端专业）
```yaml
colors:
  $bg: "#f5f5f5"
  $paper: "#ffffff"
  $primary: "#0a0a0a"   # 黑
  $accent: "#0070F3"    # 蓝（行动色）
  $text: "#1f2937"
  $muted: "#737373"
  $line: "#e5e5e5"
  $ink: "#0a0a0a"
  $green: "#10B981"
  $brass: "#D4AF37"     # 金
```

## 三、textStyles 推荐（所有主题通用，按需微调）

```yaml
textStyles:
  $cover-title: { fontSize: 48, bold: true, color: "$primary", letterSpacing: 1 }
  $title:       { fontSize: 36, bold: true, color: "$primary" }
  $h2:          { fontSize: 22, bold: true, color: "$text" }
  $h3:          { fontSize: 18, bold: true, color: "$text" }
  $body:        { fontSize: 14, bold: false, color: "$text", lineHeight: 1.4 }
  $muted:       { fontSize: 12, bold: false, color: "$muted" }
  $stat:        { fontSize: 44, bold: true, color: "$accent" }
  $stat-label:  { fontSize: 13, bold: false, color: "$muted" }
```

## 四、风格配方（圆角/间距，来自 design-style-skill）

同一配色可通过圆角/间距呈现 4 种风格（`shapeName: roundRect` 时配 `rectRadius` 无效——编译器固定圆角；用 bounds 与间距控制气质）：

| 风格 | 视觉特征 | 适用 | 页面边距 | 块间距 | 卡片圆角感 |
|---|---|---|---|---|---|
| **Sharp & Compact** | 方正、信息密度高 | 数据报表、专业报告 | 0.3" | 0.1–0.2" | 0（用 rect） |
| **Soft & Balanced** | 舒适留白、专业亲和 | 企业汇报、通用 | 0.4" | 0.15–0.25" | 小圆角 |
| **Rounded & Spacious** | 圆润宽松、现代 | 产品介绍、营销 | 0.5" | 0.2–0.35" | 中圆角 |
| **Pill & Airy** | 通透、高端感 | 发布会、品牌 | 0.6" | 0.3–0.5" | 大圆角/胶囊 |

> 编译器限制：`shapeName` 只支持 `rect` / `roundRect` / `oval` / `donut`。风格主要通过边距、间距、形状选择和色块大小表达，勿在单页堆满元素。

## 五、页面背景配色指南

- **浅色主题**（bg 为浅色）：封面/结尾可用 `$bg` 或 `$ink`（深色块）制造对比，正文页统一 `$bg`。
- **深色主题**（如科技与夜景）：封面正文用浅色 `$text`，强调用亮金 `$primary`；所有页保持深色 bg 一致，勿深浅混用。
- **一致性命门**：背景必须纯色；富文本颜色内联 #HEX/rgb()，禁止 $theme（55173 编辑器无法解析）。若无需编辑器预览（纯命令行编译），$theme 变量仍可用，但为保持 skill 通用性，一律内联。
- **一致性命门**：背景必须纯色；富文本颜色内联 #HEX/rgb()，禁止 $theme（55173 编辑器无法解析）。若无需编辑器预览（纯命令行编译），$theme 变量仍可用，但为保持 skill 通用性，一律内联。

---

## 六、DesignPreset 设计预设（JSON 引擎，借鉴 harness-anything）

JSON 数据驱动路径把配色升级为**四元组预设**：colors + fonts + spacing + rules（不只是色板，还带设计规则）。见 `scripts/presets.py`。

内置 4 大预设（`py -3 scripts/presets.py list` 查看全部）：

| 预设名 | 场景 | 主色 | 视觉占比 | 留白 | 色数上限 |
|---|---|---|---|---|---|
| `academic` | 学术答辩/基金 | 深蓝 #1A3C8B | 65% | 40% | 5 |
| `consultant` | 咨询报告 | 深蓝 #003366 | 60% | 45% | 4 |
| `business` | 商务汇报/课件 | 商务蓝 #005294 | 55% | 40% | 4 |
| `tech` | 科技/AI/数据 | 亮青 #22D3EE | 65% | 35% | 5 |

**18 套配色即预设**：`templates/themes/all-themes.yaml` 的每一套都可直接当预设用，
例如 `py -3 scripts/json2pptx.py deck.json --preset 15-pure-tech-blue`（JSON 引擎按预设名解析颜色变量）。

**预设 → JSON deck**：deck.json 的 `theme.colors` 写 `$primary` 等变量即可，
编译时由预设或内联 colors 解析。质量审查（quality_check.py）读预设的 rules 做打分。

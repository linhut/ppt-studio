# 规格：ppt-studio 渲染级质量门禁 render_check.py

- 日期：2026-09-07
- 状态：已批准（用户确认范围 A：渲染级质量门禁；检测项：文本溢出 + 画布越界、元素重叠、页面密度/留白）
- 借鉴来源：cathrynlavery/diagram-design（MIT）——`scripts/lint-render.py`（渲染级剪裁检测）、SKILL.md §6 连接器规则、§7 复杂度预算「密度 4/10」哲学

## 背景与差距

ppt-studio 现有 `quality_check.py` 是纯静态分析——基于 deck.json 几何数据做容量精算与越界估算，看不见真实渲染结果。diagram-design 的 `lint-render.py` 用无头 Chromium 做像素级校验（截图 + overflow 释放 diff），是其「编辑级质量」的工程保证。本规格把这个思路移植到 ppt-studio：用本机 PowerPoint COM 把编译产物渲染为 PNG，再做像素级分析。

## 技术路径（已全部验证可行）

```
deck.json ──→ json2pptx.py ──→ deck.pptx ──→ PowerPoint COM ──→ 每页 PNG (960×540)
                                                       │
                                            Pillow 12.2 + numpy 2.4 像素分析
                                                       │
      ┌──────────────┬──────────────┬───────────────────┴──────────────┐
      ▼              ▼              ▼                                  ▼
  文本溢出检测    画布越界检测    元素重叠检测                        密度/留白检测
      └──────────────┴──────────────┴──────────────────────────────────┘
                                 │
                    逐页报告 + 总分 + 退出码（--strict 拦截）
```

- PowerPoint COM 渲染已验证：`examples/engine-test/deck.pptx` 9 页全部导出 PNG 成功
- Pillow 12.2.0 / numpy 2.4.6 已安装

## 新增/修改文件

1. **`scripts/render_check.py`**（新增，主脚本）
2. **`scripts/quality_check.py`**（修改：增加 `--render` 集成开关）
3. **`SKILL.md`**（修改：QA 章节补充渲染级检查步骤）

## 检测逻辑详细设计

### 1. 渲染（PowerPoint COM）
- `--pptx <file>`：直接渲染已有 pptx
- `--json <file>`：先调 json2pptx 编译为临时 pptx 再渲染（一键全流程）
- 输出目录：`<deck 目录>/_render_check/`，每页 `slide-<n>.png`（960×540）
- COM 渲染要点：`Presentations.Open(path, ReadOnly, Untitled, WithWindow=false)`、`Presentation.Export(dir, "PNG", 960, 540)`，用完必须 `Close()` + `Quit()` 释放进程
- 中文路径/文件名安全：用绝对路径、COM 调用后检查导出文件存在

### 2. 背景色提取
- 取 PNG 四角区域（8×8）像素中位数作为页面背景色
- 内容像素判定：与背景色 RGB 距离 > 阈值（默认 20/255）即算内容
- 说明：背景为纯色（HARD-GATE 要求），故四角中位数可靠；渐变背景取均值即可

### 3. 文本溢出检测
- 输入：pptx 的每页 shape 列表（python-pptx 读取真实编译产物），含 text_frame 的 shape
- 原理：对比「shape 的几何 bounds（EMU→pt→px）」与「该区域内文字的渲染像素范围」
- 检测：对每个文本 shape，取 bounds 向外扩 6px 的区域；统计该区域内内容像素的连通分布——
  - 若文字渲染像素明显超出 shape bounds（shape 外 6px 条带内有文字色像素且非相邻 shape 的内容），判定溢出
  - 简化稳健版：检查 shape bounds 外侧 4px 条带中「属于本 shape 文字颜色」的像素占比
- 文字色估计：从 shape 第一个 run 的颜色（解析为 RGB）近似；多色文本取暗色簇
- 输出：`slide-2: 文本 "xxxxxxxx" 疑似溢出文本框（右边界外 12px 检测到文字像素）`

### 4. 画布越界检测
- 检查 PNG 四边缘（各 2px 宽条带）是否存在内容像素（非背景）
- 若边缘条带内容像素占比 > 阈值（0.5%），判定内容顶边/越界
- 输出：`slide-3: 内容触及画布右边缘（边缘内容像素 34 个）——可能有元素被裁剪`

### 5. 元素重叠检测
- 输入：python-pptx 读取每页 shape 的 bounds（含 text frame）
- 几何重叠：两两 shape 矩形相交且相交面积 > 较小者 15%、且非「完全包含」（卡片内文字合法）→ 可疑重叠
- 像素佐证：在重叠区域，若存在「文字色像素与另一图形填充色像素交错」→ 确认文字被图形盖住
- 输出：`slide-4: 元素重叠：文本 "标题" (120,80,400,46) 与图形 rect (100,70,430,60) 相交 32%——检查是否被遮挡`
- 过滤：跳过装饰性细条（高/宽 ≤6pt）、跳过纯背景矩形（与页面同尺寸）、跳过无文字无填充的 shape

### 6. 密度/留白检测
- 内容像素占比 = 内容像素数 / 总像素数
- 借鉴 diagram-design「目标密度 4/10」：目标内容占比 10%–40% 视为健康
  - < 8%：页面过空（可能是内容缺失或留白过度）
  - > 55%：页面过密（提示拆页）
- 输出：`slide-5: 内容密度 62% > 55%——建议拆分或精简`
- 注意：纯色大块背景图形（如全页色块）会误报，需把「与页面同尺寸的填充矩形」排除在内容像素统计外（视为背景装饰）

### 7. 评分与输出
- 每页 100 分起，每个问题按严重度扣分：越界 -15 / 溢出 -10 / 重叠 -8 / 密度异常 -5
- 总分 = 各页平均
- 输出：人类可读报告（逐页问题列表 + 总分）+ `--json` 机器可读（供 CI 集成）
- `--strict`：总分 < 90 或任一 fail 项时退出码 1（可拦截）；默认退出码 0（报告不拦截）

## 集成

- `quality_check.py` 增加 `--render`：静态检查完成后自动调用 render_check.py 做渲染级复核，汇总两段报告
- SKILL.md §5 QA 补充：
  ```
  py -3 scripts/render_check.py my-deck/deck.json   # 渲染级检查（需本机 PowerPoint）
  py -3 scripts/quality_check.py my-deck/deck.json --preset tech --auto-only --render
  ```

## 验证计划

1. `examples/engine-test`（9 页，静态 100 分）→ render_check 应基本通过，验证无假阳性
2. 人造溢出用例（构造超长文本 / 越界元素 / 重叠元素 / 过密页）→ 应准确命中
3. 中文路径与文件名 → 已验证 COM 可导出

## YAGNI（不做）

- 不做品牌 onboarding（范围外，后续再说）
- 不做 CI 基础设施（本机无 CI；以 `--strict` 退出码预留拦截点）
- 不引入 Playwright（PowerPoint COM 是真实渲染器且零额外依赖）
- 不做 draw.io / Mermaid 导入

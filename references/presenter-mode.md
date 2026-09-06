<!--
  (c) 2026 Jose AI (https://www.linhut.cn)
  https://github.com/linhut/ppt-studio
  Licensed under the MIT License. See the LICENSE file for details.
-->
# 演讲者逐字稿指南（html-ppt presenter mode → PPTD notes）

> 借鉴 html-ppt 的 presenter-mode-reveal 全 deck 模板思路（S 键演讲者模式），
> 迁移到 PPTD 工程：每页写 150–300 字口语化讲稿，存于 .page 的 `notes` 字段。

## 何时使用

用户提到：**演讲 / 分享 / 讲稿 / 逐字稿 / speaker notes / 演讲者视图 / 提词器**，
或说"我要去给团队讲 xxx""要做一场技术分享""怕讲不流畅""想要一份带逐字稿的 PPT"。

## 写稿三规则（来自 html-ppt）

1. **不是讲稿，是提示信号** — 加粗核心词，过渡句独立成段
2. **每页 150–300 字** — 2–3 分钟/页的节奏
3. **用口语，不用书面语** — "因此"→"所以"，"该方案"→"这个方案"

## 在 PPTD 中落地

每页 .page 加 `notes` 字段：

```yaml
id: "04_market"
pageType: "content"
background: { color: "$bg" }
elements: [...]
notes: |
  **开场**：这一页讲市场概况。
  先说结论：整个行业未来三年会保持 30% 以上的增速。
  然后给两个数字：今年规模 120 亿，明年预计 160 亿。
  **过渡**：那在这个大盘里，我们切的是哪块蛋糕？看下一页。
```

交付时导出全部备注为 `speaker-notes.md`（按页号组织），方便打印：

```markdown
# 演讲者逐字稿

## P01 封面
（150–300 字口语稿……）

## P04 市场概况
（……）
```

## 结构与节奏模板（技术分享 / 汇报通用）

| 页 | 类型 | 逐字稿要点 |
|---|---|---|
| 01 | cover | 30 秒自我介绍 + 主题一句话 + 讲多久 |
| 02 | toc | 快速报菜名，强调"最重点在第四部分" |
| 03–N | section/content | 每页一个观点：结论→证据→过渡 |
| N+1 | summary | 三个 takeaways 回顾 |
| N+2 | cta/thanks | 行动号召 + 致谢 + 联系方式 |

## 检查清单

- [ ] 每页 150–300 字（不是 50 字也不是 500 字）
- [ ] 口语化：无"综上所述""值得注意的是"式书面语
- [ ] 有过渡句：每页结尾指向下一页
- [ ] 核心词加粗标记（**xxx**）
- [ ] 讲稿与页面内容一致，不出现页面上没有的信息

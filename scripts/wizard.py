#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
# https://github.com/linhut/ppt-studio
# Licensed under the MIT License. See the LICENSE file for details.

"""ppt-studio 向导生成器 v3：answers.json（6 问答案 + 布局/语言偏好）→ deck.json 骨架。

用法：
  py -3 scripts/wizard.py my-deck/answers.json --output my-deck/deck.json
  py -3 scripts/wizard.py my-deck/answers.json -o my-deck/deck.json --strict

answers.json 结构（与 SKILL.md「0. 向导式需求确认」一一对应）：
{
  "topic": "ppt-studio 能力展示",          # 1 主题内容
  "structure": "full",                     # 2 页数结构: full | mini | large
  "palette": "15-pure-tech-blue",          # 3 配色方案: 18 套 key 或内置预设名
  "style": "soft",                         # 4 风格配方: sharp | soft | rounded | pill
  "mode": "light",                         # 5 深浅模式: light | dark | mixed
  "tone": "corporate",                     # 6 语言风格（内容最重要）：见下表
  "title": "ppt-studio v3",                # 封面主标题（缺省用 topic）
  "subtitle": "一次对话，产出原生 .pptx 文件",  # 封面副标题（缺省随 tone 生成）
  "tocTitle": "今天的内容",                 # 目录页标题
  "outline": ["三源融合", "设计系统", "双引擎", "复合组件", "质量闭环", "路线图"],
  "content_layouts": ["cards", "kpi", "process", "timeline", "compare"]
}
tone（语言风格）可选，缺省 corporate。完整指南见 references/tone-guide.md：
  corporate   商务正式（汇报/提案）    casual    轻松口语（分享/培训）
  tech        科技极客（技术评审）     story     叙事故事（发布会/演讲）
  persuasive  说服销售（pitch/竞标）   academic  学术报告（答辩/调研）
  government  政务公文体（汇报/党建）  语言严谨简洁、公文语体
tone 决定封面副标语、内容页占位提示、结语金句的腔调，保证全文文案风格一致。

content_layouts 可选（默认即上述 5 种），从布局样式库
（references/layout-catalog.md）挑选：
  text      要点+侧卡     cards     2x3 卡片网格   kpi       KPI 横排
  table     数据表        chart     柱状图+解读    process   流程步骤
  timeline  横向时间线    compare   双卡对比        proscons  利弊两卡
  roadmap   路线图四列    checklist 清单           quote     大引语

输出 deck.json 骨架：主题变量（配色映射 + 风格圆角/间距 + 深浅模式 + 语言风格）
+ 页面结构（封面/目录/内容页 xN/致谢），元素为占位符，填充文案后即可编译。
"""

import argparse
import json
import sys
from pathlib import Path

STUDIO_DIR = Path(__file__).resolve().parent.parent

# ---- 18 套配色 key（templates/themes/all-themes.yaml）----
def load_palette(key):
    import yaml
    yaml_path = STUDIO_DIR / "templates" / "themes" / "all-themes.yaml"
    data = yaml.safe_load(yaml_path.read_text(encoding="utf-8"))
    themes = data.get("themes", {})
    if key in themes:
        colors = {}
        for k, v in themes[key].get("colors", {}).items():
            colors[k if k.startswith("$") else "$" + k] = v
        return colors, themes[key].get("name", key)
    sys.path.insert(0, str(STUDIO_DIR / "scripts"))
    try:
        from presets import get_preset
        pr = get_preset(key)
        if pr:
            return dict(pr.colors), pr.display_name
    except Exception:
        pass
    raise SystemExit("未知配色: " + key + "（查 py -3 scripts/presets.py list）")

# ---- 风格配方（design-style-skill 英寸值 → pt 常量）----
STYLES = {
    "sharp":   {"radius": 4,  "gap": 18, "margin": 36, "padding": 12},
    "soft":    {"radius": 12, "gap": 24, "margin": 48, "padding": 16},
    "rounded": {"radius": 18, "gap": 28, "margin": 52, "padding": 20},
    "pill":    {"radius": 28, "gap": 32, "margin": 56, "padding": 24},
}

# ---- 语言风格（内容最重要！指南见 references/tone-guide.md）----
TONES = {
    "corporate": {"label": "商务正式", "sub": "严谨 · 数据 · 书面化",
                  "body": "填充本页核心内容：先给结论，再列数据与依据支撑。",
                  "quote": "以数据为证，以结果说话，以行动收尾。",
                  "author": "ppt-studio · 商务表达原则"},
    "casual": {"label": "轻松口语", "sub": "亲切 · 短句 · 对话感",
               "body": "用大白话讲清楚这件事：先讲人话，再补细节。",
               "quote": "说人话、讲重点，让每个人都能跟上。",
               "author": "ppt-studio · 口语化表达原则"},
    "tech": {"label": "科技极客", "sub": "精准 · 术语 · 简洁有力",
             "body": "说明技术要点：架构、机制、关键指标，术语精准确切。",
             "quote": "保持简单，但不简化。",
             "author": "ppt-studio · 技术写作原则"},
    "story": {"label": "叙事故事", "sub": "有情节 · 场景化 · 打动人",
              "body": "讲一个场景：痛点 → 转折 → 解法，让听众代入其中。",
              "quote": "最好的技术，讲述最动人的故事。",
              "author": "ppt-studio · 叙事表达原则"},
    "persuasive": {"label": "说服销售", "sub": "利益点 · 差异 · 行动号召",
                   "body": "突出收益与差异：对方能得到什么，为什么选我们。",
                   "quote": "选择确定性，选择我们。",
                   "author": "ppt-studio · 说服表达原则"},
    "academic": {"label": "学术报告", "sub": "客观 · 引用 · 分析框架",
                 "body": "按框架陈述：问题背景 → 方法 → 证据 → 结论。",
                 "quote": "证据驱动结论，逻辑贯穿始终。",
                 "author": "ppt-studio · 学术表达原则"},
    "government": {"label": "政务公文体", "sub": "严谨 · 规范 · 权威 · 简洁",
                   "body": "按公文语体撰写：观点先行、层次分明，用词规范准确（深入贯彻/扎实推进/统筹推进），不口语化。",
                   "quote": "大道至简，实干为要。",
                   "author": "政务表达原则"},
}
DEFAULT_TONE = "corporate"

# ---- 深浅模式：覆盖 $bg/$paper/$text/$muted/$line 明暗 ----
def apply_mode(colors, mode):
    c = dict(colors)
    if mode == "dark":
        bg = c.get("$ink") or c.get("$primary") or "#111111"
        c["$bg"] = bg
        c["$paper"] = "#1E293B"
        c["$text"] = "#F1F5F9"
        c["$muted"] = "#94A3B8"
        c["$line"] = "#334155"
    elif mode == "mixed":
        c["$bg"] = "#F8FAFC"
        c["$paper"] = "#FFFFFF"
        c["$text"] = "#1F2937"
        c["$muted"] = "#64748B"
        c["$line"] = "#E2E8F0"
    return c


def _header(kicker_text, title_text, colors):
    """内容页通用页眉：kb竖排 kicker + 标题 + 强调短横线。"""
    return [
        {"type": "text", "role": "kicker", "x": 48, "y": 14, "w": 300, "h": 18,
         "text": kicker_text, "fontSize": 12, "color": "$accent", "bold": True},
        {"type": "text", "role": "title", "x": 48, "y": 36, "w": 864, "h": 42,
         "text": title_text, "fontSize": 36, "color": "$primary", "bold": True},
        {"type": "shape", "shapeName": "rect", "x": 48, "y": 86, "w": 48, "h": 3,
         "fill": {"color": "$accent"}},
    ]


def _content_page(seq, layout, title, colors, style, tc=None):
    """按布局库生成一个内容页骨架。layout 取值见 references/layout-catalog.md。
    tc 为语言风格配置（TONES 项），决定占位提示与结语金句的腔调。"""
    tc = tc or TONES[DEFAULT_TONE]
    n = "%02d" % (4 + seq) if seq < 5 else "%02d" % (9 + seq - 5)
    e = _header("SECTION · " + n, title, colors) if not layout == "cover_asym" else []
    L = layout
    if L == "text":
        e += [
            {"type": "text", "role": "body", "x": 48, "y": 110, "w": 500, "h": 300,
             "text": tc["body"], "fontSize": 16, "color": "$text", "lineHeight": 1.6},
            {"type": "shape", "shapeName": "roundRect", "x": 600, "y": 110, "w": 312, "h": 300,
             "fill": {"color": "$paper"}, "border": {"color": "$line", "width": 1}, "rectRadius": 0.04},
        ]
    elif L == "cards":
        e += [{"type": "cards_2x3", "x": 48, "y": 110, "w": 864, "h": 352, "gap": style["gap"],
               "items": [{"title": "要点 %d" % (i + 1), "desc": "填充这一要点",
                          "color": "$accent" if i % 2 == 0 else "$green"} for i in range(6)]}]
    elif L == "kpi":
        e += [
            {"type": "cards_1x4_info", "x": 48, "y": 110, "w": 864, "h": 150, "gap": 18,
             "items": [
                 {"num": "01", "label": "填充指标", "color": "$primary"},
                 {"num": "02", "label": "填充指标", "color": "$accent"},
                 {"num": "03", "label": "填充指标", "color": "$green"},
                 {"num": "04", "label": "填充指标", "color": "$brass"},
             ]},
            {"type": "text", "role": "body", "x": 48, "y": 300, "w": 864, "h": 160,
             "text": "在这里补充这一页的说明与要点。", "fontSize": 16, "color": "$text", "lineHeight": 1.6},
        ]
    elif L == "table":
        e += [
            {"type": "table", "x": 48, "y": 110, "w": 864, "h": 300, "rows": 5, "cols": 3,
             "header_color": "$ink",
             "data": [["", "列 A", "列 B"],
                      ["维度 1", "填充内容", "填充内容"],
                      ["维度 2", "填充内容", "填充内容"],
                      ["维度 3", "填充内容", "填充内容"],
                      ["维度 4", "填充内容", "填充内容"]]},
            {"type": "text", "role": "caption", "x": 48, "y": 440, "w": 864, "h": 24,
             "text": "表下注释：填充说明文字", "fontSize": 12, "color": "$muted", "align": "center"},
        ]
    elif L == "chart":
        e += [
            {"type": "chart", "x": 48, "y": 110, "w": 560, "h": 360, "color": "$green",
             "data": {"rows": [["A", "分类", 30], ["B", "分类", 45], ["C", "分类", 25],
                               ["D", "分类", 60], ["E", "分类", 40]]}},
            {"type": "text", "role": "body", "x": 640, "y": 116, "w": 272, "h": 320,
             "text": "图表说明：在这里补充数据解读与结论。", "fontSize": 14, "color": "$text", "lineHeight": 1.6},
        ]
    elif L == "process":
        e += [{"type": "process_steps", "x": 48, "y": 130, "w": 864, "h": 300, "color": "$accent",
               "steps": [{"title": "步骤 %d" % (i + 1), "desc": "填充该步骤说明"} for i in range(4)]}]
    elif L == "timeline":
        e += [{"type": "timeline_h", "x": 48, "y": 130, "w": 864, "h": 300, "color": "$accent",
               "items": [{"time": "Q%d" % (i + 1), "title": "阶段 %d" % (i + 1),
                          "desc": "填充该阶段说明"} for i in range(4)]}]
    elif L == "compare":
        e += [{"type": "comparison_2col", "x": 48, "y": 110, "w": 864, "h": 340,
               "left": {"title": "方案 A", "points": ["填充优势一", "填充优势二", "填充优势三"]},
               "right": {"title": "方案 B", "points": ["填充优势一", "填充优势二", "填充优势三"]}}]
    elif L == "proscons":
        e += [{"type": "pros_cons", "x": 48, "y": 110, "w": 864, "h": 340,
               "pros": ["填充优点一", "填充优点二", "填充优点三"],
               "cons": ["填充不足一", "填充不足二"]}]
    elif L == "roadmap":
        e += [{"type": "roadmap_4col", "x": 48, "y": 120, "w": 864, "h": 320,
               "columns": [{"phase": ph, "items": ["填充事项一", "填充事项二"]}
                           for ph in ("NOW", "NEXT", "LATER", "VISION")]}]
    elif L == "checklist":
        e += [{"type": "checklist", "x": 48, "y": 120, "w": 864, "h": 300,
               "items": [{"text": "填充检查项 %d" % (i + 1), "done": i % 2 == 0} for i in range(6)]}]
    elif L == "quote":
        e += [{"type": "big_quote", "x": 80, "y": 160, "w": 800, "h": 240,
               "quote": tc["quote"], "author": tc["author"]}]
    else:
        e += [{"type": "text", "role": "body", "x": 48, "y": 110, "w": 864, "h": 300,
               "text": tc["body"], "fontSize": 16, "color": "$text", "lineHeight": 1.6}]
    return e


def build_deck(answers):
    palette_key = answers.get("palette", "15-pure-tech-blue")
    colors, pname = load_palette(palette_key)
    colors = apply_mode(colors, answers.get("mode", "light"))
    style = STYLES.get(answers.get("style", "soft"), STYLES["soft"])

    topic = answers.get("title") or answers.get("topic", "演示文稿")
    tone = answers.get("tone") or DEFAULT_TONE
    if tone not in TONES:
        tone = DEFAULT_TONE
    tc = TONES[tone]
    subtitle = answers.get("subtitle", "") or tc["sub"]
    toc_title = answers.get("tocTitle", "今天的内容")
    outline = answers.get("outline") or ["第一章", "第二章", "第三章", "第四章", "第五章", "第六章"]
    structure = answers.get("structure", "full")

    cover_bg = {"color": colors.get("$ink", "#03045e")} if answers.get("mode") == "dark" else {
        "type": "gradient",
        "stops": [{"color": colors.get("$ink", "#03045e")}, {"color": colors.get("$accent", "#0077b6")}],
        "angle": 135,
    }
    final_bg = {"color": colors.get("$ink", "#03045e")} if answers.get("mode") == "dark" else {
        "type": "gradient",
        "stops": [{"color": colors.get("$accent", "#0077b6")}, {"color": colors.get("$ink", "#03045e")}],
        "angle": 45,
    }

    # 内容页布局序列（默认轮换 5 种新排版，可被 content_layouts 覆盖）
    layouts = answers.get("content_layouts") or ["cards", "kpi", "process", "timeline", "compare"]
    content_pages = []
    for i, layout in enumerate(layouts[:6]):
        ot = outline[i] if i < len(outline) else ("第 %d 章" % (i + 1))
        content_pages.append({
            "id": "%02d" % (3 + i),
            "title": ot,
            "elements": _content_page(i, layout, ot, colors, style, tc),
        })
    # 结语页：大引语（若布局序列未包含 quote 则补一张）
    if "quote" not in layouts:
        content_pages.append({
            "id": "09", "title": "结语",
            "elements": _content_page(5, "quote", "结语", colors, style, tc),
        })
    elif len(content_pages) < 6:
        ot = outline[5] if len(outline) > 5 else "第 6 章"
        content_pages.append({
            "id": "09", "title": ot,
            "elements": _content_page(5, "text", ot, colors, style, tc),
        })

    deck = {
        "_doc": "ppt-studio 向导生成骨架：主题 %s（%s）+ %s 风格 + %s 模式 + %s 语言 + 布局 %s"
                % (pname, palette_key, answers.get("style", "soft"), answers.get("mode", "light"),
                   tc["label"], "/".join(layouts)),
        "canvas": {"w": 960, "h": 540},
        "theme": {"colors": colors},
        "slides": [
            {
                "id": "01", "title": "封面",
                "background": cover_bg,
                "elements": [
                    {"type": "text", "role": "kicker", "x": 96, "y": 64, "w": 768, "h": 20,
                     "text": "PPT 全能工坊 · 向导生成", "fontSize": 12, "color": "#90e0ef", "align": "center"},
                    {"type": "text", "role": "title", "x": 96, "y": 96, "w": 768, "h": 84,
                     "text": topic, "fontSize": 54, "color": "#FFFFFF", "bold": True, "font": "Arial", "align": "center"},
                    {"type": "text", "role": "body", "x": 96, "y": 214, "w": 768, "h": 32,
                     "text": subtitle, "fontSize": 18, "color": "#caf0f8", "align": "center"},
                    {"type": "tagline_bar", "y": 498,
                     "text": "ppt-studio — 一次对话生成原生 PPTX"},
                ],
            },
            {
                "id": "02", "title": "目录",
                "elements": [
                    {"type": "text", "role": "kicker", "x": 48, "y": 14, "w": 300, "h": 18,
                     "text": "CONTENTS · 目录", "fontSize": 12, "color": "$accent", "bold": True},
                    {"type": "text", "role": "title", "x": 48, "y": 36, "w": 864, "h": 42,
                     "text": toc_title, "fontSize": 36, "color": "$primary", "bold": True},
                    {"type": "shape", "shapeName": "rect", "x": 48, "y": 86, "w": 48, "h": 3,
                     "fill": {"color": "$accent"}},
                    {"type": "cards_2x3", "x": 48, "y": 110, "w": 864, "h": 352, "gap": 24,
                     "items": [
                         {"title": "%02d · %s" % (i + 1, t), "desc": "填充这一章的要点描述",
                          "color": "$accent" if i % 2 == 0 else "$green"}
                         for i, t in enumerate(outline[:6])
                     ]},
                ],
            },
        ] + content_pages + [
            {
                "id": "10", "title": "致谢",
                "background": final_bg,
                "elements": [
                    {"type": "text", "role": "title", "x": 96, "y": 180, "w": 768, "h": 84,
                     "text": "谢谢观看", "fontSize": 52, "color": "#FFFFFF", "bold": True, "align": "center"},
                    {"type": "text", "role": "body", "x": 96, "y": 296, "w": 768, "h": 32,
                     "text": "Q & A  ·  欢迎交流", "fontSize": 18, "color": "#caf0f8", "align": "center"},
                    {"type": "tagline_bar", "y": 498, "text": "ppt-studio — 一次对话生成原生 PPTX"},
                ],
            },
        ],
    }
    # 精简模式：封面/目录/第一个内容页/致谢
    if structure == "mini":
        keep = {"01", "02", "03", "10"}
        deck["slides"] = [s for s in deck["slides"] if s["id"] in keep]
    return deck


def main(argv=None):
    parser = argparse.ArgumentParser(description="ppt-studio 向导生成器 v2：answers.json → deck.json 骨架")
    parser.add_argument("input", type=Path, help="answers.json 路径")
    parser.add_argument("-o", "--output", type=Path, help="输出 deck.json 路径")
    parser.add_argument("--strict", action="store_true", help="校验字段")
    args = parser.parse_args(argv)

    if not args.input.is_file():
        print("error: 找不到 " + str(args.input), file=sys.stderr)
        return 2
    answers = json.loads(args.input.read_text(encoding="utf-8"))
    deck = build_deck(answers)
    out = args.output or (args.input.parent / "deck.json")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(deck, ensure_ascii=False, indent=2), encoding="utf-8")
    print("向导生成 %d 页骨架 → %s" % (len(deck["slides"]), out))
    return 0


if __name__ == "__main__":
    sys.exit(main())

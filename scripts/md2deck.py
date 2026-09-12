#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
# https://github.com/linhut/ppt-studio
# Licensed under the MIT License. See the LICENSE file for details.

"""ppt-studio Markdown 大纲 → deck.json 转换器（ppt md2deck）。

Markdown 约定（按此结构组织，其余行忽略）：
  # 标题                白板标题 → 封面（紧随其后的 > 引用行作为副标题）
  ## 章节标题           每个 ## 生成一页
  - 条目 / * 条目        该页条目（"标题：说明" 可拆分标题与补充说明）
  <!-- layout: xxx -->  布局提示：cards(默认) / kpi / text，放在对应 ## 前

数字型条目（含 %、★、k、万 或纯数字）自动走 KPI 布局。

用法：
  ppt md2deck outline.md -o my-deck/deck.json --theme tech
"""

import argparse
import json
import re
import sys
from pathlib import Path

STUDIO_DIR = Path(__file__).resolve().parent.parent

NUM_RE = re.compile(r"^[-+]?[\d,]+(\.\d+)?(%|★|k|万|亿)?$")
SPLIT_RE = re.compile(r"[:：]")


def _load_theme(theme):
    sys.path.insert(0, str(STUDIO_DIR / "scripts"))
    try:
        from presets import get_preset
        pr = get_preset(theme)
        if pr is None:
            print("error: 未知预设 %s（查 ppt presets list）" % theme, file=sys.stderr)
            return None
        return dict(pr.colors)
    except Exception as exc:
        print("error: 读取预设失败：%s" % exc, file=sys.stderr)
        return None


def _split_item(line):
    """按 中文/英文冒号 拆分「标题：说明」，拆不开则整行作标题。"""
    m = SPLIT_RE.split(line, maxsplit=1)
    if len(m) == 2 and m[1].strip():
        return m[0].strip(), m[1].strip()
    return line.strip(), ""


def _kpi_items(bullets):
    items = []
    for b in bullets:
        title, label = _split_item(b)
        items.append({"num": title if NUM_RE.match(title) else title, "label": label})
    return items


def _text_elements(title, bullets, colors):
    """纯文本页：标题 + 项目符号正文。"""
    body = "\n".join("- " + b.strip() for b in bullets)
    return [
        {"type": "text", "role": "kicker", "x": 48, "y": 14, "w": 300, "h": 18,
         "text": "CONTENT · 内容", "fontSize": 12, "color": "$accent", "bold": True},
        {"type": "text", "role": "title", "x": 48, "y": 36, "w": 864, "h": 42,
         "text": title, "fontSize": 36, "color": "$primary", "bold": True},
        {"type": "shape", "shapeName": "rect", "x": 48, "y": 86, "w": 48, "h": 3,
         "fill": {"color": "$accent"}},
        {"type": "text", "role": "body", "x": 60, "y": 120, "w": 840, "h": 380,
         "text": body or "填充本页内容", "fontSize": 18, "color": "$text", "lineHeight": 1.5},
    ]


def _cards_elements(title, bullets, colors):
    items = []
    for b in bullets:
        t, sub = _split_item(b)
        items.append({"title": t, "sub": sub})
    n = max(len(items), 1)
    item_h = max(40, min(58, (540 - 170) // n))
    return [
        {"type": "text", "role": "kicker", "x": 48, "y": 14, "w": 300, "h": 18,
         "text": "CONTENT · 内容", "fontSize": 12, "color": "$accent", "bold": True},
        {"type": "text", "role": "title", "x": 48, "y": 36, "w": 864, "h": 42,
         "text": title, "fontSize": 36, "color": "$primary", "bold": True},
        {"type": "shape", "shapeName": "rect", "x": 48, "y": 86, "w": 48, "h": 3,
         "fill": {"color": "$accent"}},
        {"type": "card_list_wide", "x": 60, "start_y": 120, "item_h": item_h,
         "items": items or [{"title": "填充这一章要点", "sub": "一句话说明"}]},
    ]


def _kpi_elements(title, bullets, colors):
    return [
        {"type": "text", "role": "kicker", "x": 48, "y": 14, "w": 300, "h": 18,
         "text": "KPI · 关键指标", "fontSize": 12, "color": "$accent", "bold": True},
        {"type": "text", "role": "title", "x": 48, "y": 36, "w": 864, "h": 42,
         "text": title, "fontSize": 36, "color": "$primary", "bold": True},
        {"type": "shape", "shapeName": "rect", "x": 48, "y": 86, "w": 48, "h": 3,
         "fill": {"color": "$accent"}},
        {"type": "cards_1x4_info", "x": 60, "y": 150, "w": 840, "h": 160,
         "items": _kpi_items(bullets)},
    ]


BUILDERS = {"cards": _cards_elements, "kpi": _kpi_elements, "text": _text_elements}


def parse_markdown(text):
    """解析 Markdown 大纲 → (deck_title, subtitle, slides[])。"""
    title, subtitle = "演示文稿", ""
    slides = []
    cur_title = None
    cur_bullets = []
    cur_layout = "cards"

    def flush():
        nonlocal cur_title, cur_bullets, cur_layout
        if cur_title is not None:
            slides.append({"title": cur_title, "bullets": cur_bullets,
                           "layout": cur_layout})
        cur_title, cur_bullets = None, []

    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        if line.startswith("<!--"):
            m = re.search(r"layout:\s*(\w+)", line)
            if m and m.group(1) in BUILDERS:
                cur_layout = m.group(1)
            continue
        if line.startswith("# "):
            flush()
            title = line[2:].strip()
        elif line.startswith("## "):
            flush()
            cur_title = line[3:].strip()
            cur_layout = "cards"
        elif line.startswith("### "):
            continue
        elif line.startswith(("> ", ">")):
            if not slides and not cur_title:
                subtitle = line.lstrip("> ").strip()
            continue
        elif line.startswith(("- ", "* ", "-", "*")):
            if cur_title is None:
                continue
            cur_bullets.append(line[1:].strip())
    flush()
    return title, subtitle, slides


def build_deck(title, subtitle, slides, colors):
    elements = [
        {"type": "shape", "shapeName": "rect", "x": 0, "y": 0, "w": 960, "h": 8,
         "fill": {"color": "$primary"}},
        {"type": "shape", "shapeName": "rect", "x": 0, "y": 532, "w": 960, "h": 8,
         "fill": {"color": "$accent"}},
        {"type": "text", "role": "title", "x": 60, "y": 130, "w": 840, "h": 90,
         "text": title, "fontSize": 54, "color": "$primary", "bold": True, "align": "center"},
        {"type": "text", "x": 60, "y": 240, "w": 840, "h": 44,
         "text": subtitle or "ppt-studio · Markdown 大纲生成", "fontSize": 22,
         "color": "$text", "align": "center"},
        {"type": "text", "x": 60, "y": 300, "w": 840, "h": 32,
         "text": "演讲人  |  日期", "fontSize": 16, "color": "$muted", "align": "center"},
        {"type": "tagline_bar", "y": 498, "text": "ppt-studio — 一次对话生成原生 PPTX"},
    ]
    slides_out = [{"id": "01", "title": "封面",
                   "background": {"color": "$bg"}, "elements": elements}]

    for i, s in enumerate(slides):
        # 数字型条目自动走 KPI 布局（拆「标题：说明」后取标题段判断）
        nums = [NUM_RE.match(_split_item(b)[0]) for b in s["bullets"]]
        key = s["layout"]
        if key == "cards" and s["bullets"] and all(nums):
            key = "kpi"
        if key == "kpi" and not any(nums):
            key = "cards"
        builder = BUILDERS.get(key, _cards_elements)
        slides_out.append({
            "id": "%02d" % (i + 2),
            "title": s["title"],
            "elements": builder(s["title"], s["bullets"], colors),
        })
    return {
        "_doc": "ppt-studio 生成的 Markdown 大纲 deck（%d 页）" % len(slides_out),
        "canvas": {"w": 960, "h": 540},
        "theme": {"colors": colors},
        "slides": slides_out,
    }


def main(argv=None):
    parser = argparse.ArgumentParser(description="ppt-studio Markdown 大纲 → deck.json")
    parser.add_argument("input", help="Markdown 文件路径")
    parser.add_argument("-o", "--output", default=None, help="输出 deck.json 路径")
    parser.add_argument("--theme", default="tech", help="设计预设名（默认 tech）")
    args = parser.parse_args(argv)

    src = Path(args.input)
    if not src.is_file():
        print("error: 找不到 " + str(src), file=sys.stderr)
        return 2
    colors = _load_theme(args.theme)
    if colors is None:
        return 2

    title, subtitle, slides = parse_markdown(src.read_text(encoding="utf-8"))
    if not slides:
        print("error: Markdown 中未找到 ## 章节标题", file=sys.stderr)
        return 2
    deck = build_deck(title, subtitle, slides, colors)

    out = Path(args.output) if args.output else (src.parent / "deck.json")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(deck, ensure_ascii=False, indent=2), encoding="utf-8")
    print("md2deck 生成 %d 页 → %s" % (len(deck["slides"]), out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
# https://github.com/linhut/ppt-studio
# Licensed under the MIT License. See the LICENSE file for details.

"""ppt-studio deck.json schema 校验（ppt validate）。

逐级检查 deck.json 结构与元素字段，报告带 JSONPath 定位与行号：
  slides[2].elements[1].fill.color  (第 47 行)
  error   结构性问题（必改）—— 直接退出码 1
  warning 建议性问题（--strict 时退出码 1）

用法：
  ppt validate deck.json
  ppt validate deck.json --strict
"""

import argparse
import json
import re
import sys
from pathlib import Path

STUDIO_DIR = Path(__file__).resolve().parent.parent

# 元素类型总表（与 json-engine.md 一致）
KNOWN_TYPES = {
    "text", "shape", "line", "circle", "image", "table", "chart",
    "cards_2x3", "cards_1x4_info", "card_list_wide", "tagline_bar", "num_big",
    "section_divider", "comparison_2col", "pros_cons", "timeline_h",
    "process_steps", "kpi_row", "big_quote", "roadmap_4col", "checklist",
    "cover_asym", "figure_text", "breadcrumb", "references",
}
SHAPE_NAMES = {"rect", "roundRect", "oval", "donut", "line", "arrow"}
# 需要 x/y 定位的元素（card_list_wide 用 x+start_y；tagline_bar/section_divider/num_big 自动定位）
BOUNDS_TYPES = {
    "text", "shape", "line", "circle", "image", "table", "chart",
    "cards_2x3", "cards_1x4_info", "comparison_2col", "pros_cons",
    "timeline_h", "process_steps", "kpi_row", "big_quote", "roadmap_4col",
    "checklist", "cover_asym", "figure_text", "breadcrumb", "references",
}

COLOR_RE = re.compile(r"^(\$[\w]+|#[0-9a-fA-F]{6}|#[0-9a-fA-F]{8}|rgb\([\d\s,%.]+\)|rgba\([\d\s,%.]+\))$")
# 复合组件的集合字段（类型 → (字段名, 期望类型, 条目数建议上限)；list/obj/str）
COLLECTIONS = {
    "cards_2x3": [("items", "list", 6)],
    "cards_1x4_info": [("items", "list", 4)],
    "card_list_wide": [("items", "list", 6)],
    "kpi_row": [("items", "list", None)],
    "timeline_h": [("items", "list", None)],
    "checklist": [("items", "list", None)],
    "process_steps": [("steps", "list", None)],
    "roadmap_4col": [("columns", "list", None)],
    "comparison_2col": [("left", "obj", None), ("right", "obj", None)],
    "pros_cons": [("pros", "list", None), ("cons", "list", None)],
    "big_quote": [("quote", "str", None), ("author", "str", None)],
}
KIND_NAMES = {"list": "列表", "obj": "对象", "str": "字符串"}
JSON_KV_RE = re.compile(r'"(%s)"\s*:' % "|".join(map(re.escape, (
    "canvas", "theme", "colors", "slides", "id", "title", "elements",
    "type", "x", "y", "w", "h", "text", "fontSize", "color", "bold",
    "shapeName", "fill", "items", "data", "rows", "cols", "header_color",
    "background", "lineHeight", "align", "num", "label", "sub", "desc",
    "start_y", "item_h", "gap", "kicker", "value", "name", "role",
    "steps", "columns", "left", "right", "pros", "cons", "quote", "author",
    "points", "phase", "time", "done"))))

_ISSUES = []


def build_offset_map(text):
    """扫描 JSON 原文，记录每个键路径 → 字符偏移（用于行号定位）。"""
    offsets = {}
    try:
        decoder = json.JSONDecoder()
        obj, _ = decoder.raw_decode(text)
    except Exception:
        return {}

    def walk(node, path, start):
        if isinstance(node, dict):
            pos = start
            for k, v in node.items():
                # 从当前区域向后找该键的引号位置
                m = JSON_KV_RE.search(text, pos)
                if m:
                    pos = m.start()
                    offsets[path + "/" + str(k)] = pos
                    nxt = m.end()
                    # 定位值的起始（跳过空白与冒号）
                    while nxt < len(text) and text[nxt] in " \t\r\n:":
                        nxt += 1
                    walk(v, path + "/" + str(k), nxt)
                else:
                    break
        elif isinstance(node, list):
            for i, v in enumerate(node):
                si = max(start, 0)
                m = None
                # 列表项定位：从当前位置向后找 { 或值起始
                for cand in (si, si + 1):
                    if cand < len(text) and text[cand:cand + 1] in "{[\"0-9":
                        m = cand
                        break
                if m is None:
                    m = si
                offsets[path + "/%d" % i] = m
                walk(v, path + "/%d" % i, m)

    walk(obj, "", 0)
    return offsets


def line_of(offsets, text, label):
    off = offsets.get(label)
    if off is None:
        return None
    return text.count("\n", 0, off) + 1


def issue(offsets, text, level, label, msg):
    ln = line_of(offsets, text, label)
    loc = label or "$"
    if ln:
        loc += "  (第 %d 行)" % ln
    print("  %-7s %-44s %s" % (level.upper(), loc, msg))
    _ISSUES.append((level, label, msg))


def validate_color(offsets, text, label, val):
    if isinstance(val, str) and COLOR_RE.match(val):
        return
    issue(offsets, text, "error", label,
          "颜色须为 $变量 / #RRGGBB / #RRGGBBAA / rgb()，当前：%r" % (val,))


def _type_of(e):
    return e.get("type") if isinstance(e, dict) else None


def validate_element(offsets, text, label, e):
    if not isinstance(e, dict):
        issue(offsets, text, "error", label, "元素须为对象，当前：%r" % (type(e).__name__,))
        return
    t = _type_of(e)
    if not t:
        issue(offsets, text, "error", label + "/type", "缺少必备字段 type")
        return
    if t not in KNOWN_TYPES:
        issue(offsets, text, "error", label + "/type", "未知元素类型 %r（支持见 references/json-engine.md）" % t)
        return
    base = label
    if t in BOUNDS_TYPES:
        for k in ("x", "y", "w", "h"):
            if not isinstance(e.get(k), (int, float)):
                issue(offsets, text, "error", base + "/" + k, "元素 %s 缺少数值型 %s" % (t, k))
    if t in ("text",):
        if not isinstance(e.get("text"), str) or not e.get("text"):
            issue(offsets, text, "error", base + "/text", "text 元素缺少非空字符串 text")
        if "color" in e:
            validate_color(offsets, text, base + "/color", e["color"])
        fs = e.get("fontSize")
        if fs is not None and not isinstance(fs, (int, float)):
            issue(offsets, text, "warning", base + "/fontSize", "fontSize 应为数值")
    elif t in ("shape", "line"):
        sn = e.get("shapeName")
        if t == "shape" and sn not in SHAPE_NAMES:
            issue(offsets, text, "error", base + "/shapeName",
                  "shapeName 应为 rect/roundRect/oval/donut/line，当前：%r" % (sn,))
        fill_color = (e.get("fill") or {}).get("color")
        if fill_color is not None and t == "shape":
            validate_color(offsets, text, base + "/fill/color", fill_color)
    elif t == "chart":
        rows = (e.get("data") or {}).get("rows")
        if not isinstance(rows, list) or not rows:
            issue(offsets, text, "error", base + "/data/rows", "chart 元素缺少 data.rows（[[分类,值,...],...]）")
    elif t == "table":
        data = e.get("data")
        if not isinstance(data, list):
            issue(offsets, text, "error", base + "/data", "table 元素缺少 data 二维数组")
    specs = COLLECTIONS.get(t)
    if specs:
        for f, kind, limit in specs:
            val = e.get(f)
            ok = (val is not None and (
                (kind == "list" and isinstance(val, list) and len(val) > 0) or
                (kind == "obj" and isinstance(val, dict)) or
                (kind == "str" and isinstance(val, str) and val)))
            if not ok:
                issue(offsets, text, "error", base + "/" + f,
                      "%s 元素缺少 %s（须为%s）" % (t, f, KIND_NAMES[kind]))
            if kind == "list" and limit and isinstance(val, list) and len(val) > limit:
                issue(offsets, text, "warning", base + "/" + f,
                      "%s 条目 %d 个偏多（建议 ≤%d，防溢出）" % (t, len(val), limit))
        if t == "comparison_2col":
            for side in ("left", "right"):
                s = e.get(side)
                if isinstance(s, dict):
                    if not isinstance(s.get("title"), str):
                        issue(offsets, text, "warning", base + "/" + side + "/title",
                              "建议提供字符串 title")
                    pts = s.get("points")
                    if not isinstance(pts, list) or not pts:
                        issue(offsets, text, "error", base + "/" + side + "/points",
                              "%s 缺少 points 列表" % side)


def validate_deck(data, offsets, text):
    _ISSUES.clear()
    if not isinstance(data, dict):
        issue(offsets, text, "error", "", "deck.json 顶层须为对象")
        return list(_ISSUES)
    for key in ("canvas", "theme", "slides"):
        if key not in data:
            issue(offsets, text, "error", "/" + key, "缺少必备字段 %s" % key)
    if not isinstance(data.get("slides"), list) or not data["slides"]:
        issue(offsets, text, "error", "/slides", "slides 须为非空列表")
        return list(_ISSUES)

    canvas = data.get("canvas") or {}
    for k in ("w", "h"):
        if not isinstance(canvas.get(k), (int, float)):
            issue(offsets, text, "error", "/canvas/" + k, "缺少数值型画布 %s" % k)

    colors = (data.get("theme") or {}).get("colors") or {}
    for k in ("$bg", "$text"):
        if k not in colors:
            issue(offsets, text, "error", "/theme/colors", "主题色板缺少 %s（必填）" % k)

    seen_ids = set()
    for i, slide in enumerate(data["slides"]):
        label = "/slides[%d]" % i
        if not isinstance(slide, dict):
            issue(offsets, text, "error", label, "幻灯片须为对象")
            continue
        sid = slide.get("id")
        if not isinstance(sid, str):
            issue(offsets, text, "error", label + "/id", "幻灯片缺少字符串 id")
        elif sid in seen_ids:
            issue(offsets, text, "error", label + "/id", "幻灯片 id %r 重复" % sid)
        seen_ids.add(sid)
        if not isinstance(slide.get("title"), str):
            issue(offsets, text, "warning", label + "/title", "建议提供字符串 title")
        bg = slide.get("background") or {}
        if "color" in bg:
            validate_color(offsets, text, label + "/background/color", bg["color"])
        for j, e in enumerate(slide.get("elements") or []):
            validate_element(offsets, text, "%s/elements[%d]" % (label, j), e)
    return list(_ISSUES)


def main(argv=None):
    parser = argparse.ArgumentParser(description="ppt-studio deck.json schema 校验")
    parser.add_argument("input", help="deck.json 路径")
    parser.add_argument("--strict", action="store_true", help="存在警告时退出码 1")
    args = parser.parse_args(argv)

    src = Path(args.input)
    if not src.is_file():
        print("error: 找不到 " + str(src), file=sys.stderr)
        return 2
    text = src.read_text(encoding="utf-8")
    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        print("error: JSON 解析失败：%s（第 %d 行第 %d 列）" % (exc.msg, exc.lineno, exc.colno),
              file=sys.stderr)
        return 1

    offsets = build_offset_map(text)
    print("ppt-studio deck.json 校验：%s" % src)
    issues = validate_deck(data, offsets, text)
    errs = [i for i in issues if i[0] == "error"]
    warns = [i for i in issues if i[0] == "warning"]
    print("-" * 56)
    print("结果：%d 个错误，%d 个警告" % (len(errs), len(warns)))
    if errs:
        return 1
    if warns and args.strict:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
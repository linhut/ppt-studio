#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
# https://github.com/linhut/ppt-studio
# Licensed under the MIT License. See the LICENSE file for details.

"""ppt-studio PPTX → deck.json 反向导入（ppt pptx2json）。

把原生 PPTX 的矢量结构（文本框/形状/表格/图表）提取为 deck.json，
实现「在 PowerPoint 里改模板 → 回填进 JSON 引擎」的回环工作流。
图片不复制二进制，仅保留原文件名占位（仓库不同步二进制）。

用法：
  ppt pptx2json my-deck/deck.pptx -o my-deck/deck.json
"""

import argparse
import json
import sys
from pathlib import Path

CANVAS_W, CANVAS_H = 960, 540

PPTX_CANVAS = {
    "w": CANVAS_W,
    "h": CANVAS_H,
}
PALETTE = {
    "$bg": "#FFFFFF",
    "$paper": "#F5F8FC",
    "$primary": "#1A3C8B",
    "$accent": "#E67733",
    "$ink": "#1A3C8B",
    "$text": "#222222",
    "$muted": "#666666",
    "$line": "#D0D8E4",
    "$green": "#188050",
    "$brass": "#E67733",
}

SHAPE_MAP = {
    "ROUNDED_RECTANGLE": "roundRect",
    "OVAL": "oval",
    "DIAGONAL_STRIPE": "rect",
}


def _pt(value, total, canvas):
    """EMU → pt（按画布比例缩放）。"""
    return round(value / total * canvas, 1)


def _color(rgb):
    if rgb is None:
        return None
    return "#" + str(rgb)


def _run_color(run):
    try:
        return _color(run.font.color.rgb)
    except Exception:
        return None


def _font_size(run):
    try:
        size = run.font.size
        if size is None:
            return None
        # python-pptx 返回 EMU；1pt = 12700 EMU
        return int(round(size / 12700))
    except Exception:
        return None


def _align(paragraph):
    try:
        return {0: "left", 1: "center", 2: "right", 3: "justify"}.get(
            paragraph.alignment, "left")
    except Exception:
        return "left"


def _text_of(tf, bounds, slide_w, slide_h, title_hint=False):
    """文本框 → JSON text 元素；title_hint 时优先取最大字号段作页标题候选。"""
    paras = [p for p in tf.paragraphs if p.text.strip()]
    if not paras:
        return None
    text = "\n".join(p.text for p in paras)
    runs = [r for p in paras for r in p.runs if r.text.strip()]
    size = max((_font_size(r) or 18 for r in runs), default=18)
    color = next((c for c in (_run_color(r) for r in runs) if c), "$text")
    bold = next((r.font.bold for r in runs if r.font.bold), False)
    x, y = _pt(bounds.left, slide_w, CANVAS_W), _pt(bounds.top, slide_h, CANVAS_H)
    w, h = _pt(bounds.width, slide_w, CANVAS_W), _pt(bounds.height, slide_h, CANVAS_H)
    # 白板空白文字框（短占位）在导入时可保留
    elem = {"type": "text",
            "x": x, "y": y, "w": w, "h": h,
            "text": text, "fontSize": size}
    if color:
        elem["color"] = color
    if bold:
        elem["bold"] = True
    elem["align"] = _align(paras[0])
    if title_hint:
        elem["role"] = "title"
    return elem


def _fill_color(shape):
    try:
        return _color(shape.fill.fore_color.rgb)
    except Exception:
        return None


def _shape_of(shape, slide_w, slide_h, warnings):
    x = _pt(shape.left, slide_w, CANVAS_W)
    y = _pt(shape.top, slide_h, CANVAS_H)
    w = _pt(shape.width, slide_w, CANVAS_W)
    h = _pt(shape.height, slide_h, CANVAS_H)
    if w < 2 and h < 2:
        return None
    name = SHAPE_MAP.get(str(shape.shape_type), None)
    if name is None:
        return None
    side = min(w, h)
    if side > 0 and abs(w - h) / side < 0.02 and name == "rect":
        name = "roundRect" if getattr(shape, "adjustments", None) is not None and len(
            shape.adjustments) else "rect"

    elem = {"type": "shape", "shapeName": name,
            "x": x, "y": y, "w": w, "h": h}
    color = _fill_color(shape)
    if color:
        elem["fill"] = {"color": color}
    return elem


def _line_of(shape, slide_w, slide_h, warnings):
    return {"type": "line",
            "x": _pt(shape.left, slide_w, CANVAS_W),
            "y": _pt(shape.top, slide_h, CANVAS_H),
            "w": _pt(shape.width, slide_w, CANVAS_W),
            "h": _pt(shape.height, slide_h, CANVAS_H)}


def _table_of(shape, slide_w, slide_h, warnings):
    tbl = shape.table
    data = [[cell.text for cell in row.cells] for row in tbl.rows]
    return {"type": "table",
            "x": _pt(shape.left, slide_w, CANVAS_W),
            "y": _pt(shape.top, slide_h, CANVAS_H),
            "w": _pt(shape.width, slide_w, CANVAS_W),
            "h": _pt(shape.height, slide_h, CANVAS_H),
            "rows": len(tbl.rows), "cols": len(tbl.columns),
            "header_color": "$ink",
            "data": data}


def _chart_of(shape, slide_w, slide_h, warnings):
    try:
        chart = shape.chart
        plot = chart.plots[0]
        cats = list(plot.categories) if plot.categories is not None else []
        rows = []
        for i, cat in enumerate(cats):
            row = [str(cat)]
            for series in plot.series:
                try:
                    row.append(series.values[i] if i < len(series.values) else 0)
                except Exception:
                    row.append(0)
            rows.append(row)
        if not rows:
            return None
        return {"type": "chart",
                "x": _pt(shape.left, slide_w, CANVAS_W),
                "y": _pt(shape.top, slide_h, CANVAS_H),
                "w": _pt(shape.width, slide_w, CANVAS_W),
                "h": _pt(shape.height, slide_h, CANVAS_H),
                "color": "$green",
                "data": {"rows": rows}}
    except Exception as exc:
        warnings.append("图表提取失败：%s" % exc)
        return None


def _image_of(shape, slide_w, slide_h, warnings):
    try:
        src = shape.image.filename or "image.png"
    except Exception:
        src = "image.png"
    return {"type": "image",
            "x": _pt(shape.left, slide_w, CANVAS_W),
            "y": _pt(shape.top, slide_h, CANVAS_H),
            "w": _pt(shape.width, slide_w, CANVAS_W),
            "h": _pt(shape.height, slide_h, CANVAS_H),
            "src": "<请提供图片: %s>" % src}


def _slide_of(slide, index, slide_w, slide_h, warnings, title_hint):
    elements = []
    title_text = None
    for shape in slide.shapes:
        try:
            st = shape.shape_type
        except Exception:
            continue
        if str(st) == "GROUP (6)":
            warnings.append("第 %d 页含组合对象，已跳过（请手动展开）" % (index + 1))
            continue
        if shape.has_table:
            elements.append(_table_of(shape, slide_w, slide_h, warnings))
        elif str(st) == "CHART (3)":
            elem = _chart_of(shape, slide_w, slide_h, warnings)
            if elem:
                elements.append(elem)
        elif str(st) == "PICTURE (13)":
            elements.append(_image_of(shape, slide_w, slide_h, warnings))
        elif str(st) == "LINE (9)":
            elements.append(_line_of(shape, slide_w, slide_h, warnings))
        elif shape.has_text_frame and shape.text_frame.text.strip():
            is_title = shape == slide.shapes.title
            elem = _text_of(shape.text_frame, shape, slide_w, slide_h, is_title)
            if elem and elem["text"].strip():
                elements.append(elem)
                if title_text is None and elem.get("role") == "title":
                    title_text = elem["text"].strip()[:20]
        else:
            elem = _shape_of(shape, slide_w, slide_h, warnings)
            if elem:
                elements.append(elem)

    if not elements:
        warnings.append("第 %d 页无可导入元素" % (index + 1))
    if title_text is None and elements:
        title_text = "幻灯片 %d" % (index + 1)
    return {
        "id": "%02d" % (index + 1),
        "title": title_text or ("幻灯片 %d" % (index + 1)),
        "elements": elements,
    }


def convert(pptx_path):
    from pptx import Presentation
    prs = Presentation(str(pptx_path))
    sw, sh = prs.slide_width, prs.slide_height
    warnings = []
    slides = [_slide_of(s, i, sw, sh, warnings, False) for i, s in enumerate(prs.slides)]
    deck = {
        "_doc": "pptx2json 从 %s 导入（%d 页）；图片为占位路径，需手动替换"
                % (pptx_path.name, len(slides)),
        "canvas": {"w": CANVAS_W, "h": CANVAS_H},
        "theme": {"colors": dict(PALETTE)},
        "slides": slides,
    }
    return deck, warnings


def main(argv=None):
    parser = argparse.ArgumentParser(description="ppt-studio PPTX → deck.json 反向导入")
    parser.add_argument("input", help="deck.pptx 路径")
    parser.add_argument("-o", "--output", default=None, help="输出 deck.json 路径")
    args = parser.parse_args(argv)

    src = Path(args.input)
    if not src.is_file():
        print("error: 找不到 " + str(src), file=sys.stderr)
        return 2
    deck, warnings = convert(src)
    out = Path(args.output) if args.output else (src.parent / "deck.json")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(deck, ensure_ascii=False, indent=2), encoding="utf-8")
    print("pptx2json 导入 %d 页 → %s" % (len(deck["slides"]), out))
    for w in warnings:
        print("  ! " + w, file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
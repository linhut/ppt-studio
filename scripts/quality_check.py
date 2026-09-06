#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# (c) 2026 Jose AI (https://www.linhut.cn)
# https://github.com/linhut/ppt-studio
# Licensed under the MIT License. See the LICENSE file for details.

"""ppt-studio 质量审查：对 JSON 数据驱动 deck 执行 5 维度检查 + 硬约束校验。

借鉴 harness-anything（yb2460/harness-anything，MIT）quality_checks.py：
- validate_slide() 逐页打分与警告
- review_deck() 整体评分
- 5 维度：visual(70) / pedagogy(75) / proofreading(80) / parity(85) / substance(90)

自动可查（机器执行）：
  visual       字体层级、颜色数、视觉占比、一页一主题
  hard         元素越界、标题-内容间距、标题字数、文本密度（硬约束）
人工复核（auto_only=False 时输出 checklist）：
  pedagogy     叙事弧、预备知识、示例
  proofreading 拼写、术语、溢出
  parity       PPTX/PDF 一致性
  substance    数据准确性、引用

用法：
  py -3 scripts/quality_check.py my-deck/deck.json
  py -3 scripts/quality_check.py my-deck/deck.json --preset academic --json
"""

import argparse
import json
import math
import re
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
from presets import get_preset  # noqa: E402

DIMENSIONS = [
    ("visual",       "视觉审查",   70),
    ("hard",         "硬约束审查",  60),
    ("pedagogy",     "教学法审查",  75),
    ("proofreading", "校对审查",    80),
    ("parity",       "一致性审查",  85),
    ("substance",    "实质内容审查", 90),
]

COMPOSITE_TYPES = {"cards_2x3", "cards_1x4_info", "card_list_wide",
                   "tagline_bar", "num_big", "figure_text", "breadcrumb", "references"}
VISUAL_TYPES = {"image", "chart", "table", "shape", "rect", "circle", "line"} | COMPOSITE_TYPES


def plain_text(t):
    s = re.sub(r"<[^>]+>", "", str(t.get("text", "")))
    return s.replace("\n", " ").strip()


def cn_density(texts):
    total = 0
    for t in texts:
        s = re.sub(r"<[^>]+>", "", str(t.get("text", "")))
        total += len(re.sub(r"\s", "", s))
    return total


# ---- 容量精算（借鉴 GordenPPTSkill scripts/compute_capacity.py，MIT）----
# 视觉宽度单位 vw：CJK/全角=1.0、空格=0.35、ASCII 拉丁/数字=0.5、其他=0.8
# 1 vw ≈ 1 em ≈ fontSize pt。每行 vw 容量 = 可用宽/字号；
# 最大行数 = 可用高/(字号×行高)；留 20% 容差防误报。

def vw_of(s):
    w = 0.0
    for c in s:
        if "\u4e00" <= c <= "\u9fff" or "\u3000" <= c <= "\u303f" or "\uff00" <= c <= "\uffef":
            w += 1.0          # 中日韩全角
        elif c == " ":
            w += 0.35
        elif c.isascii():
            w += 0.5          # 拉丁/数字/半角标点
        else:
            w += 0.8
    return w


def split_lines(text):
    """按 <br/> 与换行符分行，去掉 HTML 标签。"""
    t = re.sub(r"<br\s*/?>", "\n", str(text))
    t = re.sub(r"<[^>]+>", "", t)
    return t.split("\n")


def capacity_check(text, w_pt, h_pt, size_pt, wrap=True,
                   line_height=1.0, tolerance=1.2):
    """返回 (need_lines, max_lines, content_vw, cap_vw, ok)。"""
    if size_pt <= 0:
        size_pt = 14.0
    h_margin = 8.0      # 左右内边距合计（pt）
    v_margin = 4.0      # 上下内边距合计（pt）
    usable_w = max(0.0, w_pt - h_margin)
    usable_h = max(0.0, h_pt - v_margin)
    cpl = max(1, int(usable_w // size_pt))          # 每行 vw 容量
    if not wrap:
        max_lines = 1
    else:
        max_lines = max(1, int(usable_h // (size_pt * line_height)))
    cap = max(1, int(cpl * max_lines * tolerance))  # 容量预算（含容差）

    lines = split_lines(text)
    need_lines = 0
    used_vw = 0.0
    for ln in lines:
        if not ln:
            need_lines += 1
            continue
        lw = vw_of(ln)
        need_lines += max(1, int(math.ceil(lw / cpl)))
        used_vw += lw
    ok = need_lines <= max_lines
    return need_lines, max_lines, used_vw, cap, ok


def validate_capacity(slide, cw=960, ch=540):
    """容量精算：检查每个 text 元素是否超出文本框可视容量（防溢出遮挡）。"""
    warnings = []
    score = 100.0
    for e in slide.get("elements", []):
        if e.get("type") != "text":
            continue
        w = float(e.get("w", 0)); h = float(e.get("h", 0))
        if w <= 0 or h <= 0:
            continue
        size = float(e.get("fontSize", e.get("fs", 14)))
        text = e.get("text", "")
        if not str(text).strip():
            continue
        need, mx, used, cap, ok = capacity_check(text, w, h, size)
        if not ok:
            # 文字块宽超过一行宽度时按折行计；仅当需求行数超过容量行数才警告
            warnings.append(
                "容量溢出：'%s' 需 %d 行，文本框最多 %d 行（%.0fpt 字号，宽%.0f 高%.0f）"
                % (str(text)[:14], need, mx, size, w, h))
            score -= 5
    return {"score": max(0, min(100, score)), "warnings": warnings,
            "pass": score >= 70}


def validate_slide(slide, rules, cw=960, ch=540):
    """逐页自动审查。返回 {"score": 0-100, "warnings": [...], "pass": bool}。"""
    warnings = []
    score = 100.0
    elements = slide.get("elements", [])

    texts = [e for e in elements if e.get("type") == "text"]
    # 复合组件内部含大量图形元素，一并计入视觉，避免"纯文本页"误报
    COMPOSITE_TYPES = ("cards_2x3", "cards_1x4_info", "card_list_wide",
                       "tagline_bar", "num_big", "section_divider",
                       "comparison_2col", "timeline_h", "process_steps",
                       "kpi_row", "pros_cons", "big_quote", "roadmap_4col",
                       "checklist", "cover_asym", "figure_text", "breadcrumb",
                       "references")
    visuals = [e for e in elements
               if e.get("type") in VISUAL_TYPES or e.get("type") in COMPOSITE_TYPES]

    # ---- 字体层级 ----
    # role → 默认字号映射（与 json2pptx.py FS_* 常量对齐，避免误报）
    ROLE_FS = {"cover": 54, "title": 40, "h3": 24, "h4": 18, "body": 16,
               "muted": 13, "caption": 13, "eyebrow": 12, "kicker": 13,
               "stat": 44, "card_desc": 14, "table_head": 13,
               "table_cell": 12, "tagline": 15}
    for t in texts:
        role = t.get("role", "body")
        _fs = t.get("fontSize", t.get("fs", ROLE_FS.get(role, 16)))
        fs = float(_fs if _fs else ROLE_FS.get(role, 16))
        text = plain_text(t)
        if role == "title":
            if fs < 36:
                warnings.append("标题字号 %.0fpt < 36pt（正文页标题建议 36-44pt）" % fs)
                score -= 5
            limit = int(rules.get("title_max_len", 12))
            if len(text) > limit:
                warnings.append("标题 %d 字 > 建议 %d 字" % (len(text), limit))
                score -= 3
        elif role == "body":
            if fs < 14:  # design-system.md：正文 14pt（html-ppt body 16px 映射）
                warnings.append("正文字号 %.0fpt < 14pt" % fs)
                score -= 3
        elif role == "caption":
            if fs < 12 and fs > 0:  # design-system.md：弱化/标注 12pt
                warnings.append("标注字号 %.0fpt < 12pt" % fs)
                score -= 2
        if fs <= 0:
            warnings.append("文本缺少字号：" + text[:12])
            score -= 2

    # ---- 一页一主题 ----
    titles = [t for t in texts if t.get("role") == "title"]
    if len(titles) > 1:
        warnings.append("检测到 %d 个标题，建议一页一个主题" % len(titles))
        score -= 10

    # ---- 内容密度 ----
    max_bullets = int(rules.get("max_bullets", 6))
    if len(texts) > max_bullets * 1.5:
        warnings.append("文本块过多（%d > %d）" % (len(texts), int(max_bullets * 1.5)))
        score -= 8
    density = cn_density(texts)
    if density > 80:
        warnings.append("单页中文 %d 字 > 建议 80 字" % density)
        score -= 6

    # ---- 视觉占比 ----
    target = float(rules.get("visual_ratio", 0.5))
    total = len(texts) + len(visuals)
    if total > 0:
        ratio = len(visuals) / float(total)
        if ratio < target - 0.2:
            warnings.append("视觉占比 %.0f%% < 目标 %.0f%%" % (ratio * 100, target * 100))
            score -= 5

    # ---- 颜色数 ----
    used = set()
    for e in elements:
        c = e.get("color") or (e.get("fill") or {}).get("color")
        if c:
            used.add(str(c))
    max_colors = int(rules.get("max_colors", 6))
    if len(used) > max_colors + 3:
        warnings.append("颜色使用过多（%d 种 > %d）" % (len(used), max_colors))
        score -= 4

    return {"score": max(0, min(100, score)), "warnings": warnings,
            "pass": score >= 70}


def validate_hard(slide, cw=960, ch=540):
    """硬约束检查：越界、标题-内容间距。"""
    warnings = []
    score = 100.0
    elements = slide.get("elements", [])
    title_bottom = None

    for e in elements:
        if "x" in e and "y" in e and "w" in e:
            w = float(e.get("w", 0)); h = float(e.get("h", 0))
            x = float(e.get("x", 0)); y = float(e.get("y", 0))
            if x + w > cw + 1:
                warnings.append("元素越界：x=%.0f + w=%.0f > 画布宽 %d" % (x, w, cw))
                score -= 5
            if y + h > ch + 1:
                warnings.append("元素越界：y=%.0f + h=%.0f > 画布高 %d" % (y, h, ch))
                score -= 5
            if x < 0 or y < 0:
                warnings.append("元素负坐标：x=%.0f y=%.0f" % (x, y))
                score -= 3
            if e.get("role") == "title":
                title_bottom = y + h

    if title_bottom is not None:
        for e in elements:
            if e.get("role") != "title" and "y" in e and "h" in e:
                y = float(e.get("y")); h = float(e.get("h"))
                # 跳过装饰细条（高 ≤6pt 的横线/细色块）与眉题：不计入内容间距
                if h <= 6:
                    continue
                # 只检查位于标题下方的内容元素；标题上方的装饰（圆形/色块）不计入
                if y > title_bottom and y - title_bottom < 20:
                    warnings.append("标题与内容间距 %.0fpt < 24pt（建议 ≥24pt）"
                                    % (y - title_bottom))
                    score -= 4
                    break
    return {"score": max(0, min(100, score)), "warnings": warnings,
            "pass": score >= 70}


def review_deck(data, preset):
    """整个 deck 审查。"""
    slides = data.get("slides", [])
    cw = float((data.get("canvas") or {}).get("w", 960))
    ch = float((data.get("canvas") or {}).get("h", 540))
    rules = preset.rules if preset else {}
    per_slide = []
    warnings_total = 0
    auto_scores = []

    for i, slide in enumerate(slides):
        r1 = validate_slide(slide, rules, cw, ch)
        r2 = validate_hard(slide, cw, ch)
        r3 = validate_capacity(slide, cw, ch)
        score = int(round((r1["score"] + r2["score"] + r3["score"]) / 3))
        warns = list(r1["warnings"]) + list(r2["warnings"]) + list(r3["warnings"])
        warnings_total += len(warns)
        per_slide.append({"slide_index": i + 1,
                          "id": slide.get("id", ""),
                          "title": slide.get("title", ""),
                          "score": score,
                          "pass": score >= 70,
                          "warnings": warns})
        auto_scores.append(score)

    overall = int(round(sum(auto_scores) / max(len(auto_scores), 1)))
    return {
        "overall_score": overall,
        "dimensions": [{"key": k, "name": n, "threshold": th,
                        "auto": k in ("visual", "hard")}
                       for k, n, th in DIMENSIONS],
        "per_slide": per_slide,
        "total_warnings": warnings_total,
        "summary": "整体 %d 分，%d 页，%d 条警告；%s"
                   % (overall, len(slides), warnings_total,
                      "通过" if overall >= 70 else "未达 70 分，请修复警告"),
    }


def main(argv=None):
    parser = argparse.ArgumentParser(description="ppt-studio JSON deck 质量审查")
    parser.add_argument("input", type=Path, help="deck.json 路径")
    parser.add_argument("--preset", default="academic", help="设计预设名（academic/consultant/business/tech 或 18 配色 key）")
    parser.add_argument("--json", action="store_true", help="输出 JSON 报告")
    parser.add_argument("--auto-only", action="store_true", help="只跑自动检查，跳过人工复核清单")
    args = parser.parse_args(argv)

    if not args.input.is_file():
        print("error: 找不到 " + str(args.input), file=sys.stderr)
        return 2
    data = json.loads(args.input.read_text(encoding="utf-8"))
    preset = get_preset(args.preset)
    if preset is None:
        print("warning: 未找到预设 %s，使用默认规则" % args.preset, file=sys.stderr)

    report = review_deck(data, preset)

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(report["summary"])
        if not args.auto_only:
            print("\n逐页警告：")
            for p in report["per_slide"]:
                if p["warnings"]:
                    print("  [%s] %s (%d分)" % (p["id"], p["title"], p["score"]))
                    for w in p["warnings"]:
                        print("    - " + w)
            print("\n人工复核清单（无法自动判定，请逐项确认）：")
            for k, n, th in DIMENSIONS:
                if k not in ("visual", "hard"):
                    print("  [ ] %s（阈值 %d 分）" % (n, th))
        else:
            print("（--auto-only 已跳过人工复核清单）")
        if report["overall_score"] < 70:
            return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

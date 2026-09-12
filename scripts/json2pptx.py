#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
# https://github.com/linhut/ppt-studio
# Licensed under the MIT License. See the LICENSE file for details.

"""ppt-studio JSON 数据驱动引擎 v3：deck.json → 原生 PPTX。

渲染遵循 references/design-system.md（三源融合设计规范）：
- 字号层级：html-ppt base.css 真实值映射（title 36 / h3 24 / body 14 / muted 12 / stat 44）
- 间距/圆角：html-ppt 真实 token 映射（页边距 40-48 / 卡片 gap 24 / padding 16-18 / radius 12）
- 配色：MiniMax 18 套成品 → 主题变量
- 引擎：np-ppt python-pptx 内核思路；路由架构：harness-anything

JSON 结构：
{
  "canvas": {"w": 960, "h": 540},
  "theme": {"colors": {...}, "textStyles": {...}},
  "slides": [{"id": "01", "title": "封面", "elements": [ ... ]}]
}

元素类型：text / shape(rect|roundRect|oval|donut) / line / circle / image /
table / chart / cards_2x3 / cards_1x4_info / card_list_wide / tagline_bar / num_big /
section_divider / comparison_2col / pros_cons / timeline_h / process_steps / kpi_row /
big_quote / roadmap_4col / checklist / cover_asym / figure_text / breadcrumb / references（布局速查见 references/layout-catalog.md）

特性：
- $theme 变量在颜色字段直接可用（无 55173 内联限制）；背景支持纯色与渐变
- 淡入淡出过渡默认开启；--no-fade 关闭
- 越界 + 元素重叠检测（警告，--strict 失败）

用法：
  py -3 scripts/json2pptx.py my-deck/deck.json --output my-deck/deck.pptx
  py -3 scripts/json2pptx.py my-deck/deck.json -o out.pptx --no-fade --strict
"""

import argparse
import json
import re
import sys
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.util import Pt, Emu
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.dml.color import RGBColor
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
    from pptx.oxml.ns import nsdecls, qn
    from pptx.oxml import parse_xml
except ImportError:
    print("缺少依赖：pip install python-pptx", file=sys.stderr)
    sys.exit(2)

# ============================================================
# 三源融合设计常量（来源见 references/design-system.md）
# ============================================================
FONT_CN = "微软雅黑"
FONT_EN = "Arial"              # MiniMax 规范：中文雅黑 / 英文 Arial
FONT_STAT = "Arial"

# 字号层级（对齐 harness-anything 大字号体系：H1 40-48 / Body 18-20 / Caption 14-15）
# 参考：harness-anything 的 WPS 自动化工作流文档第七节（字号层级对齐思路）
FS_COVER = 54                  # 封面主标题（h1）
FS_TITLE = 40                  # 内容页标题（harness H1 40-48）
FS_H3 = 24                     # 小节标题（harness H2 24-26）
FS_H4 = 18                     # 卡片标题（harness H3 20-22 → 卡片取 18）
FS_BODY = 16                   # 正文（harness Body 18-20 → 正文取 16）
FS_MUTED = 13                  # 弱化/注释（harness Caption 14-15 → 取 13）
FS_EYEBROW = 12                # 眉题
FS_KICKER = 13                 # 引导词
FS_STAT = 44                   # kpi 大数字（harness Number 28-36 → 大屏取 44）
FS_CARD_DESC = 14              # 卡片描述（harness Body-S 16-17 → 卡片取 14）
FS_TABLE_HEAD = 13             # 表头（harness Table 13-15）
FS_TABLE_CELL = 12             # 单元格（harness Table 13-15）
FS_TAGLINE = 15                # 底部总结条（harness Tagline 14）

# 间距 / 圆角（html-ppt base.css → pt）
MARGIN_X = 48                  # 左右页边距（96px）
MARGIN_Y = 40                  # 上下页边距（72px）
GAP_CARD = 24                  # 卡片/列间距（grid gap 24px）
GAP_STACK = 14                 # 块内间距（stack 14px）
PAD_CARD_X = 18                # 卡片内边距（28px）
PAD_CARD_Y = 16                # 卡片内边距（26px）
RADIUS_CARD = 12               # 卡片圆角（18px）
RADIUS_SM = 8                  # 小圆角（12px）
CARD_BAR = 3                   # 卡片顶部色条高（3px）
DIVIDER_H = 3                  # 装饰线高（3px）
TITLE_TOP = 14                 # 内容页标题 y
TITLE_H = 46                   # 内容页标题高（40pt 字号需 46pt 框）
CONTENT_MIN_Y = 84             # 标题下第一个内容元素最小 y（间距 ≥24pt）
LINE_HEIGHT = 1.4              # 正文行高
FALLBACK_RGB = (31, 42, 36)


def log(msg):
    print("[json2pptx] " + msg, file=sys.stderr, flush=True)


def resolve_color(val, colors):
    """解析 #RRGGBB / $var / rgb(r,g,b) → 6 位 hex 大写。"""
    if not val:
        return None
    s = str(val).strip()
    if s.startswith("$"):
        s = colors.get(s[1:], "") or colors.get(s, "")
    if not s:
        return None
    m = re.search(r"rgb[a]?\s*\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)", s, re.I)
    if m:
        return "%02X%02X%02X" % (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    h = s.lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    if len(h) >= 6:
        return h[:6].upper()
    return None


def rgb_of(hex6):
    try:
        return RGBColor(int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16))
    except Exception:
        return RGBColor(*FALLBACK_RGB)


def set_fade(slide):
    try:
        trans = parse_xml(
            '<p:transition %s spd="fast"><p:fade/></p:transition>' % nsdecls("p")
        )
        slide._element.append(trans)
    except Exception:
        pass


class Builder:
    def __init__(self, data, base_dir):
        self.data = data
        self.base_dir = base_dir
        self.colors = data.get("theme", {}).get("colors", {})
        self.text_styles = data.get("theme", {}).get("textStyles", {})
        self.fonts = data.get("theme", {}).get("fonts", {})
        canvas = data.get("canvas", {})
        self.cw = float(canvas.get("w", 960))
        self.ch = float(canvas.get("h", 540))
        self.warnings = []
        self.prs = Presentation()
        self.prs.slide_width = Emu(int(self.cw * 12700))
        self.prs.slide_height = Emu(int(self.ch * 12700))

    # ---------- 工具 ----------
    def col(self, v, default=None):
        rgb = resolve_color(v, self.colors)
        if rgb:
            return rgb
        return default or "1F2A24"

    def emu(self, v):
        return Emu(int(float(v) * 12700))

    def text_style(self, name):
        return self.text_styles.get(name, {}) if name else {}

    def font_for(self, role, override=None):
        """字体解析：元素 font > theme.fonts[角色组] > 默认（中文雅黑 / 英文 Arial）。"""
        if override:
            return override
        group = ("stat" if role == "stat" else
                 ("heading" if role in ("title", "h3", "h4", "kicker") else "body"))
        if group == "stat":
            return self.fonts.get("stat", FONT_STAT)
        return self.fonts.get(group, FONT_CN)

    def add_box(self, slide, x, y, w, h):
        return slide.shapes.add_textbox(self.emu(x), self.emu(y), self.emu(w), self.emu(h))

    def put_text(self, tf, text, fs, color, bold=False, align=PP_ALIGN.LEFT,
                 font=FONT_CN, line_spacing=LINE_HEIGHT):
        """多行文本写入 text_frame。text 支持 \n 与简单 <br/>。"""
        text = re.sub(r"<[^>]+>", "", str(text))
        parts = re.split(r"\n|<br\s*/?>", text)
        tf.word_wrap = True
        tf.margin_left = Pt(0); tf.margin_right = Pt(0)
        tf.margin_top = Pt(0); tf.margin_bottom = Pt(0)
        for i, part in enumerate(parts):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.line_spacing = line_spacing
            run = p.add_run()
            run.text = part
            run.font.size = Pt(fs)
            run.font.bold = bold
            run.font.name = font
            run.font.color.rgb = rgb_of(color)

    # ---------- 形状 ----------
    def add_shape(self, slide, shape_name, e):
        x, y, w, h = e["x"], e["y"], e["w"], e.get("h", e["w"])
        mapping = {"rect": MSO_SHAPE.RECTANGLE, "roundRect": MSO_SHAPE.ROUNDED_RECTANGLE,
                   "oval": MSO_SHAPE.OVAL, "donut": MSO_SHAPE.DONUT,
                   "line": MSO_SHAPE.RECTANGLE}
        shp = slide.shapes.add_shape(mapping.get(shape_name, MSO_SHAPE.RECTANGLE),
                                     self.emu(x), self.emu(y), self.emu(w), self.emu(h))
        fill = e.get("fill", {})
        if fill.get("color"):
            shp.fill.solid()
            shp.fill.fore_color.rgb = rgb_of(self.col(fill.get("color"), "$paper"))
        else:
            shp.fill.background()
        border = e.get("border", {})
        if border:
            try:
                shp.line.color.rgb = rgb_of(self.col(border.get("color"), "$line"))
                shp.line.width = Pt(float(border.get("width", 1)))
            except Exception:
                pass
        else:
            shp.line.fill.background()
        # 圆角半径（roundRect 调整）
        if shape_name == "roundRect":
            try:
                adj = shp.adjustments[0]
                adj.value = float(e.get("rectRadius", 0.08))
            except Exception:
                pass
        return shp

    # ---------- 文本 ----------
    def add_text(self, slide, e):
        """按角色给默认字号（design-system 层级），元素可覆盖。"""
        role = e.get("role", "body")
        default_fs = {"title": FS_TITLE, "h3": FS_H3, "h4": FS_H4,
                      "body": FS_BODY, "caption": FS_MUTED,
                      "muted": FS_MUTED, "kicker": FS_KICKER,
                      "eyebrow": FS_EYEBROW, "stat": FS_STAT}.get(role, FS_BODY)
        style = self.text_style(e.get("style"))
        fs = float(e.get("fontSize", e.get("fs", style.get("fontSize", default_fs))))
        color = e.get("color", style.get("color", "$text" if role not in ("kicker", "stat") else "$accent"))
        align_s = e.get("align", "left")
        align_map = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
        box = self.add_box(slide, e["x"], e["y"], e["w"], e.get("h", 40))
        tf = box.text_frame
        valign = e.get("valign", "top")
        tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
                              "bottom": MSO_ANCHOR.BOTTOM}.get(valign, MSO_ANCHOR.TOP)
        self.put_text(tf, e.get("text", ""), fs, self.col(color),
                      bold=bool(e.get("bold", role in ("title", "h3", "h4", "stat", "kicker"))),
                      align=align_map.get(align_s, PP_ALIGN.LEFT),
                      font=self.font_for(role, e.get("font")),
                      line_spacing=float(e.get("lineHeight", style.get("lineHeight", LINE_HEIGHT))))
        return box

    # ---------- 图片 / 表格 / 图表 ----------
    def add_image(self, slide, e):
        fp = Path(e.get("file", ""))
        if not fp.is_absolute():
            fp = self.base_dir / fp
        if not fp.is_file():
            self.warn("图片不存在：" + str(fp))
            return
        slide.shapes.add_picture(str(fp), self.emu(e["x"]), self.emu(e["y"]),
                                 self.emu(e["w"]), self.emu(e["h"]))

    def add_table(self, slide, e):
        rows, cols = e.get("rows"), e.get("cols")
        data = e.get("data", [])
        if not rows or not cols:
            return
        x, y, w, h = e["x"], e["y"], e["w"], e.get("h", 300)
        gf = slide.shapes.add_table(rows, cols, self.emu(x), self.emu(y),
                                    self.emu(w), self.emu(h))
        table = gf.table
        hdr = self.col(e.get("header_color", "$ink"), "1F2A24")
        cell_font = self.font_for("body") if not self.fonts.get("table") else self.fonts["table"]
        for r in range(rows):
            for c in range(cols):
                cell = table.cell(r, c)
                val = data[r][c] if r < len(data) and c < len(data[r]) else ""
                cell.text = str(val)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.margin_left = Pt(8); cell.margin_right = Pt(8)
                cell.margin_top = Pt(2); cell.margin_bottom = Pt(2)
                for p in cell.text_frame.paragraphs:
                    p.alignment = PP_ALIGN.CENTER if c > 0 or r == 0 else PP_ALIGN.LEFT
                    for run in p.runs:
                        run.font.size = Pt(FS_TABLE_HEAD if r == 0 else FS_TABLE_CELL)
                        run.font.bold = bool(r == 0)
                        run.font.name = cell_font
                        run.font.color.rgb = rgb_of("FFFFFF" if r == 0 else "333333")
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb_of(hdr if r == 0 else ("FFFFFF" if r % 2 == 0 else "F5F0ED"))
                # 不换行
                try:
                    cell.text_frame.word_wrap = False
                except Exception:
                    pass

    def add_chart(self, slide, e):
        rows = e.get("data", {}).get("rows", e.get("rows", []))
        if not rows:
            return
        ctype = str(e.get("chartType", "bar")).lower()
        x, y, w, h = e["x"], e["y"], e["w"], e.get("h", 380)

        def num(v):
            try:
                return float(v)
            except (TypeError, ValueError):
                return None

        if ctype == "scatter":
            from pptx.chart.data import XyChartData
            xd = XyChartData()
            points = []
            for r in rows:
                if len(r) >= 3 and num(r[1]) is not None and num(r[2]) is not None:
                    points.append((num(r[1]), num(r[2])))
            if not points:
                return
            xd.add_series("散点", points)
            gf = slide.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER,
                                        self.emu(x), self.emu(y), self.emu(w), self.emu(h), xd)
        else:
            # 数值列检测：跳过「单位」类非数值列（兼容 [分类, 单位, 值] 旧格式）
            series_spec = []
            for j in range(1, len(rows[0])):
                vals = [num(r[j]) if j < len(r) else None for r in rows]
                if all(v is not None for v in vals):
                    series_spec.append((j, vals))
            if not series_spec:
                return
            if ctype == "pie":
                series_spec = series_spec[:1]
            chart_data = CategoryChartData()
            chart_data.categories = [str(r[0]) for r in rows]
            for idx, (_, vals) in enumerate(series_spec):
                chart_data.add_series("数值" if len(series_spec) == 1 else "系列%d" % (idx + 1),
                                      [v for v in vals])
            xtype = {"line": XL_CHART_TYPE.LINE_MARKERS,
                     "pie": XL_CHART_TYPE.PIE,
                     "radar": XL_CHART_TYPE.RADAR,
                     "stacked": XL_CHART_TYPE.COLUMN_STACKED,
                     }.get(ctype, XL_CHART_TYPE.COLUMN_CLUSTERED)
            gf = slide.shapes.add_chart(xtype,
                                        self.emu(x), self.emu(y), self.emu(w), self.emu(h), chart_data)
        chart = gf.chart
        try:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
            series_colors = ["$green", "$accent", "$primary", "$brass"]
            for idx, series in enumerate(chart.plots[0].series):
                series.format.fill.solid()
                series.format.fill.fore_color.rgb = rgb_of(
                    self.col(e.get("color", series_colors[idx % len(series_colors)])))
            if len(chart.plots[0].series) == 1:
                chart.plots[0].series[0].has_data_labels = True
        except Exception:
            pass

    # ---------- 复合组件（harness-anything 路由 + html-ppt 视觉） ----------
    def add_cards_2x3(self, slide, e):
        """2 行 x 3 列卡片网格（html-ppt toc 卡片风格：顶部 3pt 色条 + 圆角 12 + 内边距 16/18）。"""
        items = e.get("items", [])
        if not items:
            return
        x, y, w, h = e["x"], e["y"], e["w"], e.get("h", 320)
        gap = float(e.get("gap", GAP_CARD))
        cw = (w - 2 * gap) / 3.0
        chh = (h - gap) / 2.0
        for i, item in enumerate(items):
            cx = x + (i % 3) * (cw + gap)
            cy = y + (i // 3) * (chh + gap)
            col_hex = self.col(item.get("color", e.get("card_color", "$primary")), "1F2A24")
            # 顶部色条 3pt
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                         self.emu(cx), self.emu(cy),
                                         self.emu(cw), self.emu(CARD_BAR))
            bar.fill.solid(); bar.fill.fore_color.rgb = rgb_of(col_hex); bar.line.fill.background()
            # 卡片底（圆角 12pt）
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                          self.emu(cx), self.emu(cy),
                                          self.emu(cw), self.emu(chh))
            card.fill.solid(); card.fill.fore_color.rgb = rgb_of(self.col("$paper", "FFFFFF"))
            card.line.color.rgb = rgb_of(self.col("$line", "E5E7EB")); card.line.width = Pt(1)
            try:
                card.adjustments[0].value = 0.08
            except Exception:
                pass
            # 标题 h4=16pt（design-system）
            tb = self.add_box(slide, cx + PAD_CARD_X, cy + PAD_CARD_Y + CARD_BAR, cw - 2 * PAD_CARD_X, 26)
            self.put_text(tb.text_frame, item.get("title", ""), FS_H4,
                          self.col("$text", "222222"), bold=True, font=FONT_CN)
            # 描述 ≤2 行，12pt
            db = self.add_box(slide, cx + PAD_CARD_X, cy + PAD_CARD_Y + CARD_BAR + 26,
                              cw - 2 * PAD_CARD_X, chh - PAD_CARD_Y - CARD_BAR - 30)
            desc = str(item.get("desc", ""))
            if len(desc) > 34:
                desc = desc[:33] + "…"
            self.put_text(db.text_frame, desc, FS_CARD_DESC,
                          self.col("$muted", "888888"), bold=False, font=FONT_CN, line_spacing=1.3)

    def add_cards_1x4_info(self, slide, e):
        """4 列统计卡（html-ppt kpi-grid：eyebrow + 44pt 数字 + 增量标签）。"""
        items = e.get("items", [])[:4]
        if not items:
            return
        x, y, w, h = e["x"], e["y"], e["w"], e.get("h", 130)
        gap = float(e.get("gap", 16))
        cw = (w - 3 * gap) / 4.0
        for i, item in enumerate(items):
            cx = x + i * (cw + gap)
            col_hex = self.col(item.get("color", e.get("accent_color", "$accent")), "1F2A24")
            # 卡片底
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                          self.emu(cx), self.emu(y),
                                          self.emu(cw), self.emu(h))
            card.fill.solid(); card.fill.fore_color.rgb = rgb_of(self.col("$paper", "FFFFFF"))
            card.line.color.rgb = rgb_of(self.col("$line", "E5E7EB")); card.line.width = Pt(1)
            try:
                card.adjustments[0].value = 0.08
            except Exception:
                pass
            # 数字 40pt（占 45%）+ 标签 12pt（占 40%），三段式
            num_h = h * 0.45
            nb = self.add_box(slide, cx + 6, y + 10, cw - 12, num_h)
            self.put_text(nb.text_frame, str(item.get("num", "")),
                          float(item.get("fs", 40)), col_hex,
                          bold=True, align=PP_ALIGN.CENTER, font=FONT_STAT)
            lb = self.add_box(slide, cx + 6, y + num_h + 6, cw - 12, h - num_h - 12)
            self.put_text(lb.text_frame, str(item.get("label", "")), 12,
                          self.col("$muted", "888888"), align=PP_ALIGN.CENTER, font=FONT_CN)

    def add_card_list_wide(self, slide, e):
        """编号圆 + 标题 + 副标题列表（html-ppt toc 行卡片 + MiniMax 编号圆）。"""
        items = e.get("items", [])
        x = float(e.get("x", 100))
        y = float(e.get("start_y", 84))
        ih = float(e.get("item_h", 56))
        colors = [self.col(e.get("color_a", "$primary")), self.col(e.get("color_b", "$accent"))]
        for i, item in enumerate(items):
            cy = y + i * ih
            col_hex = colors[i % 2]
            # 编号圆 28pt
            c = slide.shapes.add_shape(MSO_SHAPE.OVAL, self.emu(x), self.emu(cy + 4),
                                       self.emu(28), self.emu(28))
            c.fill.solid(); c.fill.fore_color.rgb = rgb_of(col_hex); c.line.fill.background()
            nbox = self.add_box(slide, x, cy + 4, 28, 28)
            self.put_text(nbox.text_frame, str(item.get("num", "")), 12,
                          "FFFFFF", bold=True, align=PP_ALIGN.CENTER, font=FONT_EN)
            # 标题 16pt bold
            tbox = self.add_box(slide, x + 44, cy, 560, 24)
            self.put_text(tbox.text_frame, str(item.get("title", "")), FS_H4,
                          self.col("$text", "1A1A1A"), bold=True, font=FONT_CN)
            # 副标题 12pt muted
            sbox = self.add_box(slide, x + 44, cy + 24, 700, 20)
            self.put_text(sbox.text_frame, str(item.get("sub", "")), FS_MUTED,
                          self.col("$muted", "666666"), font=FONT_CN)

    def add_tagline_bar(self, slide, e):
        """底部品牌色总结条。"""
        x = float(e.get("x", self.cw * 0.03))
        y = float(e.get("y", self.ch - 42))
        w = float(e.get("w", self.cw * 0.94))
        h = float(e.get("h", 28))
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, self.emu(x), self.emu(y),
                                     self.emu(w), self.emu(h))
        bar.fill.solid(); bar.fill.fore_color.rgb = rgb_of(self.col(e.get("color", "$primary")))
        bar.line.fill.background()
        tb = self.add_box(slide, x + 10, y + 3, w - 20, h - 6)
        self.put_text(tb.text_frame, e.get("text", ""), FS_TAGLINE, "FFFFFF",
                      bold=True, align=PP_ALIGN.CENTER, font=self.fonts.get("tagline", FONT_CN))

    def add_num_big(self, slide, e):
        """大数字 + 标签（三段式 40/10/50，design-system）。"""
        x, y, w, h = e["x"], e["y"], e["w"], e.get("h", 100)
        num_h = h * 0.40; gap = h * 0.10; lbl_h = h * 0.50
        nb = self.add_box(slide, x, y, w, num_h)
        self.put_text(nb.text_frame, str(e.get("num", "")),
                      float(e.get("fs", FS_STAT)), self.col(e.get("color", "$accent")),
                      bold=True, align=PP_ALIGN.CENTER, font=self.font_for("stat"))
        lb = self.add_box(slide, x, y + num_h + gap, w, lbl_h)
        self.put_text(lb.text_frame, str(e.get("label", "")), FS_MUTED,
                      self.col("$muted", "888888"), align=PP_ALIGN.CENTER, font=self.font_for("body"))

        # ---------- 新增布局组件（html-ppt 31 布局 + MiniMax 页面生成器 + GordenPPTSkill 版式） ----------
    def add_section_divider(self, slide, e):
        """分隔页：大数字 + 标题 + 可选说明（MiniMax Bold Center / Left Accent / Gorden 章节扉页）。"""
        x = float(e.get("x", 60)); y = float(e.get("y", 120))
        w = float(e.get("w", 840)); h = float(e.get("h", 260))
        variant = e.get("variant", "center")
        num = str(e.get("num", "")); title = str(e.get("title", ""))
        intro = str(e.get("intro", ""))
        accent = rgb_of(self.col(e.get("color", "$primary")))
        if variant == "leftaccent":
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, self.emu(x), self.emu(y),
                                         self.emu(8), self.emu(h))
            bar.fill.solid(); bar.fill.fore_color.rgb = accent; bar.line.fill.background()
            if num:
                nb = self.add_box(slide, x + 40, y + 6, w - 60, 58)
                self.put_text(nb.text_frame, num, 58, accent, bold=True, font=FONT_EN)
            tb = self.add_box(slide, x + 40, y + 70, w - 60, 54)
            self.put_text(tb.text_frame, title, FS_TITLE, self.col("$text"), bold=True, font=FONT_CN)
            if intro:
                ib = self.add_box(slide, x + 40, y + 128, w - 60, 40)
                self.put_text(ib.text_frame, intro, FS_BODY, self.col("$muted"), font=FONT_CN)
        else:
            if num:
                nb = self.add_box(slide, x, y, w, 92)
                self.put_text(nb.text_frame, num, 86, accent,
                              bold=True, align=PP_ALIGN.CENTER, font=FONT_EN)
            tb = self.add_box(slide, x, y + 96, w, 54)
            self.put_text(tb.text_frame, title, FS_TITLE, self.col("$text"),
                          bold=True, align=PP_ALIGN.CENTER, font=FONT_CN)
            if intro:
                ib = self.add_box(slide, x, y + 154, w, 32)
                self.put_text(ib.text_frame, intro, FS_BODY, self.col("$muted"),
                              align=PP_ALIGN.CENTER, font=FONT_CN)

    def add_comparison_2col(self, slide, e):
        """左右对比：两卡片 + 顶部色条（html-ppt comparison + MiniMax Comparison）。"""
        x = float(e.get("x", 48)); y = float(e.get("y", 90))
        w = float(e.get("w", 864)); h = float(e.get("h", 320))
        left = e.get("left", {}) or {}; right = e.get("right", {}) or {}
        gap = float(e.get("gap", 24))
        col_w = (w - gap) / 2.0
        lc = rgb_of(self.col(left.get("color", e.get("left_color", "$accent"))))
        rc = rgb_of(self.col(right.get("color", e.get("right_color", "$green"))))
        for (side, cx, ccol) in ((left, x, lc), (right, x + col_w + gap, rc)):
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                          self.emu(cx), self.emu(y),
                                          self.emu(col_w), self.emu(h))
            card.fill.solid(); card.fill.fore_color.rgb = rgb_of(self.col("$paper"))
            card.line.color.rgb = rgb_of(self.col("$line")); card.line.width = Pt(1)
            try:
                card.adjustments[0].value = 0.06
            except Exception:
                pass
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, self.emu(cx), self.emu(y + 4),
                                         self.emu(col_w), self.emu(4))
            bar.fill.solid(); bar.fill.fore_color.rgb = ccol; bar.line.fill.background()
            tt = self.add_box(slide, cx + 20, y + 24, col_w - 40, 32)
            self.put_text(tt.text_frame, str(side.get("title", "")), FS_H3,
                          ccol, bold=True, font=FONT_CN)
            py_ = y + 68
            for pt in (side.get("points", []) or [])[:6]:
                bt = self.add_box(slide, cx + 20, py_, col_w - 40, 26)
                self.put_text(bt.text_frame, "•  " + str(pt), FS_BODY,
                              self.col("$text"), font=FONT_CN)
                py_ += 34

    def add_timeline_h(self, slide, e):
        """横向时间线：轴线 + 节点圆 + 时间/标题/说明（html-ppt timeline + Gorden 时间线）。"""
        x = float(e.get("x", 48)); y = float(e.get("y", 120))
        w = float(e.get("w", 864)); h = float(e.get("h", 260))
        items = e.get("items", []) or []
        n = max(len(items), 1)
        seg = w / float(n)
        acc = rgb_of(self.col(e.get("color", "$accent")))
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, self.emu(x + 20), self.emu(y + 24),
                                      self.emu(w - 40), self.emu(3))
        line.fill.solid(); line.fill.fore_color.rgb = rgb_of(self.col("$line"))
        line.line.fill.background()
        for i, it in enumerate(items[:6]):
            cx = x + seg * i + seg / 2.0 - 20
            dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, self.emu(cx), self.emu(y + 14),
                                         self.emu(22), self.emu(22))
            dot.fill.solid(); dot.fill.fore_color.rgb = acc; dot.line.fill.background()
            tbox = self.add_box(slide, cx - 40, y + 42, 100, 24)
            self.put_text(tbox.text_frame, str(it.get("time", "T%d" % (i + 1))),
                          FS_H4, acc, bold=True, align=PP_ALIGN.CENTER, font=FONT_CN)
            tt = self.add_box(slide, cx - 52, y + 70, 124, 24)
            self.put_text(tt.text_frame, str(it.get("title", "")), FS_BODY,
                          self.col("$text"), bold=True, align=PP_ALIGN.CENTER, font=FONT_CN)
            db = self.add_box(slide, cx - 52, y + 98, 124, 60)
            self.put_text(db.text_frame, str(it.get("desc", "")), FS_MUTED,
                          self.col("$muted"), align=PP_ALIGN.CENTER, font=FONT_CN, line_spacing=1.2)

    def add_process_steps(self, slide, e):
        """流程步骤卡：编号圆 + 卡片 + 箭头（html-ppt process-steps + MiniMax Timeline/Process）。"""
        x = float(e.get("x", 48)); y = float(e.get("y", 120))
        w = float(e.get("w", 864)); h = float(e.get("h", 240))
        steps = e.get("steps", []) or []
        n = max(len(steps), 1)
        gap = 20.0; arrows = 26.0
        seg = (w - (n - 1) * (gap + arrows)) / float(n)
        for i, st in enumerate(steps[:6]):
            cx = x + i * (seg + gap + arrows)
            acc = rgb_of(self.col(st.get("color", e.get("color", "$accent"))))
            c = slide.shapes.add_shape(MSO_SHAPE.OVAL, self.emu(cx + seg / 2 - 16), self.emu(y + 10),
                                       self.emu(32), self.emu(32))
            c.fill.solid(); c.fill.fore_color.rgb = acc; c.line.fill.background()
            nb = self.add_box(slide, cx + seg / 2 - 16, y + 13, 32, 26)
            self.put_text(nb.text_frame, "%02d" % (i + 1), 13, "FFFFFF", bold=True,
                          align=PP_ALIGN.CENTER, font=FONT_EN)
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                          self.emu(cx), self.emu(y + 52),
                                          self.emu(seg), self.emu(h - 60))
            card.fill.solid(); card.fill.fore_color.rgb = rgb_of(self.col("$paper"))
            card.line.color.rgb = rgb_of(self.col("$line")); card.line.width = Pt(1)
            try:
                card.adjustments[0].value = 0.08
            except Exception:
                pass
            tt = self.add_box(slide, cx + 14, y + 62, seg - 28, 26)
            self.put_text(tt.text_frame, str(st.get("title", "")), FS_H4,
                          self.col("$text"), bold=True, align=PP_ALIGN.CENTER, font=FONT_CN)
            db = self.add_box(slide, cx + 14, y + 94, seg - 28, h - 104)
            self.put_text(db.text_frame, str(st.get("desc", "")), FS_MUTED,
                          self.col("$muted"), align=PP_ALIGN.CENTER, font=FONT_CN, line_spacing=1.25)
            if i < n - 1 and e.get("arrow", True):
                ar_x = cx + seg + gap / 2.0 - 10
                tri = slide.shapes.add_shape(MSO_SHAPE.ISOSCELES_TRIANGLE,
                                             self.emu(ar_x), self.emu(y + 34),
                                             self.emu(22), self.emu(18))
                tri.fill.solid(); tri.fill.fore_color.rgb = rgb_of(self.col("$line"))
                tri.line.fill.background()
                try:
                    tri.rotation = 90.0
                except Exception:
                    pass

    def add_kpi_row(self, slide, e):
        """KPI 横排：数字 + 标签 + 增量（html-ppt kpi-grid with deltas）。"""
        x = float(e.get("x", 48)); y = float(e.get("y", 100))
        w = float(e.get("w", 864)); h = float(e.get("h", 150))
        items = e.get("items", []) or []
        n = max(len(items), 1); gap = float(e.get("gap", 16))
        cw = (w - (n - 1) * gap) / float(n)
        for i, it in enumerate(items[:4]):
            cx = x + i * (cw + gap)
            acc = rgb_of(self.col(it.get("color", e.get("color", "$accent"))))
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                          self.emu(cx), self.emu(y),
                                          self.emu(cw), self.emu(h))
            card.fill.solid(); card.fill.fore_color.rgb = rgb_of(self.col("$paper"))
            card.line.color.rgb = rgb_of(self.col("$line")); card.line.width = Pt(1)
            try:
                card.adjustments[0].value = 0.08
            except Exception:
                pass
            nb = self.add_box(slide, cx + 10, y + 12, cw - 20, 48)
            self.put_text(nb.text_frame, str(it.get("num", "")),
                          float(it.get("fs", 34)), acc, bold=True,
                          align=PP_ALIGN.CENTER, font=FONT_STAT)
            lab = self.add_box(slide, cx + 10, y + 66, cw - 20, 22)
            self.put_text(lab.text_frame, str(it.get("label", "")), FS_MUTED,
                          self.col("$muted"), align=PP_ALIGN.CENTER, font=FONT_CN)
            dl = it.get("delta", "")
            if dl:
                up = "+" in str(dl) or "↑" in str(dl) or "增长" in str(dl)
                dbox = self.add_box(slide, cx + 10, y + 94, cw - 20, 24)
                self.put_text(dbox.text_frame, str(dl), 13,
                              "1E7B4F" if up else "B0413E", bold=True,
                              align=PP_ALIGN.CENTER, font=FONT_CN)
    def add_pros_cons(self, slide, e):
        """利弊两卡（html-ppt pros-cons）。pros 绿 / cons 蓝。"""
        x = float(e.get("x", 48)); y = float(e.get("y", 90))
        w = float(e.get("w", 864)); h = float(e.get("h", 320))
        pros = e.get("pros", []) or []; cons = e.get("cons", []) or []
        gap = float(e.get("gap", 24)); col_w = (w - gap) / 2.0
        p_c = rgb_of(self.col(e.get("pro_color", "$green")))
        c_c = rgb_of(self.col(e.get("con_color", "$accent")))
        pro_title = str(e.get("pro_title", "优点"))
        con_title = str(e.get("con_title", "不足"))
        for (htxt, items, cx, ccol) in ((pro_title, pros, x, p_c),
                                        (con_title, cons, x + col_w + gap, c_c)):
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                          self.emu(cx), self.emu(y),
                                          self.emu(col_w), self.emu(h))
            card.fill.solid(); card.fill.fore_color.rgb = rgb_of(self.col("$paper"))
            card.line.color.rgb = rgb_of(self.col("$line")); card.line.width = Pt(1)
            try:
                card.adjustments[0].value = 0.06
            except Exception:
                pass
            tag = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                         self.emu(cx + 20), self.emu(y + 18),
                                         self.emu(64), self.emu(26))
            tag.fill.solid(); tag.fill.fore_color.rgb = ccol; tag.line.fill.background()
            try:
                tag.adjustments[0].value = 0.5
            except Exception:
                pass
            tb = self.add_box(slide, cx + 20, y + 20, 64, 22)
            self.put_text(tb.text_frame, htxt, 14, "FFFFFF", bold=True,
                          align=PP_ALIGN.CENTER, font=FONT_CN)
            py_ = y + 62
            for it in items[:6]:
                bt = self.add_box(slide, cx + 20, py_, col_w - 40, 26)
                mark = "✓  " if htxt == pro_title else "✗  "
                self.put_text(bt.text_frame, mark + str(it), FS_BODY,
                              self.col("$text"), font=FONT_CN)
                py_ += 34

    def add_big_quote(self, slide, e):
        """大引语：左侧品牌粗线 + 引文 + 作者（html-ppt big-quote）。"""
        x = float(e.get("x", 80)); y = float(e.get("y", 150))
        w = float(e.get("w", 800)); h = float(e.get("h", 240))
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, self.emu(x - 30), self.emu(y),
                                     self.emu(8), self.emu(h))
        bar.fill.solid(); bar.fill.fore_color.rgb = rgb_of(self.col(e.get("color", "$accent")))
        bar.line.fill.background()
        qb = self.add_box(slide, x, y, w, h - 60)
        self.put_text(qb.text_frame, "“" + str(e.get("quote", "")) + "”", 26,
                      self.col("$text"), bold=False, font=FONT_CN, line_spacing=1.4)
        ab = self.add_box(slide, x, y + h - 56, w, 30)
        self.put_text(ab.text_frame, "—— " + str(e.get("author", "")), FS_BODY,
                      self.col("$muted"), align=PP_ALIGN.RIGHT, font=FONT_CN)

    def add_roadmap_4col(self, slide, e):
        """NOW / NEXT / LATER / VISION 四列（html-ppt roadmap）。"""
        x = float(e.get("x", 48)); y = float(e.get("y", 110))
        w = float(e.get("w", 864)); h = float(e.get("h", 300))
        cols = e.get("columns", []) or []
        n = max(len(cols), 1); gap = float(e.get("gap", 18))
        cw = (w - (n - 1) * gap) / float(n)
        default_h = {"NOW": "$green", "NEXT": "$accent", "LATER": "$primary", "VISION": "$muted"}
        for i, col in enumerate(cols[:4]):
            cx = x + i * (cw + gap)
            ph = str(col.get("phase", ""))
            cc = rgb_of(self.col(col.get("color", default_h.get(ph, "$accent"))))
            card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                          self.emu(cx), self.emu(y),
                                          self.emu(cw), self.emu(h))
            card.fill.solid(); card.fill.fore_color.rgb = rgb_of(self.col("$paper"))
            card.line.color.rgb = rgb_of(self.col("$line")); card.line.width = Pt(1)
            try:
                card.adjustments[0].value = 0.06
            except Exception:
                pass
            head = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, self.emu(cx), self.emu(y + 4),
                                          self.emu(cw), self.emu(34))
            head.fill.solid(); head.fill.fore_color.rgb = cc; head.line.fill.background()
            hb = self.add_box(slide, cx, y + 10, cw, 24)
            self.put_text(hb.text_frame, ph, 14, "FFFFFF", bold=True,
                          align=PP_ALIGN.CENTER, font=FONT_EN)
            py_ = y + 52
            for it in (col.get("items", []) or [])[:5]:
                ib = self.add_box(slide, cx + 14, py_, cw - 28, 24)
                self.put_text(ib.text_frame, "•  " + str(it), 13,
                              self.col("$text"), font=FONT_CN)
                py_ += 30

    def add_checklist(self, slide, e):
        """清单：完成 ✓ 品牌色 / 未做 □ 灰（html-ppt todo-checklist）。"""
        x = float(e.get("x", 80)); y = float(e.get("y", 110))
        w = float(e.get("w", 800)); h = float(e.get("h", 300))
        items = e.get("items", []) or []
        row_h = 40.0
        for i, it in enumerate(items[:6]):
            ry = y + i * row_h
            done = bool(it.get("done", False))
            mark = "✓" if done else "□"
            mc = "1E7B4F" if done else self.col("$muted", "888888")
            mb = self.add_box(slide, x, ry, 34, 30)
            self.put_text(mb.text_frame, mark, 18, mc, bold=True, font=FONT_CN)
            tb = self.add_box(slide, x + 44, ry, w - 60, 30)
            self.put_text(tb.text_frame, str(it.get("text", "")), FS_BODY,
                          self.col("$text"), font=FONT_CN)

    def add_cover_asym(self, slide, e):
        """非对称封面：左文右品牌色块（MiniMax Asymmetric Left-Right）。"""
        x = float(e.get("x", 0)); y = float(e.get("y", 0))
        w = float(e.get("w", 960)); h = float(e.get("h", 540))
        block_w = float(e.get("block_w", 300))
        side = e.get("block_side", "right")
        if side == "right":
            bx = x + w - block_w
            tx = x + 60; tw = w - block_w - 100
        else:
            bx = x
            tx = x + block_w + 50; tw = w - block_w - 110
        block = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, self.emu(bx), self.emu(y),
                                       self.emu(block_w), self.emu(h))
        block.fill.solid(); block.fill.fore_color.rgb = rgb_of(self.col(e.get("block_color", "$primary")))
        block.line.fill.background()
        acc = rgb_of(self.col(e.get("kicker_color", "$accent")))
        if e.get("kicker"):
            kb = self.add_box(slide, tx, y + 120, tw, 24)
            self.put_text(kb.text_frame, str(e.get("kicker", "")), FS_KICKER,
                          acc, bold=True, font=FONT_CN)
        tb = self.add_box(slide, tx, y + 152, tw, 120)
        self.put_text(tb.text_frame, str(e.get("title", "")), FS_TITLE + 8,
                      self.col("$text"), bold=True, font=FONT_CN)
        if e.get("subtitle"):
            sb = self.add_box(slide, tx, y + 288, tw, 34)
            self.put_text(sb.text_frame, str(e.get("subtitle", "")), FS_BODY,
                          self.col("$muted"), font=FONT_CN)
        if e.get("meta"):
            mb_ = self.add_box(slide, tx, y + 440, tw, 24)
            self.put_text(mb_.text_frame, str(e.get("meta", "")), FS_MUTED,
                          self.col("$muted"), font=FONT_CN)

    # ---------- 检查 ----------


    # ---------- 新增：academic-pptx-skill + deckary 吸收组件 ----------
    def add_figure_text(self, slide, e):
        """结果页：图左文右 + 底部来源（academic-pptx-skill Results 模式）。
        左区 = 图表/图占位（圆角纸色底框 + 图题），右区 = 核心发现（引导词 + 要点），底 = 来源。"""
        x = float(e.get("x", 48)); y = float(e.get("y", 110))
        w = float(e.get("w", 864)); h = float(e.get("h", 340))
        ratio = float(e.get("figure_ratio", 0.55))
        gap = 24.0
        fw = w * ratio
        tw = w - fw - gap
        # 左：图区（圆角纸色底框）
        fig = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                     self.emu(x), self.emu(y), self.emu(fw), self.emu(h - 34))
        fig.fill.solid(); fig.fill.fore_color.rgb = rgb_of(self.col(e.get("figure_bg", "$paper")))
        fig.line.color.rgb = rgb_of(self.col("$line", "CCCCCC")); fig.line.width = Pt(1)
        try:
            fig.adjustments[0].value = 0.03
        except Exception:
            pass
        fig_title = e.get("figure_title", "")
        if fig_title:
            ftb = self.add_box(slide, x + 12, y + 10, fw - 24, 22)
            self.put_text(ftb.text_frame, fig_title, FS_MUTED, self.col("$muted"), font=FONT_CN)
        placeholder = e.get("figure_placeholder", "")
        if placeholder:
            phb = self.add_box(slide, x + 12, y + h / 2 - 14, fw - 24, 28)
            self.put_text(phb.text_frame, placeholder, 13, self.col("$muted"),
                          align=PP_ALIGN.CENTER, font=FONT_CN)
        # 右：核心发现
        tx = x + fw + gap
        if e.get("heading"):
            hb = self.add_box(slide, tx, y, tw, 30)
            self.put_text(hb.text_frame, str(e.get("heading", "")), FS_H4,
                          self.col(e.get("heading_color", "$accent")), bold=True, font=FONT_CN)
        points = e.get("points", []) or []
        pyy = y + 38
        lead_color = self.col(e.get("point_color", "$primary"), "1F2A24")
        for pt in points[:5]:
            if isinstance(pt, dict):
                lead = str(pt.get("lead", ""))
                txt = str(pt.get("text", ""))
            else:
                lead = ""; txt = str(pt)
            if lead:
                pst = lead + "："
                pb = self.add_box(slide, tx, pyy, tw, 24)
                tb = pb.text_frame
                tb.word_wrap = True
                p = tb.paragraphs[0]
                run1 = p.add_run(); run1.text = pst
                run1.font.size = Pt(14); run1.font.bold = True
                run1.font.name = FONT_CN; run1.font.color.rgb = rgb_of(lead_color)
                run2 = p.add_run(); run2.text = txt
                run2.font.size = Pt(14); run2.font.bold = False
                run2.font.name = FONT_CN; run2.font.color.rgb = rgb_of(self.col("$text"))
            else:
                pb = self.add_box(slide, tx, pyy, tw, 24)
                self.put_text(pb.text_frame, "•  " + txt, 14, self.col("$text"), font=FONT_CN)
            pyy += 34
        # 底：来源
        source = e.get("source", "")
        if source:
            sb = self.add_box(slide, x, y + h - 28, w, 22)
            self.put_text(sb.text_frame, "来源：" + str(source), 12,
                          self.col("$muted"), font=FONT_CN)

    def add_breadcrumb(self, slide, e):
        """顶部面包屑导航条（academic-pptx-skill Breadcrumb Bar）。
        全宽浅色条 + 章节均分；当前章节加粗 + 品牌色 + 底部下划线。页面其他内容 y 需 >= 条高+8。"""
        y = float(e.get("y", 0)); hh = float(e.get("height", 30))
        sections = e.get("sections", []) or []
        active = int(e.get("active", 0))
        if not sections:
            return
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, self.emu(0), self.emu(y),
                                     self.emu(self.cw), self.emu(hh))
        bar.fill.solid(); bar.fill.fore_color.rgb = rgb_of(self.col(e.get("bar_bg", "$paper")))
        bar.line.fill.background()
        n = len(sections)
        sw = self.cw / float(n)
        active_color = self.col(e.get("color", "$accent"), "2E75B6")
        inactive_color = self.col("$muted", "777777")
        for i, sec in enumerate(sections):
            is_active = (i == active)
            sb = self.add_box(slide, i * sw, y + 2, sw, hh - 4)
            self.put_text(sb.text_frame, str(sec), 11,
                          active_color if is_active else inactive_color,
                          bold=is_active, align=PP_ALIGN.CENTER, font=FONT_CN)
            if is_active:
                under = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                               self.emu(i * sw + sw * 0.12), self.emu(y + hh - 4),
                                               self.emu(sw * 0.76), self.emu(3))
                under.fill.solid(); under.fill.fore_color.rgb = rgb_of(active_color)
                under.line.fill.background()

    def add_references(self, slide, e):
        """参考文献页：小字条目列表（academic-pptx-skill References）。
        items 为字符串列表；12pt、条目间 6pt 空行、自动换行。"""
        x = float(e.get("x", 48)); y = float(e.get("y", 100))
        w = float(e.get("w", 864)); h = float(e.get("h", 380))
        items = e.get("items", []) or []
        pyy = y
        row_h = 26.0
        for i, ref in enumerate(items[:12]):
            rb = self.add_box(slide, x + 12, pyy, w - 24, row_h + 10)
            self.put_text(rb.text_frame, "[%d]  " % (i + 1) + str(ref), FS_TABLE_CELL,
                          self.col("$text"), font=FONT_EN, line_spacing=1.15)
            pyy += row_h + 12
            if pyy > y + h - row_h:
                break

    def warn(self, msg):
        self.warnings.append(msg)
        log("警告: " + msg)

    def check_bounds(self, e):
        w = float(e.get("w", 0)); h = float(e.get("h", 0))
        x = float(e.get("x", 0)); y = float(e.get("y", 0))
        if x + w > self.cw + 1 or y + h > self.ch + 1:
            self.warn("元素越界：x=%g y=%g w=%g h=%g（画布 %gx%g）" % (x, y, w, h, self.cw, self.ch))

    def check_overlap(self, els):
        """实体碰撞检测：仅对两个非 text 元素报重叠。
        text 文字浮在形状/卡片之上是正常排版（胶囊标签、卡片内文字），不报。
        两个实体（shape/image/table/chart/复合组件）相交才算真遮挡。"""
        boxes = []
        for i, e in enumerate(els):
            if all(k in e for k in ("x", "y", "w", "h")):
                boxes.append((i, e.get("type", "text"), float(e["x"]), float(e["y"]),
                              float(e["x"]) + float(e["w"]), float(e["y"]) + float(e["h"])))
        for a in range(len(boxes)):
            for b in range(a + 1, len(boxes)):
                ia, ta, ax1, ay1, ax2, ay2 = boxes[a]
                ib, tb, bx1, by1, bx2, by2 = boxes[b]
                if ta == "text" or tb == "text":
                    continue  # 文字浮层跳过
                ox = min(ax2, bx2) - max(ax1, bx1)
                oy = min(ay2, by2) - max(ay1, by1)
                if ox > 8 and oy > 8:  # 容差 8pt，忽略细装饰线
                    self.warn("元素重叠：#%d(%s) 与 #%d(%s) 相交 %g×%gpt"
                              % (ia, ta, ib, tb, ox, oy))

    def build_background(self, slide, bg):
        if not bg:
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = rgb_of(self.col("$bg", "FFFFFF"))
            return
        if bg.get("type") == "gradient":
            stops = bg.get("stops", [])
            c0 = self.col(stops[0].get("color") if stops else "$bg", "FFFFFF")
            c1 = self.col(stops[1].get("color") if len(stops) > 1 else "$accent", "3B6CFF")
            angle = int(bg.get("angle", 115))
            xml = (
                '<p:bg %s %s><p:bgPr><a:gradFill flip="none" rotWithShape="1">'
                '<a:gsLst><a:gs pos="0"><a:srgbClr val="%s"/></a:gs>'
                '<a:gs pos="100000"><a:srgbClr val="%s"/></a:gs></a:gsLst>'
                '<a:lin ang="%d"/></a:gradFill><a:effectLst/></p:bgPr></p:bg>'
            ) % (nsdecls("p"), nsdecls("a"), c0, c1, angle * 60000)
            # 修正：p:bg 必须位于 p:cSld 内、p:spTree 之前（OOXML schema 顺序 bg?, spTree）
            cSld = slide._element.find(qn("p:cSld"))
            if cSld is not None:
                cSld.insert(0, parse_xml(xml))
            else:
                slide._element.insert(0, parse_xml(xml))
        else:
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = rgb_of(self.col(bg.get("color", "$bg"), "FFFFFF"))

    def build(self, no_fade=False, strict=False):
        shape_map = {"line": "rect", "rect": "rect", "circle": "oval", "shape": None}
        for slide_data in self.data.get("slides", []):
            slide = self.prs.slides.add_slide(self.prs.slide_layouts[6])
            self.build_background(slide, slide_data.get("background"))
            title = slide_data.get("title", "")
            if title:
                log("页: " + str(slide_data.get("id", "?")) + " " + title)
            elements = slide_data.get("elements", [])
            self.check_overlap(elements)
            for e in elements:
                self.check_bounds(e)
                t = e.get("type", "text")
                if t in shape_map:
                    name = shape_map[t]
                    if name is None:
                        name = e.get("shapeName", "rect")
                    self.add_shape(slide, name, e)
                elif t == "text":
                    self.add_text(slide, e)
                elif t == "image":
                    self.add_image(slide, e)
                elif t == "chart":
                    self.add_chart(slide, e)
                elif t == "table":
                    self.add_table(slide, e)
                elif t == "cards_2x3":
                    self.add_cards_2x3(slide, e)
                elif t == "cards_1x4_info":
                    self.add_cards_1x4_info(slide, e)
                elif t == "card_list_wide":
                    self.add_card_list_wide(slide, e)
                elif t == "tagline_bar":
                    self.add_tagline_bar(slide, e)
                elif t == "num_big":
                    self.add_num_big(slide, e)
                elif t == "section_divider":
                    self.add_section_divider(slide, e)
                elif t == "comparison_2col":
                    self.add_comparison_2col(slide, e)
                elif t == "timeline_h":
                    self.add_timeline_h(slide, e)
                elif t == "process_steps":
                    self.add_process_steps(slide, e)
                elif t == "kpi_row":
                    self.add_kpi_row(slide, e)
                elif t == "pros_cons":
                    self.add_pros_cons(slide, e)
                elif t == "big_quote":
                    self.add_big_quote(slide, e)
                elif t == "roadmap_4col":
                    self.add_roadmap_4col(slide, e)
                elif t == "checklist":
                    self.add_checklist(slide, e)
                elif t == "cover_asym":
                    self.add_cover_asym(slide, e)
                elif t == "figure_text":
                    self.add_figure_text(slide, e)
                elif t == "breadcrumb":
                    self.add_breadcrumb(slide, e)
                elif t == "references":
                    self.add_references(slide, e)
                else:
                    self.warn("未知元素类型: " + str(t))
            if not no_fade:
                set_fade(slide)
        return self.warnings


def main(argv=None):
    parser = argparse.ArgumentParser(description="ppt-studio JSON → PPTX 引擎 v3")
    parser.add_argument("input", type=Path, help="deck.json 数据文件")
    parser.add_argument("-o", "--output", type=Path, help="输出 .pptx 路径")
    parser.add_argument("--no-fade", action="store_true", help="关闭淡入淡出过渡")
    parser.add_argument("--strict", action="store_true", help="存在警告时非零退出")
    args = parser.parse_args(argv)

    if not args.input.is_file():
        print("error: 找不到输入文件 " + str(args.input), file=sys.stderr)
        return 2
    data = json.loads(args.input.read_text(encoding="utf-8"))
    out = args.output or (args.input.parent / "deck.pptx")
    b = Builder(data, args.input.parent)
    warnings = b.build(no_fade=args.no_fade)
    b.prs.save(str(out))
    log("导出 %d 页 → %s" % (len(data.get("slides", [])), out))
    if warnings:
        log("%d 条警告（--strict 将失败）" % len(warnings))
        if args.strict:
            return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

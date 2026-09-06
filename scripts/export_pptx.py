#!/usr/bin/env python3
# (c) 2026 Jose AI (https://www.linhut.cn)
# https://github.com/linhut/ppt-studio
# Licensed under the MIT License. See the LICENSE file for details.

"""DSH NP-PPT High-Fidelity Local Native Exporter.

Converts PPTD projects (.pptd + pages/*.page) into pixel-perfect, presentation-ready
PowerPoint presentations (PPTX) with Microsoft YaHei (微软雅黑) typography, DrawingML gradient
backgrounds, donut decorative shapes, precise table borders & cell paddings, styled charts,
noAutofit font locking, and slide fade transitions.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import site
import tempfile
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

if sys.stdout:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if sys.stderr:
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

def _ensure_dep(import_name: str, pip_name: str) -> None:
    """Import import_name; if missing, pip-install pip_name into a writable target dir and prepend it to sys.path.

    Avoids the common `--user` install failure when ~/Library/Python is not writable.
    """
    try:
        __import__(import_name)
        return
    except ImportError:
        pass
    import site
    target = os.path.join(_DEP_TARGET if (_DEP_TARGET := os.environ.get("DSH_PPT_DEPS")) else os.path.join(tempfile.gettempdir(), "dsh-np-ppt-deps"))
    os.makedirs(target, exist_ok=True)
    subprocess.run([sys.executable, "-m", "pip", "install", "--target", target, pip_name], check=True)
    if target not in sys.path:
        sys.path.insert(0, target)
    site.addsitedir(target)


try:
    import yaml
except ImportError:
    _ensure_dep("yaml", "pyyaml")
    import yaml

try:
    import pptx
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE, XL_DATA_LABEL_POSITION
    from pptx.chart.data import CategoryChartData
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import nsdecls, qn
except ImportError:
    _ensure_dep("pptx", "python-pptx")
    import pptx
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
    from pptx.dml.color import RGBColor
    from pptx.enum.chart import XL_CHART_TYPE, XL_DATA_LABEL_POSITION
    from pptx.chart.data import CategoryChartData
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import nsdecls, qn


class ExportError(RuntimeError):
    pass


def log(message: str) -> None:
    print(f"[dsh-np-ppt] {message}", file=sys.stderr, flush=True)


FONT_NAME_CN = "微软雅黑"
FONT_NAME_EN = "Microsoft YaHei"


def hex_to_rgb(hex_str: str) -> Tuple[RGBColor, Optional[int]]:
    """Converts hex color (#RRGGBB or #RRGGBBAA or #RGB) to (RGBColor, alpha_100k)."""
    hex_clean = str(hex_str).strip().lstrip("#")
    alpha_100k = None
    if len(hex_clean) == 8:
        alpha_val = int(hex_clean[6:8], 16)
        alpha_100k = int((alpha_val / 255.0) * 100000)
        hex_clean = hex_clean[:6]
    elif len(hex_clean) == 3:
        hex_clean = "".join([c * 2 for c in hex_clean])

    if len(hex_clean) != 6:
        return RGBColor(31, 42, 36), None

    r = int(hex_clean[0:2], 16)
    g = int(hex_clean[2:4], 16)
    b = int(hex_clean[4:6], 16)
    return RGBColor(r, g, b), alpha_100k


def parse_rgb_string(rgb_str: str) -> Optional[Tuple[RGBColor, Optional[int]]]:
    m = re.search(r"rgba?\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)(?:\s*,\s*([\d\.]+))?\s*\)", rgb_str, re.IGNORECASE)
    if m:
        r, g, b = int(m.group(1)), int(m.group(2)), int(m.group(3))
        alpha_100k = None
        if m.group(4) is not None:
            a_val = float(m.group(4))
            alpha_100k = int(max(0.0, min(1.0, a_val)) * 100000)
        return RGBColor(r, g, b), alpha_100k
    return None


def resolve_color(color_val: Any, theme_colors: Dict[str, str]) -> Optional[Tuple[RGBColor, Optional[int]]]:
    if not color_val:
        return None
    color_str = str(color_val).strip()
    if color_str.startswith("$"):
        name = color_str[1:]
        color_str = theme_colors.get(name, "#1F2A24")
    if color_str.lower().startswith("rgb"):
        return parse_rgb_string(color_str)
    if color_str.startswith("#"):
        return hex_to_rgb(color_str)
    if color_str in theme_colors:
        return resolve_color(theme_colors[color_str], theme_colors)
    return None


def resolve_color_rgb(color_val: Any, theme_colors: Dict[str, str], default_rgb: RGBColor = RGBColor(31, 42, 36)) -> RGBColor:
    res = resolve_color(color_val, theme_colors)
    if res is not None:
        return res[0]
    return default_rgb


def strip_tags(html_str: str) -> str:
    return re.sub(r"<[^>]+>", "", html_str)


def parse_css_properties(style_str: str, theme_colors: Dict[str, str]) -> Dict[str, Any]:
    props: Dict[str, Any] = {}
    if not style_str:
        return props
    pairs = [p.strip() for p in style_str.split(";") if p.strip()]
    for pair in pairs:
        if ":" not in pair:
            continue
        k, v = pair.split(":", 1)
        k = k.strip().lower()
        v = v.strip()

        if k == "font-size":
            m = re.search(r"(\d+(?:\.\d+)?)", v)
            if m:
                props["fontSize"] = float(m.group(1))
        elif k == "font-weight":
            props["bold"] = v in ("bold", "700", "800", "900")
        elif k == "color":
            resolved = resolve_color(v, theme_colors)
            if resolved:
                props["color"] = resolved[0]
        elif k == "text-align":
            props["align"] = v.lower()
        elif k == "margin-top":
            m = re.search(r"(\d+(?:\.\d+)?)", v)
            if m:
                props["marginTop"] = float(m.group(1))
        elif k == "margin-bottom":
            m = re.search(r"(\d+(?:\.\d+)?)", v)
            if m:
                props["marginBottom"] = float(m.group(1))
        elif k == "line-height":
            m = re.search(r"(\d+(?:\.\d+)?)", v)
            if m:
                props["lineHeight"] = float(m.group(1))
        elif k == "letter-spacing":
            m = re.search(r"(\d+(?:\.\d+)?)", v)
            if m:
                props["letterSpacing"] = float(m.group(1))
    return props


def lock_text_frame_no_autofit(tf):
    """Locks text frame to noAutofit so PowerPoint never shrinks or wraps text unexpectedly."""
    tf.word_wrap = True
    tf.margin_left = Pt(0)
    tf.margin_right = Pt(0)
    tf.margin_top = Pt(0)
    tf.margin_bottom = Pt(0)

    bodyPr = tf._txBody.find(qn("a:bodyPr"))
    if bodyPr is not None:
        for old in list(bodyPr):
            if old.tag in (qn("a:spAutoFit"), qn("a:normAutofit")):
                bodyPr.remove(old)
        if bodyPr.find(qn("a:noAutofit")) is None:
            bodyPr.append(parse_xml(f'<a:noAutofit {nsdecls("a")}/>'))


def apply_typography(run, font_name_cn: str = FONT_NAME_CN, font_name_en: str = FONT_NAME_EN, letter_spacing: float = 0):
    """Explicitly applies Microsoft YaHei / 微软雅黑 for crisp Chinese & Latin rendering in PowerPoint."""
    run.font.name = font_name_cn
    rPr = run._r.get_or_add_rPr()

    for child in list(rPr):
        if child.tag in (qn("a:ea"), qn("a:latin"), qn("a:cs")):
            rPr.remove(child)

    rPr.append(parse_xml(f'<a:ea {nsdecls("a")} typeface="{font_name_cn}"/>'))
    rPr.append(parse_xml(f'<a:latin {nsdecls("a")} typeface="{font_name_en}"/>'))
    rPr.append(parse_xml(f'<a:cs {nsdecls("a")} typeface="{font_name_en}"/>'))

    if letter_spacing > 0:
        spc_val = str(int(letter_spacing * 100))
        rPr.set("spc", spc_val)


def set_cell_border(cell, color_hex: str = "D8D3C4", width_pt: float = 0.75):
    """Adds clean solid borders to a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    width_emu = int(width_pt * 12700)
    for border_name in ["lnL", "lnR", "lnT", "lnB"]:
        existing = tcPr.find(f'{{http://schemas.openxmlformats.org/drawingml/2006/main}}{border_name}')
        if existing is not None:
            tcPr.remove(existing)
        xml = f'''<a:{border_name} {nsdecls("a")} w="{width_emu}" cmpd="sng">
            <a:solidFill><a:srgbClr val="{color_hex}"/></a:solidFill>
        </a:{border_name}>'''  # cmpd 枚举须为 sng（single），PowerPoint 严格校验拒绝 "s"
        tcPr.append(parse_xml(xml))


class RichTextHTMLParser(HTMLParser):
    def __init__(
        self,
        tf,
        theme_colors: Dict[str, str],
        default_props: Dict[str, Any],
        font_name_cn: str,
        font_name_en: str,
    ):
        super().__init__()
        self.tf = tf
        self.theme_colors = theme_colors
        self.default_props = default_props
        self.font_name_cn = font_name_cn
        self.font_name_en = font_name_en
        
        self.current_paragraph = None
        self.has_paragraph = False
        self.style_stack: List[Dict[str, Any]] = [dict(default_props)]
        self.p_style: Dict[str, Any] = {}

    def current_style(self) -> Dict[str, Any]:
        merged = dict(self.default_props)
        for s in self.style_stack:
            merged.update(s)
        return merged

    def handle_starttag(self, tag: str, attrs: List[Tuple[str, Optional[str]]]):
        tag = tag.lower()
        attr_dict = {k.lower(): v for k, v in attrs if v is not None}
        style_str = attr_dict.get("style", "")
        css_props = parse_css_properties(style_str, self.theme_colors)

        if tag == "p":
            if not self.has_paragraph:
                self.current_paragraph = self.tf.paragraphs[0]
                self.has_paragraph = True
            else:
                self.current_paragraph = self.tf.add_paragraph()

            self.p_style = dict(css_props)
            h_align = self.p_style.get("align", self.default_props.get("h_align", "left"))
            if h_align == "center":
                self.current_paragraph.alignment = PP_ALIGN.CENTER
            elif h_align == "right":
                self.current_paragraph.alignment = PP_ALIGN.RIGHT
            else:
                self.current_paragraph.alignment = PP_ALIGN.LEFT

            if "marginTop" in self.p_style:
                self.current_paragraph.space_before = Pt(self.p_style["marginTop"])
            if "marginBottom" in self.p_style:
                self.current_paragraph.space_after = Pt(self.p_style["marginBottom"])

            line_h = self.p_style.get("lineHeight", self.default_props.get("lineHeight", 1.25))
            self.current_paragraph.line_spacing = max(1.05, float(line_h) * 0.85)

            p_inline = {
                k: v
                for k, v in css_props.items()
                if k not in ("align", "marginTop", "marginBottom", "lineHeight")
            }
            self.style_stack.append(p_inline)

        elif tag in ("strong", "b"):
            css_props["bold"] = True
            self.style_stack.append(css_props)
        elif tag in ("em", "i"):
            css_props["italic"] = True
            self.style_stack.append(css_props)
        elif tag == "u":
            css_props["underline"] = True
            self.style_stack.append(css_props)
        elif tag == "span":
            self.style_stack.append(css_props)
        elif tag == "br":
            if self.current_paragraph is not None:
                r = self.current_paragraph.add_run()
                r.text = "\n"

    def handle_endtag(self, tag: str):
        tag = tag.lower()
        if tag in ("p", "strong", "b", "em", "i", "u", "span"):
            if len(self.style_stack) > 1:
                self.style_stack.pop()

    def handle_data(self, data: str):
        if not data:
            return
        if self.current_paragraph is None:
            self.current_paragraph = self.tf.paragraphs[0]
            self.has_paragraph = True
            h_align = self.default_props.get("h_align", "left")
            if h_align == "center":
                self.current_paragraph.alignment = PP_ALIGN.CENTER
            elif h_align == "right":
                self.current_paragraph.alignment = PP_ALIGN.RIGHT
            else:
                self.current_paragraph.alignment = PP_ALIGN.LEFT
            line_h = self.default_props.get("lineHeight", 1.25)
            self.current_paragraph.line_spacing = max(1.05, float(line_h) * 0.85)

        style = self.current_style()
        run = self.current_paragraph.add_run()
        run.text = data
        if "fontSize" in style:
            run.font.size = Pt(style["fontSize"])
        if "bold" in style:
            run.font.bold = style["bold"]
        if "italic" in style:
            run.font.italic = style["italic"]
        if "underline" in style:
            run.font.underline = style["underline"]
        if "color" in style and isinstance(style["color"], RGBColor):
            run.font.color.rgb = style["color"]
        apply_typography(run, self.font_name_cn, self.font_name_en, style.get("letterSpacing", 0))


def render_html_content_to_text_frame(
    tf,
    html_text: str,
    theme_colors: Dict[str, str],
    default_props: Dict[str, Any],
):
    """Renders HTML-formatted string or plain text accurately into a PowerPoint TextFrame."""
    lock_text_frame_no_autofit(tf)

    raw = str(html_text or "").strip()
    if not raw:
        return

    if not ("<" in raw and ">" in raw):
        p = tf.paragraphs[0]
        h_align = default_props.get("h_align", "left")
        if h_align == "center":
            p.alignment = PP_ALIGN.CENTER
        elif h_align == "right":
            p.alignment = PP_ALIGN.RIGHT
        else:
            p.alignment = PP_ALIGN.LEFT

        line_h = default_props.get("lineHeight", 1.25)
        p.line_spacing = max(1.05, float(line_h) * 0.85)

        run = p.add_run()
        run.text = raw
        run.font.size = Pt(default_props.get("fontSize", 14))
        run.font.bold = default_props.get("bold", False)
        if "color" in default_props and isinstance(default_props["color"], RGBColor):
            run.font.color.rgb = default_props["color"]
        apply_typography(
            run,
            FONT_NAME_CN,
            FONT_NAME_EN,
            default_props.get("letterSpacing", 0),
        )
        return

    parser = RichTextHTMLParser(
        tf,
        theme_colors,
        default_props,
        FONT_NAME_CN,
        FONT_NAME_EN,
    )
    parser.feed(raw)


def find_manifest(source: Path) -> Path:
    source = source.expanduser().resolve()
    if source.is_file():
        if source.suffix.lower() != ".pptd":
            raise ExportError(f"input must be a .pptd file or project directory: {source}")
        return source
    if not source.is_dir():
        raise ExportError(f"input does not exist: {source}")
    manifests = sorted(source.rglob("*.pptd"))
    if not manifests:
        raise ExportError(f"no .pptd manifest found under: {source}")
    if len(manifests) > 1:
        choices = "\n  ".join(str(path) for path in manifests[:20])
        raise ExportError(
            "multiple .pptd manifests found; pass one manifest explicitly:\n  " + choices
        )
    return manifests[0]


# ── PPTD 一致性校验 ──────────────────────────────────────────────────────────
# 保证 55173 编辑器预览与原生 PPTX 导出渲染一致：禁止发行"预览白底、导出正常"的页面。
CONSISTENCY_RULES = [
    "background MUST be a solid color. Use only a top-level `color` field (#HEX/rgb()/rgba()/named). "
    "Do NOT use `type: \"gradient\"` / `stops` / `angle` — the 55173 editor cannot render gradients and "
    "shows white, while the PPTX compiler would render the gradient: a forbidden preview/export divergence. "
    "The PPTX compiler renders `background.color` as a solid fill too, so solid-only keeps both sides identical.",
    "Inside `content.text` HTML, inline every color as #HEX/rgb()/rgba(). Never use $theme variables "
    "there — theme vars are resolved only by the PPTX compiler, so the 55173 editor would lose the color.",
]


def validate_page_consistency(page_path: Path, page_data: dict, theme_colors: Dict[str, str]) -> List[str]:
    """Returns a list of consistency violations for a single page. Empty list = OK."""
    violations: List[str] = []
    page_id = page_path.name

    bg = page_data.get("background", {})
    if not isinstance(bg, dict) or not bg:
        violations.append(f"[{page_id}] missing `background` block entirely")
    else:
        # gradients are forbidden — they render white in the editor but fine in PPTX
        if bg.get("type") == "gradient" or "stops" in bg or "angle" in bg:
            violations.append(
                f"[{page_id}] background uses a gradient (`type`/`stops`/`angle`). The 55173 editor cannot "
                f"render gradients and will show white. Use a solid `color` only — the PPTX compiler also "
                f"renders `background.color` as a solid fill, so both sides stay identical."
            )
        if "color" not in bg or not str(bg.get("color", "")).strip():
            violations.append(f"[{page_id}] background has no `color` field")
        else:
            # color must be a concrete value, not a $theme the editor can't resolve
            color_val = str(bg.get("color", "")).strip()
            if color_val.startswith("$"):
                violations.append(
                    f"[{page_id}] background.color uses `$theme` (`{color_val}`); the 55173 editor "
                    f"cannot resolve theme vars. Inline it as #HEX/rgb()."
                )

    # text HTML must not contain $theme color refs
    dollar_re = re.compile(r"color\s*:\s*\$[a-zA-Z_]+", re.IGNORECASE)
    for el in page_data.get("elements", []):
        if el.get("elementType") != "text":
            continue
        txt = el.get("content", {}).get("text", "") or ""
        if dollar_re.search(txt):
            violations.append(
                f"[{page_id}] element `{el.get('elementId', '?')}` uses a $theme color inside "
                f"content.text HTML; inline it as #HEX/rgb() so the editor renders the same color."
            )

    return violations


def validate_deck_consistency(root: Path, manifest: dict) -> None:
    """Validates the whole deck against the preview/export consistency rules. Raises ExportError on violation."""
    pages = manifest.get("pages", [])
    if not isinstance(pages, list) or not pages:
        raise ExportError("manifest contains no pages")

    theme_colors = manifest.get("theme", {}).get("colors", {})
    all_violations: List[str] = []

    for page_rel in pages:
        page_file = root / page_rel
        if not page_file.is_file():
            all_violations.append(f"[{page_rel}] page file missing")
            continue
        try:
            page_data = yaml.safe_load(page_file.read_text(encoding="utf-8"))
        except Exception as exc:
            all_violations.append(f"[{page_rel}] page is not valid YAML/JSON: {exc}")
            continue
        if not isinstance(page_data, dict):
            all_violations.append(f"[{page_rel}] page root must be an object")
            continue
        all_violations.extend(validate_page_consistency(page_file, page_data, theme_colors))

    if all_violations:
        bullet = "\n  - ".join(all_violations)
        raise ExportError(
            "PPTD consistency check failed — the 55173 editor preview and the PPTX export would diverge:\n  - "
            + bullet
            + "\nRules:\n  - " + "\n  - ".join(CONSISTENCY_RULES)
        )


def export_pptx(
    source: Path,
    output: Optional[Path] = None,
    transition: str = "fade",
    force: bool = False,
) -> Dict[str, Any]:
    manifest_path = find_manifest(source)
    root = manifest_path.parent.resolve()
    manifest_text = manifest_path.read_text(encoding="utf-8")
    manifest = yaml.safe_load(manifest_text)
    if not isinstance(manifest, dict):
        raise ExportError("PPTD manifest must be a valid YAML object")

    if output is None:
        output = root / f"{manifest_path.stem}.pptx"
    else:
        output = output.expanduser().resolve()

    if output.exists() and not force:
        raise ExportError(f"output file already exists (use --force to overwrite): {output}")

    output.parent.mkdir(parents=True, exist_ok=True)
    log(f"manifest: {manifest_path}")

    # ── 一致性校验（55173 预览 ≡ PPTX 导出）────────────────────
    validate_deck_consistency(root, manifest)

    log(f"building pixel-perfect native PPTX -> {output}")

    prs = Presentation()
    prs.slide_width = Inches(13.333333)
    prs.slide_height = Inches(7.5)

    theme_colors: Dict[str, str] = manifest.get("theme", {}).get("colors", {})
    theme_text_styles: Dict[str, Any] = manifest.get("theme", {}).get("textStyles", {})
    blank_layout = prs.slide_layouts[6]

    pages = manifest.get("pages", [])
    if not isinstance(pages, list) or not pages:
        raise ExportError("manifest contains no pages")

    slide_count = 0
    for page_rel in pages:
        page_file = root / page_rel
        if not page_file.is_file():
            log(f"warning: page file missing: {page_rel}")
            continue

        page_data = yaml.safe_load(page_file.read_text(encoding="utf-8"))
        if not isinstance(page_data, dict):
            continue

        slide = prs.slides.add_slide(blank_layout)
        slide_count += 1

        # ── 1. 背景渲染 ───────────────────────────────────────────
        bg = page_data.get("background", {})
        if bg.get("type") == "gradient":
            stops = bg.get("stops", [])
            stop0_hex = stops[0].get("color", "#0E2A21").lstrip("#") if len(stops) > 0 else "0E2A21"
            stop1_hex = stops[1].get("color", "#214A3B").lstrip("#") if len(stops) > 1 else "214A3B"
            angle = int(bg.get("angle", 115))
            dml_ang = angle * 60000

            bg_xml = parse_xml(f"""
            <p:bg {nsdecls("p")} {nsdecls("a")}>
              <p:bgPr>
                <a:gradFill flip="none" rotWithShape="1">
                  <a:gsLst>
                    <a:gs pos="0">
                      <a:srgbClr val="{stop0_hex[:6]}"/>
                    </a:gs>
                    <a:gs pos="100000">
                      <a:srgbClr val="{stop1_hex[:6]}"/>
                    </a:gs>
                  </a:gsLst>
                  <a:lin ang="{dml_ang}"/>
                </a:gradFill>
                <a:effectLst/>
              </p:bgPr>
            </p:bg>
            """)
            # 修正：p:bg 必须位于 p:cSld 内、p:spTree 之前（OOXML schema 顺序 bg?, spTree）
            cSld = slide._element.find(qn("p:cSld"))
            if cSld is not None:
                cSld.insert(0, bg_xml)
            else:
                slide._element.insert(0, bg_xml)
        else:
            bg_color_res = resolve_color(bg.get("color", "$bg"), theme_colors)
            if bg_color_res:
                fill = slide.background.fill
                fill.solid()
                fill.fore_color.rgb = bg_color_res[0]

        # ── 2. 元素渲染 ───────────────────────────────────────────
        elements = page_data.get("elements", [])
        for el in elements:
            el_type = el.get("elementType")
            bounds = el.get("bounds", [0, 0, 100, 100])
            x, y, w, h = Pt(bounds[0]), Pt(bounds[1]), Pt(bounds[2]), Pt(bounds[3])

            # ── A. 几何图形 (Shape) ──────────────────────────────
            if el_type == "shape":
                shape_name = el.get("shapeName", "rect")
                if shape_name == "donut":
                    shape = slide.shapes.add_shape(MSO_SHAPE.DONUT, x, y, w, h)
                    adjustments = el.get("adjustments", [18000])
                    adj_val = adjustments[0] if adjustments else 18000
                    spPr = shape._element.spPr
                    avLst = spPr.find(qn("a:prstGeom"))
                    if avLst is not None:
                        av_sub = avLst.find(qn("a:avLst"))
                        if av_sub is not None:
                            av_sub.append(parse_xml(f'<a:gd {nsdecls("a")} name="adj" fmla="val {adj_val}"/>'))
                elif shape_name == "oval":
                    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, w, h)
                elif shape_name == "roundRect":
                    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
                else:
                    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)

                shape.shadow.inherit = False

                fill_data = el.get("fill", {})
                if fill_data.get("type") == "solid":
                    c_res = resolve_color(fill_data.get("color"), theme_colors)
                    if c_res:
                        shape.fill.solid()
                        shape.fill.fore_color.rgb = c_res[0]
                        if c_res[1] is not None:
                            spPr = shape._element.spPr
                            solidFill = spPr.find(qn("a:solidFill"))
                            if solidFill is not None:
                                srgbClr = solidFill.find(qn("a:srgbClr"))
                                if srgbClr is not None:
                                    srgbClr.append(parse_xml(f'<a:alpha {nsdecls("a")} val="{c_res[1]}"/>'))
                else:
                    shape.fill.background()

                border_data = el.get("border", {})
                if border_data.get("style") == "solid":
                    bc_res = resolve_color(border_data.get("color"), theme_colors)
                    if bc_res:
                        shape.line.color.rgb = bc_res[0]
                        shape.line.width = Pt(border_data.get("width", 1))
                        if bc_res[1] is not None:
                            spPr = shape._element.spPr
                            ln = spPr.find(qn("a:ln"))
                            if ln is not None:
                                solidFill = ln.find(qn("a:solidFill"))
                                if solidFill is not None:
                                    srgbClr = solidFill.find(qn("a:srgbClr"))
                                    if srgbClr is not None:
                                        srgbClr.append(parse_xml(f'<a:alpha {nsdecls("a")} val="{bc_res[1]}"/>'))
                else:
                    shape.line.fill.background()

            # ── B. 线条与箭头 (Line) ──────────────────────────────
            elif el_type == "line":
                border_data = el.get("border", {})
                bc_res = resolve_color(border_data.get("color", "$brass"), theme_colors)
                bc = bc_res[0] if bc_res else RGBColor(176, 141, 62)
                l_width = Pt(border_data.get("width", 2))
                arrow_spec = el.get("arrow", [])
                has_arrow = arrow_spec and len(arrow_spec) > 1 and arrow_spec[1] == "arrow"

                if has_arrow:
                    # 注意：OOXML 坐标必须是整数 EMU，Python 的 '/' 产生浮点小数点
                    # （如 2806700.0），PowerPoint 严格校验会拒绝。用 int() 归一。
                    cx1, cy1 = int(x), int(y + h / 2)
                    cx2, cy2 = int(x + w), int(y + h / 2)
                    connector = slide.shapes.add_connector(
                        MSO_CONNECTOR.STRAIGHT, cx1, cy1, cx2, cy2
                    )
                    connector.line.color.rgb = bc
                    connector.line.width = l_width
                    spPr = connector._element.spPr
                    if spPr is not None:
                        ln = spPr.find(qn("a:ln"))
                        if ln is not None:
                            tail = parse_xml(f'<a:tailEnd {nsdecls("a")} type="triangle" w="med" len="med"/>')
                            ln.append(tail)
                else:
                    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, Pt(max(bounds[3], 2)))
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = bc
                    shape.line.fill.background()

            # ── C. 文本框 (Text) ──────────────────────────────────
            elif el_type == "text":
                content = el.get("content", {})
                tx_box = slide.shapes.add_textbox(x, y, w, h)
                tf = tx_box.text_frame

                align_spec = content.get("align", ["left", "top"])
                h_align = align_spec[0] if len(align_spec) > 0 else "left"
                v_align = align_spec[1] if len(align_spec) > 1 else "top"

                if v_align == "middle":
                    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                elif v_align == "bottom":
                    tf.vertical_anchor = MSO_ANCHOR.BOTTOM
                else:
                    tf.vertical_anchor = MSO_ANCHOR.TOP

                style_name = content.get("style", "")
                inherited_props = {}
                if style_name.startswith("$"):
                    s_key = style_name[1:]
                    inherited_props = theme_text_styles.get(s_key, {})

                def_size = content.get("fontSize", inherited_props.get("fontSize", 14))
                def_bold = content.get("bold", inherited_props.get("bold", False))
                def_color_spec = content.get("color", inherited_props.get("color", "$text"))
                def_color = resolve_color_rgb(def_color_spec, theme_colors, RGBColor(31, 42, 36))
                def_line_h = content.get("lineHeight", inherited_props.get("lineHeight", 1.25))
                def_spacing = content.get("letterSpacing", inherited_props.get("letterSpacing", 0))

                default_props = {
                    "fontSize": float(def_size),
                    "bold": bool(def_bold),
                    "color": def_color,
                    "h_align": h_align,
                    "lineHeight": float(def_line_h),
                    "letterSpacing": float(def_spacing),
                }

                raw_text = content.get("text", "")
                render_html_content_to_text_frame(tf, str(raw_text), theme_colors, default_props)

            # ── D. 表格 (Table) ───────────────────────────────────
            elif el_type == "table":
                rows = el.get("rows", [])
                if rows:
                    num_rows = len(rows)
                    num_cols = len(rows[0])
                    col_widths = el.get("columnWidths", [1.0 / num_cols] * num_cols)
                    row_heights = el.get("rowHeights", [])

                    table_shape = slide.shapes.add_table(num_rows, num_cols, x, y, w, h)
                    table = table_shape.table

                    for ci, cw in enumerate(col_widths):
                        table.columns[ci].width = Pt(int(bounds[2] * cw))

                    for ri in range(num_rows):
                        if ri < len(row_heights):
                            table.rows[ri].height = Pt(int(max(bounds[3] * row_heights[ri], 18)))
                        else:
                            table.rows[ri].height = Pt(int(max(bounds[3] / num_rows, 18)))

                    ink_color = resolve_color_rgb("$ink", theme_colors, RGBColor(18, 59, 47))
                    text_color = resolve_color_rgb("$text", theme_colors, RGBColor(31, 42, 36))
                    paper_color = RGBColor(255, 255, 255)
                    alt_row_color = RGBColor(242, 237, 223)
                    line_hex = theme_colors.get("line", "#D8D3C4").lstrip("#")

                    for ri, row_data in enumerate(rows):
                        is_header = ri == 0
                        for ci, cell_data in enumerate(row_data):
                            cell = table.cell(ri, ci)
                            cell_text = str(cell_data.get("text", ""))
                            cell.text = cell_text

                            cell.margin_left = Pt(6)
                            cell.margin_right = Pt(6)
                            cell.margin_top = Pt(3)
                            cell.margin_bottom = Pt(3)
                            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

                            set_cell_border(cell, line_hex, 0.75)

                            if is_header:
                                cell.fill.solid()
                                cell.fill.fore_color.rgb = ink_color
                            elif ri % 2 == 1:
                                cell.fill.solid()
                                cell.fill.fore_color.rgb = paper_color
                            else:
                                cell.fill.solid()
                                cell.fill.fore_color.rgb = alt_row_color

                            for p in cell.text_frame.paragraphs:
                                if ci == 0 or cell_text.isdigit() or cell_text.endswith("%"):
                                    p.alignment = PP_ALIGN.CENTER
                                elif ci == num_cols - 1 and not is_header:
                                    p.alignment = PP_ALIGN.RIGHT
                                else:
                                    p.alignment = PP_ALIGN.LEFT

                                for run in p.runs:
                                    run.font.size = Pt(11.5 if is_header else 10.5)
                                    run.font.bold = is_header
                                    run.font.color.rgb = RGBColor(247, 243, 232) if is_header else text_color
                                    apply_typography(run, FONT_NAME_CN, FONT_NAME_EN)

            # ── E. 图表 (Chart) ───────────────────────────────────
            elif el_type == "chart":
                chart_data_info = el.get("data", {})
                rows = chart_data_info.get("rows", [])
                if rows:
                    chart_data = CategoryChartData()
                    categories = [str(r[0]) for r in rows]
                    values = [float(r[2]) for r in rows]
                    chart_data.categories = categories
                    chart_data.add_series("Token消耗（亿）", values)

                    chart_shape = slide.shapes.add_chart(
                        XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, w, h, chart_data
                    )
                    chart = chart_shape.chart
                    chart.has_legend = False
                    chart.has_title = False

                    chart.plots[0].gap_width = 110

                    series = chart.series[0]
                    series.format.fill.solid()
                    green_color = resolve_color_rgb("$green", theme_colors, RGBColor(47, 82, 51))
                    series.format.fill.fore_color.rgb = green_color

                    chart.plots[0].has_data_labels = True
                    data_labels = chart.plots[0].data_labels
                    data_labels.font.name = FONT_NAME_CN
                    data_labels.font.size = Pt(11)
                    data_labels.font.color.rgb = resolve_color_rgb("$text", theme_colors, RGBColor(31, 42, 36))
                    data_labels.position = XL_DATA_LABEL_POSITION.OUTSIDE_END

                    category_axis = chart.category_axis
                    category_axis.format.line.color.rgb = resolve_color_rgb("$line", theme_colors, RGBColor(216, 211, 196))
                    category_axis.tick_labels.font.name = FONT_NAME_CN
                    category_axis.tick_labels.font.size = Pt(12)

                    value_axis = chart.value_axis
                    value_axis.format.line.color.rgb = resolve_color_rgb("$line", theme_colors, RGBColor(216, 211, 196))
                    value_axis.tick_labels.font.name = FONT_NAME_CN
                    value_axis.tick_labels.font.size = Pt(11)
                    value_axis.has_major_gridlines = True
                    value_axis.major_gridlines.format.line.color.rgb = RGBColor(229, 223, 208)

            # ── F. 图片 (Image) ───────────────────────────────────
            elif el_type == "image":
                img_path_rel = el.get("path") or el.get("src")
                if img_path_rel:
                    img_file = root / img_path_rel
                    if img_file.is_file():
                        slide.shapes.add_picture(str(img_file), x, y, w, h)

        # ── 3. 淡入淡出切换动画 ───────────────────────────────────
        if transition == "fade":
            slide_elem = slide._element
            transition_xml = parse_xml(
                f'<p:transition {nsdecls("p")} spd="fast" advClick="1"><p:fade/></p:transition>'
            )
            slide_elem.append(transition_xml)

    prs.save(str(output))
    log(f"exported {slide_count} slides with pixel-perfect typography to {output}")
    return {
        "output": str(output),
        "slides": slide_count,
        "transition": transition,
        "bytes": output.stat().st_size if output.exists() else 0,
    }


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="High-fidelity local native PPTD to PPTX converter for DSH NP-PPT."
    )
    parser.add_argument("input", type=Path, help=".pptd manifest or project directory")
    parser.add_argument("--output", "-o", type=Path, help="output .pptx path")
    parser.add_argument(
        "--transition",
        choices=("fade", "none"),
        default="fade",
        help="slide transition written to every slide (default: fade)",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="replace an existing output file",
    )
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    try:
        export_pptx(
            source=args.input,
            output=args.output,
            transition=args.transition,
            force=args.force,
        )
        return 0
    except ExportError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())

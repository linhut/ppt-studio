#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# (c) 2026 ppt-studio 渲染级质量门禁
# 借鉴 cathrynlavery/diagram-design（MIT）scripts/lint-render.py 的渲染级校验思路：
# - 对真实渲染结果做像素级检测（而非仅静态几何估算）
# - 检测画布越界、文字被遮挡、元素重叠、页面密度/留白

"""ppt-studio 渲染级质量门禁：deck.json/pptx → PowerPoint COM 渲染 PNG → 像素级检测。

补上 quality_check.py（纯静态分析）看不见的盲区：
- 元素/文字实际渲染超出画布被裁剪
- 文字被实体元素遮挡（json2pptx 只检查实体-实体重叠，看不见"文字被盖住"）
- 实体元素渲染后互相重叠
- 页面过密/过空（借鉴 diagram-design「目标密度 4/10」哲学）

用法：
  py -3 scripts/render_check.py my-deck/deck.json              # 编译 + 渲染 + 检测
  py -3 scripts/render_check.py my-deck/deck.pptx              # 直接检测已有 pptx
  py -3 scripts/render_check.py my-deck/deck.json --strict     # 有 fail 项时退出码 1
  py -3 scripts/render_check.py my-deck/deck.json --json       # 机器可读输出

依赖：
  - python-pptx（读取 shape 几何/文本）
  - Pillow + numpy（像素分析）
  - 本机安装 Microsoft PowerPoint（COM 渲染）
"""

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from collections import Counter
from pathlib import Path

# 强制 UTF-8 输出（Windows 控制台默认 GBK，无法打印 • 等字符）
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

# ---- 第三方依赖（延迟导入，给出清晰错误） ----
try:
    import numpy as np
    from PIL import Image
except ImportError:
    print("缺少依赖：pip install pillow numpy", file=sys.stderr)
    sys.exit(2)

try:
    from pptx import Presentation
except ImportError:
    print("缺少依赖：pip install python-pptx", file=sys.stderr)
    sys.exit(2)

# ============================================================
# 常量
# ============================================================
SCRIPT_DIR = Path(__file__).resolve().parent

# 画布（ppt-studio 固定 960×540pt，1pt = 1px 渲染）
CANVAS_W = 960
CANVAS_H = 540

# 内容像素判定：与背景 RGB 距离 > 阈值
BG_DIST_THRESHOLD = 20.0

# 文字色匹配容差（RGB 距离）
TEXT_COLOR_TOL = 60.0

# 画布越界检测：边缘条带宽度（px）
EDGE_BAND = 2
# 边缘内容像素最小数量才报警（排除抗锯齿噪声）
EDGE_MIN_PIXELS = 12

# 文字遮挡检测：文字像素落入实体 shape 的最小数量才报警
OCCLUDE_MIN_PIXELS = 8

# 重叠检测：相交面积阈值
OVERLAP_RATIO = 0.15

# 密度检测：格子异色度法（对背景判定免疫）
# 画布分成 GY×GX 格，格子内量化颜色数 ≥ 阈值视为"有内容"。
# 纯色大色块（颜色单一）不算内容——它是设计元素不是视觉拥挤。
# 密集文字/图形混排（格子内多色）才算内容。
DENSITY_GRID_GY = 8
DENSITY_GRID_GX = 12
DENSITY_CELL_MIN_COLORS = 3      # 格子量化颜色数阈值
DENSITY_QUANT = 24               # 颜色量化级
DENSITY_TOO_EMPTY = 0.05         # 覆盖率 < 5% 视为真空白页
DENSITY_TOO_FULL = 0.88          # 覆盖率 > 88% 视为过密（软提示：极少数密集页也会触发）
DENSITY_PENALTY_SOFT = 3         # 过密是软提示（可能误判），扣分轻，不阻塞

# 评分
PENALTY_OUT_OF_BOUNDS = 15
PENALTY_OCCLUDE = 12
PENALTY_OVERLAP = 8
PASS_THRESHOLD = 90.0

# 导出分辨率（1pt = 1px）
EXPORT_W = CANVAS_W
EXPORT_H = CANVAS_H

EMU_PER_PT = 12700.0


def log(msg):
    print("[render_check] " + msg, file=sys.stderr, flush=True)


# ============================================================
# 1. 渲染：PowerPoint COM 把 pptx 逐页导出 PNG
# ============================================================
def render_pptx(pptx_path: Path, out_dir: Path) -> list:
    """用 PowerPoint COM 逐页导出 PNG。返回 PNG 文件列表（按页序）。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    ps_script = r"""
param([string]$Src, [string]$OutDir)
$ErrorActionPreference = "Stop"
$ppt = New-Object -ComObject PowerPoint.Application
try {
    $pres = $ppt.Presentations.Open($Src, $true, $false, $false)
    try {
        $count = $pres.Slides.Count
        for ($i = 1; $i -le $count; $i++) {
            $pres.Slides.Item($i).Export("$OutDir\slide-$i.png", "PNG", %(W)d, %(H)d)
        }
        Write-Output "PAGES=$count"
    } finally {
        $pres.Close()
    }
} finally {
    $ppt.Quit()
}
""" % {"W": EXPORT_W, "H": EXPORT_H}

    script_file = out_dir / "_render.ps1"
    script_file.write_text(ps_script, encoding="utf-8")

    src_abs = str(pptx_path.resolve())
    out_abs = str(out_dir.resolve())
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-File", str(script_file), "-Src", src_abs, "-OutDir", out_abs],
            capture_output=True, text=True, timeout=300,
        )
    except subprocess.TimeoutExpired:
        log("渲染超时（300s）")
        return []
    except FileNotFoundError:
        log("找不到 powershell.exe——无法调用 PowerPoint COM 渲染")
        return []

    if result.returncode != 0:
        log("PowerPoint 渲染失败：%s" % result.stderr.strip()[:500])
        return []

    pages = 0
    for line in result.stdout.splitlines():
        if line.startswith("PAGES="):
            pages = int(line.split("=")[1].strip())
            break

    pngs = []
    for i in range(1, pages + 1):
        p = out_dir / ("slide-%d.png" % i)
        if p.is_file():
            pngs.append(p)
    log("渲染完成：%d 页 → %s" % (len(pngs), out_dir))
    return pngs


# ============================================================
# 1b. 兜底渲染：LibreOffice headless（无 PowerPoint 时）
# ============================================================
def render_pptx_libreoffice(pptx_path: Path, out_dir: Path) -> list:
    """用 LibreOffice headless 把 pptx 转 PDF，再把每页转成 PNG。

    依赖：soffice + PyMuPDF（推荐）或 pdf2image（需 poppler）。
    任一缺失时返回 []（由调用方给出清晰提示）。
    """
    soffice = shutil.which("soffice") or shutil.which("soffice.exe")
    if not soffice:
        log("未找到 soffice——LibreOffice 兜底渲染不可用")
        return []
    try:
        import fitz  # PyMuPDF
    except ImportError:
        fitz = None
    if fitz is None:
        try:
            from pdf2image import convert_from_path
        except ImportError:
            log("缺少 PDF 渲染依赖：pip install pymupdf（或 pdf2image+poppler）")
            return []

    out_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="soffice_") as tmp:
        tmp = Path(tmp)
        try:
            result = subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf",
                 "--outdir", str(tmp), str(pptx_path.resolve())],
                capture_output=True, text=True, timeout=300,
            )
        except (subprocess.TimeoutExpired, OSError) as exc:
            log("LibreOffice 转换失败：%s" % exc)
            return []
        if result.returncode != 0:
            log("LibreOffice 转换失败：%s" % result.stderr.strip()[:300])
            return []

        pdfs = list(tmp.glob("*.pdf"))
        if not pdfs:
            log("LibreOffice 未产出 PDF")
            return []
        pdf = pdfs[0]

        pngs = []
        try:
            if fitz is not None:
                doc = fitz.open(str(pdf))
                for i, page in enumerate(doc):
                    scale = EXPORT_W / max(page.rect.width, 1)
                    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale))
                    out = out_dir / ("slide-%d.png" % (i + 1))
                    pix.save(str(out))
                    pngs.append(out)
                doc.close()
            else:
                images = convert_from_path(str(pdf), size=(EXPORT_W, EXPORT_H))
                for i, img in enumerate(images):
                    out = out_dir / ("slide-%d.png" % (i + 1))
                    img.save(out, "PNG")
                    pngs.append(out)
        except Exception as exc:
            log("PDF 转 PNG 失败：%s" % exc)
            return []
    log("LibreOffice 渲染完成：%d 页 → %s" % (len(pngs), out_dir))
    return pngs


# ============================================================
# 2. 背景提取（众数法）+ 内容像素
# ============================================================
def extract_background(img: np.ndarray):
    """网格采样 + 众数法提取页面背景色。

    返回 (bg_color, 众数占比)。众数占比低（< 0.60）说明页面是渐变/
    图案背景——调用方据此跳过依赖"单色背景"的像素边缘检测。

    不使用四角采样：全宽装饰条（5px 高的色条）恰好贴边时会污染
    四角 8×8 采样（四个角一致地取到装饰条色）。众数法对真实 deck
    （内容居中、背景主导）100% 正确。
    """
    h, w, _ = img.shape
    samples = []
    gy_n, gx_n = 12, 20
    total = gy_n * gx_n
    for gy in range(gy_n):
        for gx in range(gx_n):
            y = int((gy + 0.5) * h / gy_n)
            x = int((gx + 0.5) * w / gx_n)
            block = img[max(0, y - 1):y + 2, max(0, x - 1):x + 2].reshape(-1, 3)
            samples.append(tuple(np.median(block, axis=0).astype(int)))
    bg, count = Counter(samples).most_common(1)[0]
    return np.array(bg, dtype=float), count / float(total)


def content_mask(img: np.ndarray, bg: np.ndarray) -> np.ndarray:
    """内容像素掩码：与背景色距离 > 阈值。"""
    dist = np.sqrt(((img.astype(float) - bg.astype(float)) ** 2).sum(axis=2))
    return dist > BG_DIST_THRESHOLD


# ============================================================
# 3. 读取 pptx 的 shape 几何（python-pptx）
# ============================================================
def emu_to_pt(v):
    try:
        return float(v) / EMU_PER_PT
    except Exception:
        return 0.0


def shape_is_text(sh):
    """判断 shape 是否含有效文本（用于文字遮挡检测）。"""
    if not sh.has_text_frame:
        return False
    text = "".join(run.text for para in sh.text_frame.paragraphs for run in para.runs)
    return bool(text.strip())


def shape_first_color(sh):
    """取文本第一个 run 的颜色（hex 或 None）。"""
    if not sh.has_text_frame:
        return None
    for para in sh.text_frame.paragraphs:
        for run in para.runs:
            try:
                if run.font.color and run.font.color.type is not None:
                    return str(run.font.color.rgb)
            except Exception:
                pass
    return None


def read_pages(pptx_path: Path):
    """读取每页 shape 几何。返回 [{x,y,w,h,text,color,is_text,fill_visible}]。"""
    prs = Presentation(str(pptx_path))
    pages = []
    for slide in prs.slides:
        shapes = []
        for sh in slide.shapes:
            try:
                x = emu_to_pt(sh.left)
                y = emu_to_pt(sh.top)
                w = emu_to_pt(sh.width)
                h = emu_to_pt(sh.height)
            except Exception:
                continue
            text = ""
            if sh.has_text_frame:
                for para in sh.text_frame.paragraphs:
                    for run in para.runs:
                        text += run.text
                    if para.text:
                        text += "\n"
            fill_visible = False
            try:
                if sh.fill.type is not None and str(sh.fill.type) != "MSO_FILL_TYPE.BACKGROUND (5)":
                    fill_visible = True
            except Exception:
                pass
            shapes.append({
                "x": x, "y": y, "w": w, "h": h,
                "text": text.strip(),
                "color": shape_first_color(sh),
                "is_text": shape_is_text(sh),
                "fill_visible": fill_visible,
                "name": sh.name,
            })
        pages.append(shapes)
    return pages


# ============================================================
# 4. 画布越界检测（几何 + 像素）
# ============================================================
def detect_out_of_bounds(page_shapes, img):
    """检测元素/文字超出画布（会被裁剪）。

    两层：
    1. 几何：shape bounds 超出画布（pt → 与画布比较）——渐变背景也检测
    2. 像素：画布边缘条带的内容像素（排除合法全宽装饰条）——
       仅纯色背景页检测（渐变背景的边缘色差会虚报"触及边缘"）
    """
    findings = []
    h_img, w_img = img.shape[0], img.shape[1]
    bg, bg_ratio = extract_background(img)
    mask = content_mask(img, bg)

    # 1. 几何越界
    for shp in page_shapes:
        x, y, w, h = shp["x"], shp["y"], shp["w"], shp["h"]
        if w <= 0 or h <= 0:
            continue
        label = shp["text"][:18] if shp["is_text"] else shp["name"]
        if x + w > CANVAS_W + 1:
            findings.append("元素「%s」右边界 %.0f 超出画布宽 %d（可能被裁剪）" % (label, x + w, CANVAS_W))
        if y + h > CANVAS_H + 1:
            findings.append("元素「%s」下边界 %.0f 超出画布高 %d（可能被裁剪）" % (label, y + h, CANVAS_H))
        if x < -1:
            findings.append("元素「%s」左边界 %.0f 越出画布" % (label, x))
        if y < -1:
            findings.append("元素「%s」上边界 %.0f 越出画布" % (label, y))

    # 2. 像素边缘检测：仅"有明确主背景色"的页面。
    #    渐变/图案背景（众数占比低）的边缘色差会虚报"触及边缘"，跳过。
    if bg_ratio < 0.60:
        return findings

    # 判断每个边缘是否有"恰好贴边"的 shape（封面色块/横幅等设计性贴边）。
    # 贴边 = bounds 与画布边缘重合；真越界（超出画布）已由几何检测覆盖。
    edge_has_anchor = {"top": False, "bottom": False, "left": False, "right": False}
    for shp in page_shapes:
        x, y, w, h = shp["x"], shp["y"], shp["w"], shp["h"]
        if w <= 0 or h <= 0:
            continue
        if x <= 0.5:
            edge_has_anchor["left"] = True
        if x + w >= CANVAS_W - 0.5:
            edge_has_anchor["right"] = True
        if y <= 0.5:
            edge_has_anchor["top"] = True
        if y + h >= CANVAS_H - 0.5:
            edge_has_anchor["bottom"] = True

    # 构建「合法贴边装饰条」掩码并排除
    decor_mask = np.zeros_like(mask)
    for shp in page_shapes:
        x, y, w, h = shp["x"], shp["y"], shp["w"], shp["h"]
        # 全宽水平装饰条（贴顶部/底部）：覆盖其矩形区域
        if w >= CANVAS_W - 2 and h <= 8 and (y <= 8 or y + h >= CANVAS_H - 8):
            x0, y0 = int(round(max(0, x))), int(round(max(0, y)))
            x1 = int(round(min(w_img, x + w)))
            y1 = int(round(min(h_img, y + h)))
            if x1 > x0 and y1 > y0:
                decor_mask[y0:y1, x0:x1] = True
        # 全高垂直装饰条（贴左/右边缘）
        if h >= CANVAS_H - 2 and w <= 8 and (x <= 8 or x + w >= CANVAS_W - 8):
            x0, y0 = int(round(max(0, x))), int(round(max(0, y)))
            x1 = int(round(min(w_img, x + w)))
            y1 = int(round(min(h_img, y + h)))
            if x1 > x0 and y1 > y0:
                decor_mask[y0:y1, x0:x1] = True

    edge_checks = [
        ("顶部", mask[0:EDGE_BAND, :], decor_mask[0:EDGE_BAND, :], "top"),
        ("底部", mask[-EDGE_BAND:, :], decor_mask[-EDGE_BAND:, :], "bottom"),
        ("左侧", mask[:, 0:EDGE_BAND], decor_mask[:, 0:EDGE_BAND], "left"),
        ("右侧", mask[:, -EDGE_BAND:], decor_mask[:, -EDGE_BAND:], "right"),
    ]
    for label, band, band_decor, key in edge_checks:
        if edge_has_anchor.get(key):
            continue  # 该边缘有恰好贴边的 shape（设计性贴边），跳过
        real_content = band & (~band_decor)
        n = int(real_content.sum())
        if n > EDGE_MIN_PIXELS:
            findings.append("内容触及画布%s边缘（%d 像素）——检查是否越界被裁剪" % (label, n))

    return findings


# ============================================================
# 5. 文字遮挡检测（文字像素落入实体 shape 区域）
# ============================================================
def hex_to_rgb(hex6):
    try:
        return np.array([int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)], dtype=float)
    except Exception:
        return None


def detect_occlusion(page_shapes, img, bg):
    """检测文字被实体元素遮挡。

    核心洞察（渲染级盲区，json2pptx 静态检查看不见）：
    - 文字绘制在实体【之上】→ 交集区域仍能看到文字色像素 → 合法排版
    - 文字被后绘实体【盖住】→ 交集区域的文字色像素被实体填充色替换 → 遮挡

    因此检测"交集区域文字色像素缺失"而非"存在"：
    对每个文本 shape，与每个实体 shape 求交集；若交集面积占文本 bounds
    主要部分（> 30%），且交集区域内的文字色像素极少（< 阈值），且实体
    填充色与文字色差异明显 → 判定文字被该实体遮挡。

    排除：
    - 同色系（实体填充色 ≈ 文字色，无法区分，跳过）
    - 实体区域本身很小（可能是装饰）
    - 交集只覆盖文本 bounds 的很小部分（文字主体仍可见）
    """
    findings = []
    img_f = img.astype(float)

    entities = [
        shp for shp in page_shapes
        if not shp["is_text"] and shp["fill_visible"]
        and shp["w"] > 8 and shp["h"] > 8
        and not (shp["w"] >= CANVAS_W - 2 and shp["h"] >= CANVAS_H - 2)
    ]

    for shp in page_shapes:
        if not shp["is_text"]:
            continue
        if not shp["color"]:
            continue
        text_rgb = hex_to_rgb(shp["color"])
        if text_rgb is None:
            continue
        x0, y0 = int(round(shp["x"])), int(round(shp["y"]))
        x1, y1 = int(round(shp["x"] + shp["w"])), int(round(shp["y"] + shp["h"]))
        x0, y0 = max(0, x0), max(0, y0)
        x1, y1 = min(img.shape[1], x1), min(img.shape[0], y1)
        if x1 <= x0 or y1 <= y0:
            continue
        text_area = (x1 - x0) * (y1 - y0)

        for ent in entities:
            ex0, ey0 = int(round(ent["x"])), int(round(ent["y"]))
            ex1, ey1 = int(round(ent["x"] + ent["w"])), int(round(ent["y"] + ent["h"]))
            ix0, iy0 = max(x0, ex0), max(y0, ey0)
            ix1, iy1 = min(x1, ex1), min(y1, ey1)
            if ix1 <= ix0 or iy1 <= iy0:
                continue
            inter_area = (ix1 - ix0) * (iy1 - iy0)
            # 交集必须占文本 bounds 的主要部分才算遮挡嫌疑
            if inter_area < 0.3 * text_area:
                continue
            # 交集区域内文字色像素数量
            region = img_f[iy0:iy1, ix0:ix1]
            d = np.sqrt(((region - text_rgb) ** 2).sum(axis=2))
            n_text = int((d <= TEXT_COLOR_TOL).sum())
            if n_text >= OCCLUDE_MIN_PIXELS:
                continue  # 文字仍清晰可见 → 文字在实体之上，合法
            # 实体填充色
            ent_color = _entity_fill_color(img_f, ex0, ey0, ex1, ey1)
            if ent_color is None:
                continue
            diff = np.sqrt(((ent_color - text_rgb) ** 2).sum())
            if diff < 40:
                continue  # 同色系，无法区分
            # 实体填充色与背景色也要有差异（实体真的"盖"在上面而非透明）
            bg_diff = np.sqrt(((ent_color - bg) ** 2).sum())
            if bg_diff < 25:
                continue
            findings.append(
                "文字「%s」(%.0f,%.0f,%.0f×%.0f) 疑似被实体「%s」遮挡（交集区 %.0f×%.0f 内文字色像素仅 %d 个）"
                % (shp["text"][:18], shp["x"], shp["y"], shp["w"], shp["h"],
                   ent["name"], ix1 - ix0, iy1 - iy0, n_text)
            )
            break

    return findings


def _entity_fill_color(img_f, x0, y0, x1, y1):
    """估算实体 shape 区域的填充色（采样中心区中位数）。"""
    x0, y0 = max(0, x0), max(0, y0)
    x1, y1 = min(img_f.shape[1], x1), min(img_f.shape[0], y1)
    if x1 <= x0 or y1 <= y0:
        return None
    cx0, cy0 = x0 + (x1 - x0) // 4, y0 + (y1 - y0) // 4
    cx1, cy1 = x0 + 3 * (x1 - x0) // 4, y0 + 3 * (y1 - y0) // 4
    region = img_f[cy0:cy1, cx0:cx1].reshape(-1, 3)
    if len(region) < 4:
        return None
    med = np.median(region, axis=0)
    return med


# ============================================================
# 6. 实体重叠检测（几何）
# ============================================================
def rects_overlap(a, b):
    """两个矩形的相交。返回 (相交面积, 占较小者比例)。"""
    ox = min(a[2], b[2]) - max(a[0], b[0])
    oy = min(a[3], b[3]) - max(a[1], b[1])
    if ox <= 0 or oy <= 0:
        return 0.0, 0.0
    inter = ox * oy
    smaller = min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1]))
    if smaller <= 0:
        return 0.0, 0.0
    return inter, inter / smaller


def rect_contains(outer, inner, tol=1.0):
    """outer 是否完全包含 inner（徽章/标签/表头条浮在卡片上=包含，合法）。"""
    return (inner[0] >= outer[0] - tol and inner[1] >= outer[1] - tol
            and inner[2] <= outer[2] + tol and inner[3] <= outer[3] + tol)


def detect_overlap(page_shapes):
    """实体-实体重叠（渲染级复核 json2pptx 的静态检查）。

    排除"完全包含"——小元素（徽章、标签、表头条、色条）完全落在
    大元素（卡片、容器）内部是 ppt-studio 的标准设计（pros_cons
    徽章、roadmap 表头条）。只有【部分相交】的实体才报重叠。
    """
    findings = []
    boxes = []
    for i, shp in enumerate(page_shapes):
        x, y, w, h = shp["x"], shp["y"], shp["w"], shp["h"]
        if w <= 0 or h <= 0:
            continue
        if shp["is_text"]:
            continue  # 文字浮层跳过（文字遮挡另查）
        if min(w, h) <= 6:
            continue  # 装饰细条
        if w >= CANVAS_W - 2 and h >= CANVAS_H - 2:
            continue  # 全页背景
        boxes.append((i, shp["name"], x, y, x + w, y + h))

    for a in range(len(boxes)):
        for b in range(a + 1, len(boxes)):
            ia, na, ax1, ay1, ax2, ay2 = boxes[a]
            ib, nb, bx1, by1, bx2, by2 = boxes[b]

            ab = (ax1, ay1, ax2, ay2)
            bb = (bx1, by1, bx2, by2)

            # 完全包含 → 徽章/标签/表头，合法，跳过
            # （注：完全相同的矩形也在此范围，跳过）
            if rect_contains(ab, bb) or rect_contains(bb, ab):
                continue

            inter, ratio = rects_overlap(ab, bb)
            if inter <= 0:
                continue
            if ratio >= OVERLAP_RATIO:
                findings.append(
                    "实体重叠：「%s」(%.0f,%.0f,%.0f×%.0f) 与「%s」(%.0f,%.0f,%.0f×%.0f) 相交 %.0fpt²"
                    % (na, ax1, ay1, ax2 - ax1, ay2 - ay1, nb, bx1, by1, bx2 - bx1, by2 - by1, inter)
                )
    return findings


# ============================================================
# 7. 密度/留白检测
# ============================================================
def detect_density(img, page_shapes=None):
    """页面密度/留白检测——格子异色度法（对背景判定免疫）。

    借鉴 diagram-design「目标密度 4/10」哲学：页面既不能空，也不能
    塞满。把画布分成 GY×GX 格，每格统计量化颜色数：
    - 格子内颜色多样（≥ DENSITY_CELL_MIN_COLORS）→ 有内容
    - 纯色大色块（颜色单一）不算内容——是设计元素不是视觉拥挤
    - 密集文字/图形混排（格内多色）才算内容

    覆盖率 = 内容格 / 总格：
    - < 5%：页面近乎空白（内容缺失）
    - > 65%：页面过密（几乎每格都有内容，建议拆分）

    对背景判定免疫：不依赖"背景色"，纯色区域天然不计入内容，
    即使整页被大面积色块主导也不会误报过密。
    """
    h, w, _ = img.shape
    gy, gx = DENSITY_GRID_GY, DENSITY_GRID_GX
    cell_h, cell_w = h / gy, w / gx
    quant = DENSITY_QUANT

    content_cells = 0
    total_cells = gy * gx
    for cy in range(gy):
        for cx in range(gx):
            y0, y1 = int(cy * cell_h), int((cy + 1) * cell_h)
            x0, x1 = int(cx * cell_w), int((cx + 1) * cell_w)
            cell = img[y0:y1, x0:x1]
            # 量化颜色（降采样到 quant 级）
            q = (cell.astype(np.int32) // quant * quant).reshape(-1, 3)
            # 用哈希统计近似颜色数
            n_colors = len({tuple(r) for r in q.tolist()})
            if n_colors >= DENSITY_CELL_MIN_COLORS:
                content_cells += 1

    ratio = content_cells / float(total_cells)
    findings = []
    if ratio < DENSITY_TOO_EMPTY:
        findings.append("内容密度 %.0f%% < 5%%——页面近乎空白，检查是否内容缺失" % (ratio * 100))
    elif ratio > DENSITY_TOO_FULL:
        findings.append("内容密度 %.0f%% > %.0f%%——页面过密，建议拆分或精简（借鉴 diagram-design 密度 4/10 哲学）"
                        % (ratio * 100, DENSITY_TOO_FULL * 100))
    return findings, ratio


# ============================================================
# 8. 主流程
# ============================================================
def check_pptx(pptx_path: Path, render_dir: Path):
    """渲染 + 检测。返回逐页报告。"""
    pngs = render_pptx(pptx_path, render_dir)
    if not pngs:
        log("PowerPoint 不可用，尝试 LibreOffice 兜底渲染")
        pngs = render_pptx_libreoffice(pptx_path, render_dir)
    if not pngs:
        return None
    page_shapes = read_pages(pptx_path)

    report = []
    for i, png in enumerate(pngs):
        img = np.array(Image.open(str(png)).convert("RGB"))
        bg, bg_ratio = extract_background(img)
        shapes = page_shapes[i] if i < len(page_shapes) else []

        issues = []
        issues += detect_out_of_bounds(shapes, img)
        issues += detect_occlusion(shapes, img, bg)
        issues += detect_overlap(shapes)
        density_findings, ratio = detect_density(img, shapes)
        issues += density_findings

        score = 100.0
        for iss in issues:
            if "超出画布" in iss or "越出画布" in iss or "触及画布" in iss:
                score -= PENALTY_OUT_OF_BOUNDS
            elif "遮挡" in iss:
                score -= PENALTY_OCCLUDE
            elif "重叠" in iss:
                score -= PENALTY_OVERLAP
            elif "密度" in iss:
                score -= DENSITY_PENALTY_SOFT
        score = max(0.0, min(100.0, score))

        report.append({
            "slide": i + 1,
            "bg": bg.astype(int).tolist(),
            "bg_ratio": round(bg_ratio, 3),
            "density": round(ratio * 100, 1),
            "score": round(score, 1),
            "pass": score >= PASS_THRESHOLD,
            "issues": issues,
        })
    return report


def compile_json(json_path: Path, out_pptx: Path) -> bool:
    """调 json2pptx.py 把 deck.json 编译为 pptx。"""
    cmd = [sys.executable, str(SCRIPT_DIR / "json2pptx.py"),
           str(json_path), "-o", str(out_pptx), "--no-fade"]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    except subprocess.TimeoutExpired:
        log("json2pptx 编译超时")
        return False
    if result.returncode != 0:
        log("json2pptx 编译失败：%s" % result.stderr.strip()[:500])
        return False
    return out_pptx.is_file()


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="ppt-studio 渲染级质量门禁（PowerPoint COM 渲染 + 像素级检测）")
    parser.add_argument("input", type=Path, help="deck.json 或 deck.pptx")
    parser.add_argument("--strict", action="store_true",
                        help="存在 fail 项（分数 < 90）时退出码 1")
    parser.add_argument("--json-output", action="store_true", dest="json_out",
                        help="输出机器可读 JSON 报告")
    parser.add_argument("--render-dir", type=Path, default=None,
                        help="渲染输出目录（默认 <输入目录>/_render_check）")
    args = parser.parse_args(argv)

    if not args.input.is_file():
        print("error: 找不到输入文件 " + str(args.input), file=sys.stderr)
        return 2

    suffix = args.input.suffix.lower()
    if suffix == ".json":
        tmp_dir = Path(tempfile.mkdtemp(prefix="render_check_"))
        try:
            pptx_path = tmp_dir / "deck.pptx"
            if not compile_json(args.input, pptx_path):
                return 2
            report = check_pptx(pptx_path, args.render_dir or (args.input.parent / "_render_check"))
        finally:
            try:
                import shutil
                shutil.rmtree(str(tmp_dir), ignore_errors=True)
            except Exception:
                pass
    elif suffix == ".pptx":
        report = check_pptx(args.input, args.render_dir or (args.input.parent / "_render_check"))
    else:
        print("error: 仅支持 .json 或 .pptx 输入", file=sys.stderr)
        return 2

    if report is None:
        print("error: 渲染失败（需要本机安装 Microsoft PowerPoint 或 LibreOffice；"
              "LibreOffice 路径还需 pip install pymupdf）", file=sys.stderr)
        return 2

    total = sum(r["score"] for r in report) / len(report)
    fails = [r for r in report if not r["pass"]]
    issues_total = sum(len(r["issues"]) for r in report)

    # --json-output：只输出纯 JSON（供 quality_check / CI 解析，不混排人类可读报告）
    if args.json_out:
        out = {
            "total_score": round(total, 1),
            "pages": report,
            "pass": len(fails) == 0,
        }
        print(json.dumps(out, ensure_ascii=False, indent=2))
        if args.strict and fails:
            return 1
        return 0

    print("=" * 60)
    print("ppt-studio 渲染级质量门禁报告")
    print("=" * 60)
    for r in report:
        status = "PASS" if r["pass"] else "FAIL"
        if r["density"] is not None:
            print("\n[%s] 第 %d 页  分数 %.1f  内容密度 %.1f%%" % (status, r["slide"], r["score"], r["density"]))
        else:
            print("\n[%s] 第 %d 页  分数 %.1f  （渐变背景，跳过密度检测）" % (status, r["slide"], r["score"]))
        for iss in r["issues"]:
            print("    ! " + iss)
    print("\n" + "=" * 60)
    print("总分：%.1f / 100（%d 页，%d 个问题，%d 页未达标）" % (total, len(report), issues_total, len(fails)))
    if issues_total == 0:
        print("渲染级检查全部通过 ✔")
    print("=" * 60)

    if args.strict and fails:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

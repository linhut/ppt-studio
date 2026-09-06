#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
# https://github.com/linhut/ppt-studio
# Licensed under the MIT License. See the LICENSE file for details.

"""ppt-studio 设计预设库（DesignPreset）。

借鉴 harness-anything（yb2460/harness-anything，MIT）的 DesignPreset 结构：
colors + fonts + spacing + rules 四元组，让配色不只停留在色板，还能带设计规则。

内置 4 大预设（academic / consultant / business / tech），
并可从 templates/themes/all-themes.yaml 加载 18 套配色为 DesignPreset。

用法：
  py -3 scripts/presets.py list                # 列出预设
  py -3 scripts/presets.py show academic       # 显示某预设细节
"""

import json
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
STUDIO_DIR = SCRIPT_DIR.parent


class DesignPreset:
    """统一设计预设：colors(hex 变量) + fonts + spacing + rules。"""

    def __init__(self, name, display_name, colors, fonts, spacing, rules):
        self.name = name
        self.display_name = display_name
        self.colors = colors      # {"$primary": "#03045e", ...}
        self.fonts = fonts        # {"title": "SimHei", "body": "Microsoft YaHei", ...}
        self.spacing = spacing    # {"margin": 40, "gap": 24, "card_padding": 14, "line_height": 1.3}
        self.rules = rules        # {"visual_ratio": 0.6, "max_colors": 4, ...}

    def to_dict(self):
        return {
            "name": self.name,
            "display_name": self.display_name,
            "colors": self.colors,
            "fonts": self.fonts,
            "spacing": self.spacing,
            "rules": self.rules,
        }

    @classmethod
    def from_colors(cls, name, display_name, colors):
        """从 18 配色（colors 键 $bg/$paper/$primary/...）生成预设，字体与规则用通用值。"""
        return cls(
            name=name,
            display_name=display_name,
            colors=colors,
            fonts={"title": "SimHei", "body": "Microsoft YaHei",
                   "latin_title": "Arial", "latin_body": "Arial"},
            spacing={"margin": 40, "gap": 24, "card_padding": 14, "line_height": 1.3},
            rules={"visual_ratio": 0.5, "max_colors": 5, "colorblind_safe": False,
                   "one_idea_per_slide": True, "white_space": 0.4,
                   "max_bullets": 6, "title_max_len": 16, "prefer_sans_serif": True},
        )


# ============================================================
# 4 大预设（harness-anything design_presets 移植，hex 化）
# ============================================================

PRESETS = {}

PRESETS["academic"] = DesignPreset(
    name="academic", display_name="学术答辩",
    colors={
        "$bg": "#FFFFFF", "$paper": "#F5F8FC", "$primary": "#1A3C8B",
        "$accent": "#E67733", "$ink": "#1A3C8B", "$text": "#222222",
        "$muted": "#666666", "$line": "#D0D8E4", "$green": "#188050",
        "$brass": "#E67733",
    },
    fonts={"title": "SimHei", "body": "Microsoft YaHei",
           "latin_title": "Arial", "latin_body": "Arial"},
    spacing={"margin": 80, "gap": 24, "card_padding": 20, "line_height": 1.5},
    rules={"visual_ratio": 0.65, "max_colors": 5, "colorblind_safe": True,
           "one_idea_per_slide": True, "white_space": 0.4,
           "max_bullets": 6, "title_max_len": 12, "prefer_sans_serif": True},
)

PRESETS["consultant"] = DesignPreset(
    name="consultant", display_name="咨询顾问",
    colors={
        "$bg": "#FFFFFF", "$paper": "#F2F6FA", "$primary": "#003366",
        "$accent": "#00A8E8", "$ink": "#003366", "$text": "#212121",
        "$muted": "#666666", "$line": "#D8E2EC", "$green": "#2E8B57",
        "$brass": "#FF8C00",
    },
    fonts={"title": "SimHei", "body": "Microsoft YaHei",
           "latin_title": "Arial", "latin_body": "Arial"},
    spacing={"margin": 60, "gap": 24, "card_padding": 18, "line_height": 1.4},
    rules={"visual_ratio": 0.6, "max_colors": 4, "colorblind_safe": True,
           "one_idea_per_slide": True, "white_space": 0.45,
           "max_bullets": 6, "title_max_len": 12, "prefer_sans_serif": True},
)

PRESETS["business"] = DesignPreset(
    name="business", display_name="商务汇报",
    colors={
        "$bg": "#FFFFFF", "$paper": "#F4F7FA", "$primary": "#005294",
        "$accent": "#F0A500", "$ink": "#005294", "$text": "#1F2937",
        "$muted": "#6B7280", "$line": "#DCE3EB", "$green": "#10B981",
        "$brass": "#F0A500",
    },
    fonts={"title": "SimHei", "body": "Microsoft YaHei",
           "latin_title": "Arial", "latin_body": "Arial"},
    spacing={"margin": 48, "gap": 20, "card_padding": 16, "line_height": 1.35},
    rules={"visual_ratio": 0.55, "max_colors": 4, "colorblind_safe": True,
           "one_idea_per_slide": True, "white_space": 0.4,
           "max_bullets": 6, "title_max_len": 14, "prefer_sans_serif": True},
)

PRESETS["tech"] = DesignPreset(
    name="tech", display_name="科技极客",
    colors={
        "$bg": "#0F1423", "$paper": "#171E30", "$primary": "#22D3EE",
        "$accent": "#A78BFA", "$ink": "#0F1423", "$text": "#E5E9F0",
        "$muted": "#8A93A6", "$line": "#2A3350", "$green": "#34D399",
        "$brass": "#FBBF24",
    },
    fonts={"title": "SimHei", "body": "Microsoft YaHei",
           "latin_title": "Consolas", "latin_body": "Arial"},
    spacing={"margin": 48, "gap": 20, "card_padding": 16, "line_height": 1.35},
    rules={"visual_ratio": 0.65, "max_colors": 5, "colorblind_safe": False,
           "one_idea_per_slide": True, "white_space": 0.35,
           "max_bullets": 5, "title_max_len": 16, "prefer_sans_serif": True},
)


def load_themes_as_presets():
    """从 templates/themes/all-themes.yaml 加载 18 套配色为 DesignPreset。"""
    import yaml
    yaml_path = STUDIO_DIR / "templates" / "themes" / "all-themes.yaml"
    if not yaml_path.is_file():
        return {}
    data = yaml.safe_load(yaml_path.read_text(encoding="utf-8"))
    out = {}
    for key, item in (data.get("themes") or {}).items():
        colors = item.get("colors", {})
        # 归一化：去掉 $ 前缀差异，统一为 $xxx
        norm = {}
        for k, v in colors.items():
            nk = k if k.startswith("$") else "$" + k
            norm[nk] = v
        out[key] = DesignPreset.from_colors(key, item.get("name", key), norm)
    return out


def get_preset(name):
    """按名称取预设：先查内置 4 大，再查 18 配色。"""
    if name in PRESETS:
        return PRESETS[name]
    themes = load_themes_as_presets()
    return themes.get(name)


def list_presets():
    lines = ["内置预设："]
    for p in PRESETS.values():
        lines.append("  %-12s %s  主色 %s" % (p.name, p.display_name, p.colors.get("$primary")))
    themes = load_themes_as_presets()
    if themes:
        lines.append("18 套配色（templates/themes/all-themes.yaml）：")
        for key, p in themes.items():
            lines.append("  %-24s %s  主色 %s" % (key, p.display_name, p.colors.get("$primary")))
    return "\n".join(lines)


def main(argv=None):
    args = argv if argv is not None else sys.argv[1:]
    if not args or args[0] == "list":
        print(list_presets())
        return 0
    if args[0] == "show":
        p = get_preset(args[1])
        if not p:
            print("未找到预设: " + str(args[1]))
            return 1
        print(json.dumps(p.to_dict(), ensure_ascii=False, indent=2))
        return 0
    print(__doc__)
    return 2


if __name__ == "__main__":
    sys.exit(main())

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
# https://github.com/linhut/ppt-studio
# Licensed under the MIT License. See the LICENSE file for details.

"""ppt-studio 新 deck 脚手架（ppt new）。

用法：
  ppt new my-deck --theme tech --pages 8 --dir .
  ppt new my-deck --theme 15-pure-tech-blue --pages 12

生成 my-deck/deck.json（封面/目录/N 页内容/致谢，内容页按场景轮换布局），
主题色直接套用设计预设（ppt presets list 可查全部 key）。
"""

import argparse
import json
import sys
from pathlib import Path

STUDIO_DIR = Path(__file__).resolve().parent.parent

DEFAULT_LAYOUTS = ["cards", "kpi", "process", "timeline", "compare"]


def _check_theme(theme):
    sys.path.insert(0, str(STUDIO_DIR / "scripts"))
    try:
        from presets import get_preset
        pr = get_preset(theme)
        if pr is None:
            print("error: 未知预设 %s（查 ppt presets list）" % theme, file=sys.stderr)
            return None
        return theme
    except Exception as exc:
        print("error: 读取预设失败：%s" % exc, file=sys.stderr)
        return None


def build_new(name, theme, pages, target_dir):
    sys.path.insert(0, str(STUDIO_DIR / "scripts"))
    try:
        import wizard as wizard_mod
        from presets import list_presets
    except Exception as exc:
        print("error: 加载引擎模块失败：%s" % exc, file=sys.stderr)
        return 2

    theme = _check_theme(theme)
    if theme is None:
        return 2

    # 内容页数 = 总页数 - 封面 - 目录 - 致谢
    content_n = max(1, min(9, pages - 3))
    layouts = (DEFAULT_LAYOUTS * ((content_n // len(DEFAULT_LAYOUTS)) + 1))[:content_n]

    answers = {
        "topic": name,
        "title": name,
        "subtitle": "ppt-studio 生成的演示文稿",
        "palette": theme,
        "style": "soft",
        "mode": "light",
        "tone": "corporate",
        "structure": "full",
        "outline": ["章节 %d" % (i + 1) for i in range(content_n)],
        "content_layouts": layouts,
    }
    deck = wizard_mod.build_deck(answers)

    # 重排页号，精确到 pages 页：封面 + 目录 + 内容 + 致谢
    slides = deck["slides"]
    keep = [slides[0], slides[1]] + slides[2:-1][:content_n] + [slides[-1]]
    for i, s in enumerate(keep):
        s["id"] = "%02d" % (i + 1)
    deck["slides"] = keep

    out_dir = Path(target_dir) / name
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / "deck.json"
    out.write_text(json.dumps(deck, ensure_ascii=False, indent=2), encoding="utf-8")

    print("已创建 %s/deck.json（%d 页，预设 %s，内容布局 %s）"
          % (out_dir, len(keep), theme, "/".join(layouts)))
    print("下一步：")
    print("  ppt build %s -o %s/deck.pptx" % (out, out_dir))
    print("  ppt check %s --preset %s" % (out, theme))
    return 0


def main(argv=None):
    parser = argparse.ArgumentParser(description="ppt-studio 新 deck 脚手架")
    parser.add_argument("name", help="项目目录名（如 my-deck）")
    parser.add_argument("--theme", default="tech", help="设计预设名（默认 tech，ppt presets list 可查）")
    parser.add_argument("--pages", type=int, default=8, help="页数（默认 8）")
    parser.add_argument("--dir", default=".", help="创建位置（默认当前目录）")
    args = parser.parse_args(argv)
    if args.pages < 4:
        print("error: 页数至少 4（封面/目录/内容/致谢）", file=sys.stderr)
        return 2
    return build_new(args.name, args.theme, args.pages, args.dir)


if __name__ == "__main__":
    sys.exit(main())
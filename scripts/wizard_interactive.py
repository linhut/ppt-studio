#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
# https://github.com/linhut/ppt-studio
# Licensed under the MIT License. See the LICENSE file for details.

"""ppt-studio 交互式向导（ppt wizard）。

按 SKILL.md「向导式需求确认」的 6 问逐项提问（带推荐默认值，直接回车取默认），
生成 answers.json 语义的 deck.json 骨架。

非交互场景：
  ppt wizard --defaults -o deck.json      # 全部取推荐默认值（CI / 脚本可用）
  ppt wizard --answers answers.json -o deck.json   # 预填答案，跳过问答
"""

import argparse
import json
import sys
from pathlib import Path

STUDIO_DIR = Path(__file__).resolve().parent.parent

QUESTIONS = [
    ("topic", "主题内容", "能力展示 / 产品介绍 / 汇报 / 技术分享", "技术分享"),
    ("structure", "页数结构", "8-9 页完整 / 5-6 页精简 / 12+ 大 deck", "8-9 页完整"),
    ("palette", "配色方案", "见 ppt presets list（内置 4 大 + 18 套配色）", "15-pure-tech-blue"),
    ("style", "风格配方", "soft / sharp / rounded / pill", "soft"),
    ("mode", "深浅模式", "light / dark / mixed", "light"),
    ("tone", "语言风格", "corporate 商务 / casual 口语 / tech 极客 / story 叙事 / "
                        "persuasive 销售 / academic 学术 / government 公文", "corporate"),
]

STYLE_MAP = {"soft", "sharp", "rounded", "pill"}
MODE_MAP = {"light", "dark", "mixed"}
TONE_MAP = {"corporate", "casual", "tech", "story", "persuasive", "academic", "government"}


def _prompt(label, choices, default):
    """交互提问；非 TTY 或 --defaults 时直接返回默认值。"""
    if not sys.stdin.isatty():
        return default
    hint = "[%s]" % default
    ans = input("  %s (%s) %s: " % (label, choices, hint)).strip()
    return ans or default


def collect_answers(defaults_only=False):
    answers = {}
    if defaults_only:
        answers = {"topic": "技术分享", "structure": "full",
                   "palette": "15-pure-tech-blue", "style": "soft",
                   "mode": "light", "tone": "corporate"}
    else:
        print("ppt-studio 向导（直接回车取推荐默认值，Ctrl+C 取消）")
        answers["topic"] = _prompt("主题内容", QUESTIONS[0][2], "技术分享")
        structure = _prompt("页数结构", QUESTIONS[1][2], "8-9 页完整")
        answers["structure"] = "mini" if structure.startswith("5") else "full"
        answers["palette"] = _prompt("配色方案", "见 ppt presets list", "15-pure-tech-blue")
        style = _prompt("风格配方", "soft / sharp / rounded / pill", "soft")
        answers["style"] = style if style in STYLE_MAP else "soft"
        mode = _prompt("深浅模式", "light / dark / mixed", "light")
        answers["mode"] = mode if mode in MODE_MAP else "light"
        tone = _prompt("语言风格", "corporate / casual / tech / story / persuasive / academic / government",
                       "corporate")
        answers["tone"] = tone if tone in TONE_MAP else "corporate"

    src = STUDIO_DIR / "templates" / "wizard" / "answers.json"
    try:
        tmpl = json.loads(src.read_text(encoding="utf-8"))
        for k in ("title", "subtitle", "tocTitle", "outline"):
            answers.setdefault(k, tmpl.get(k))
    except Exception:
        answers.setdefault("outline", ["第一章", "第二章", "第三章", "第四章", "第五章", "第六章"])
    return answers


def build(answers, out):
    sys.path.insert(0, str(STUDIO_DIR / "scripts"))
    try:
        import wizard as wizard_mod
    except Exception as exc:
        print("error: 加载 wizard 模块失败：%s" % exc, file=sys.stderr)
        return 2
    deck = wizard_mod.build_deck(answers)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(deck, ensure_ascii=False, indent=2), encoding="utf-8")
    print("向导生成 %d 页骨架 → %s" % (len(deck["slides"]), out))
    return 0


def main(argv=None):
    parser = argparse.ArgumentParser(description="ppt-studio 交互式向导")
    parser.add_argument("--output", "-o", default=None, help="输出 deck.json 路径（默认 deck.json）")
    parser.add_argument("--answers", default=None, help="预填 answers.json（跳过问答）")
    parser.add_argument("--defaults", action="store_true", help="全部取推荐默认值（非交互）")
    args = parser.parse_args(argv)

    if args.answers:
        try:
            answers = json.loads(Path(args.answers).read_text(encoding="utf-8"))
        except Exception as exc:
            print("error: 读取 answers.json 失败：%s" % exc, file=sys.stderr)
            return 2
    else:
        answers = collect_answers(defaults_only=args.defaults)

    out = Path(args.output) if args.output else Path("deck.json")
    return build(answers, out)


if __name__ == "__main__":
    sys.exit(main())
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""render_check.py 回归验证：正常 deck 应 0 误报，对抗性用例应全部命中。

借鉴 diagram-design（MIT）的验证文化：每个 checker 都配 adversarial 测试，
"检查的检查"保证检测器既不会漏报真问题、也不会误报正常页面。

用法：
  py -3 scripts/test_render_check.py          # 跑全部
  py -3 scripts/test_render_check.py --deck   # 只跑正常 deck 回归（0 误报）
  py -3 scripts/test_render_check.py --adversarial  # 只跑对抗性用例（应命中）

退出码：0 = 全部通过；1 = 有失败（误报或漏报）
"""

import argparse
import json
import shutil
import sys
import tempfile
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
import render_check as rc  # noqa: E402

STUDIO = SCRIPT_DIR.parent
EXAMPLES = STUDIO / "examples"

# 正常 deck：不应有任何未达标页（0 误报）
NORMAL_DECKS = [
    "engine-test", "layout-showcase", "gov-report",
    "json-demo", "academic-demo", "csbm-research",
]

# 对抗性用例：应命中指定问题（slide 索引 → 期望问题关键字）
ADVERSARIAL = [
    ("render-check-adversarial/deck.json", {
        0: ["超出画布"],        # 越界
        1: ["遮挡"],            # 文字被遮挡
        3: [],                   # 过空页（密度 8.3%，当前阈值不报——只验证不误报 FAIL）
    }),
    ("render-check-adversarial/grid-dense.json", {
        0: ["过密"],            # 网格密集页
    }),
]


def check_deck(deck_path: Path):
    """编译 + 渲染 + 检测，返回 (report, 退出码)。"""
    tmp = Path(tempfile.mkdtemp(prefix="render_check_reg_"))
    try:
        pptx = tmp / "deck.pptx"
        if not rc.compile_json(deck_path, pptx):
            return None, 2
        report = rc.check_pptx(pptx, tmp / "png")
        return report, 0
    finally:
        shutil.rmtree(str(tmp), ignore_errors=True)


def test_normal_decks():
    """正常 deck 0 误报。"""
    print("=" * 60)
    print("回归验证：正常 deck 应 0 误报")
    print("=" * 60)
    failed = []
    for deck in NORMAL_DECKS:
        jp = EXAMPLES / deck / "deck.json"
        if not jp.is_file():
            print("  SKIP %s（无 deck.json）" % deck)
            continue
        report, code = check_deck(jp)
        if report is None:
            print("  FAIL %s：渲染失败" % deck)
            failed.append(deck)
            continue
        total = sum(r["score"] for r in report) / len(report)
        fails = [r for r in report if not r["pass"]]
        print("  %-18s %2d页 总分 %5.1f  %s" % (
            deck, len(report), total,
            "PASS" if not fails else "FAIL(%d 页)" % len(fails)))
        for f in fails:
            print("    slide-%d:" % f["slide"])
            for iss in f["issues"]:
                print("      ! " + iss)
        if fails:
            failed.append(deck)
    return failed


def test_adversarial():
    """对抗性用例应命中指定问题。"""
    print("\n" + "=" * 60)
    print("对抗性验证：人造缺陷应被命中")
    print("=" * 60)
    failed = []
    for rel, expected in ADVERSARIAL:
        jp = EXAMPLES / rel
        if not jp.is_file():
            print("  SKIP %s（不存在）" % rel)
            continue
        report, code = check_deck(jp)
        if report is None:
            print("  FAIL %s：渲染失败" % rel)
            failed.append(rel)
            continue
        for slide_idx, keywords in expected.items():
            if slide_idx >= len(report):
                print("  FAIL %s：缺 slide-%d" % (rel, slide_idx + 1))
                failed.append(rel)
                continue
            r = report[slide_idx]
            issues_text = " ".join(r["issues"])
            if not keywords:
                # 期望无特定问题——只要求页是 PASS
                if not r["pass"]:
                    print("  FAIL %s slide-%d：意外问题：%s"
                          % (rel, slide_idx + 1, issues_text))
                    failed.append(rel)
                else:
                    print("  PASS %s slide-%d（%s）" % (rel, slide_idx + 1, r["issues"][0][:40] if r["issues"] else "无问题"))
                continue
            hit = [k for k in keywords if k in issues_text]
            if hit:
                print("  PASS %s slide-%d 命中：%s" % (rel, slide_idx + 1, "、".join(hit)))
            else:
                print("  FAIL %s slide-%d 漏报！期望 %s，实际：%s"
                      % (rel, slide_idx + 1, "、".join(keywords), issues_text or "无问题"))
                failed.append(rel)
    return failed


def main():
    parser = argparse.ArgumentParser(description="render_check 回归验证")
    parser.add_argument("--deck", action="store_true", help="只跑正常 deck 回归")
    parser.add_argument("--adversarial", action="store_true", help="只跑对抗性用例")
    args = parser.parse_args()

    failures = []
    if not args.adversarial:
        failures += test_normal_decks()
    if not args.deck:
        failures += test_adversarial()

    print("\n" + "=" * 60)
    if failures:
        print("验证失败：%s" % "、".join(failures))
        print("=" * 60)
        return 1
    print("全部验证通过 ✔（正常 deck 0 误报 + 对抗性全部命中）")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())

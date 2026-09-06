#!/usr/bin/env python3
# (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
# https://github.com/linhut/ppt-studio
# Licensed under the MIT License. See the LICENSE file for details.

"""ppt-studio 辅助脚本：把 .page 文件里的 $theme 颜色变量内联为实际色值。

背景：dsh-np-ppt 编译器的 55173 一致性校验要求 background.color 必须是内联 #HEX/rgb()，
不能用 $theme 变量（编辑器无法解析）；content.text 富文本内颜色也必须内联。
fill/border/content.color 用 $theme 是允许的，但统一内联后更保险、更可移植。

本脚本读取 deck.deck.pptd 的 theme.colors，批量把 pages/*.page 中所有
形如  color: "$xxx"  的写法替换为实际色值，再编译即可通过一致性校验。

用法：
  py -3 scripts/inline_bg.py deck/deck.deck.pptd        # 原地修改页面文件
  py -3 scripts/inline_bg.py deck/deck.deck.pptd --dry  # 只显示将修改哪些文件
"""

import sys
import re
from pathlib import Path


def load_manifest(manifest_path: Path):
    import yaml
    with manifest_path.open(encoding="utf-8") as f:
        return yaml.safe_load(f)


def inline_page_colors(page_file: Path, colors: dict) -> bool:
    text = page_file.read_text(encoding="utf-8")
    original = text

    def repl(m):
        name = m.group(2)
        if name in colors:
            return m.group(1) + colors[name] + m.group(3)
        return m.group(0)

    q = chr(39)
    pattern = re.compile(r'(?m)^(\s{2,}color:\s*["' + q + r'])\$([A-Za-z_][A-Za-z0-9_]*)(["' + q + r'])')
    text = pattern.sub(repl, text)
    if text != original:
        page_file.write_text(text, encoding="utf-8")
        return True
    return False


def main(argv) -> int:
    if len(argv) < 1:
        print(__doc__)
        return 2
    dry = "--dry" in argv
    manifest_path = Path(argv[0]).resolve()
    root = manifest_path.parent
    manifest = load_manifest(manifest_path)
    raw_colors = manifest.get("theme", {}).get("colors", {})
    # YAML 里 colors 的 key 带 $ 前缀（如 $bg），统一去掉便于匹配
    colors = {str(k).lstrip("$"): v for k, v in raw_colors.items()}
    if not colors:
        print("[inline_bg] theme.colors empty, nothing to do")
        return 0

    changed = 0
    for page_rel in manifest.get("pages", []):
        page_file = root / page_rel
        if not page_file.is_file():
            print("[inline_bg] skip missing: " + page_rel)
            continue
        if inline_page_colors(page_file, colors):
            changed += 1
            print("[inline_bg] " + ("would update " if dry else "updated ") + page_rel)
    print("[inline_bg] %d file(s) %s" % (changed, "to update" if dry else "updated"))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
# https://github.com/linhut/ppt-studio
# Licensed under the MIT License. See the LICENSE file for details.

"""ppt-studio 统一命令行入口（ppt）。

一条命令覆盖 JSON/PPTD 双引擎的全流程：
  build       JSON → PPTX（json2pptx.py）
  export      PPTD → PPTX（export_pptx.py）
  check       质量审查（quality_check.py）
  render      渲染级检查（render_check.py）
  presets     查预设（presets.py）
  doctor      环境预检（依赖 / PowerPoint / LibreOffice）
  new         新 deck 脚手架
  wizard      交互式向导（answers 问答 → deck.json）
  md2deck     Markdown 大纲 → deck.json
  validate    deck.json schema 校验（JSONPath 定位）
  pptx2json   PPTX → deck.json 反向导入
  version     版本信息

用法示例：
  ppt build my-deck/deck.json -o my-deck/deck.pptx --strict
  ppt check my-deck/deck.json --preset tech --auto-only --html
  ppt render my-deck/deck.json --strict
  ppt export deck.deck.pptd --output deck.pptx --force
  ppt new my-deck --theme tech --pages 8
  ppt md2deck outline.md -o my-deck/deck.json
"""

import argparse
import importlib
import subprocess
import sys
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent


def _run(script, argv):
    """以子进程调用 scripts/ 下既有脚本，透传参数与退出码。"""
    return subprocess.call([sys.executable, str(SCRIPTS / script)] + list(argv))


def _load(module):
    """懒加载 scripts/ 下新增模块（doctor / md2deck / validate / pptx2json）。"""
    if str(SCRIPTS) not in sys.path:
        sys.path.insert(0, str(SCRIPTS))
    try:
        return importlib.import_module(module)
    except ImportError as exc:
        print("error: 模块 %s 不可用：%s" % (module, exc), file=sys.stderr)
        return None


def cmd_build(args):
    return _run("json2pptx.py", args.args)


def cmd_export(args):
    return _run("export_pptx.py", args.args)


def cmd_check(args):
    return _run("quality_check.py", args.args)


def cmd_render(args):
    return _run("render_check.py", args.args)


def cmd_presets(args):
    return _run("presets.py", args.args)


def cmd_doctor(args):
    mod = _load("doctor")
    if mod is None:
        return 2
    return mod.main(args.args)


def cmd_new(args):
    mod = _load("new_project")
    if mod is None:
        return 2
    argv = [args.name, "--theme", args.theme, "--pages", str(args.pages)]
    if args.dir != ".":
        argv += ["--dir", args.dir]
    return mod.main(argv)


def cmd_wizard(args):
    mod = _load("wizard_interactive")
    if mod is None:
        return 2
    argv = []
    if args.answers:
        argv += ["--answers", args.answers]
    if args.output:
        argv += ["--output", args.output]
    if getattr(args, "defaults", False):
        argv += ["--defaults"]
    return mod.main(argv)


def cmd_md2deck(args):
    mod = _load("md2deck")
    if mod is None:
        return 2
    argv = [args.input]
    if args.output:
        argv += ["-o", args.output]
    argv += ["--theme", args.theme]
    return mod.main(argv)


def cmd_validate(args):
    mod = _load("validate")
    if mod is None:
        return 2
    argv = [args.input]
    if args.strict:
        argv += ["--strict"]
    return mod.main(argv)


def cmd_pptx2json(args):
    mod = _load("pptx2json")
    if mod is None:
        return 2
    argv = [args.input]
    if args.output:
        argv += ["-o", args.output]
    return mod.main(argv)


def cmd_version(args):
    print("ppt-studio 1.0.0 (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)")
    return 0


def build_parser():
    parser = argparse.ArgumentParser(
        prog="ppt",
        description="ppt-studio 统一命令行入口（PPT 全能工坊）",
    )
    sub = parser.add_subparsers(dest="command")

    def add_passthrough(name, help_text):
        p = sub.add_parser(name, help=help_text)
        p.add_argument("args", nargs="*", help="透传给底层脚本的参数")
        return p

    add_passthrough("build", "JSON deck → 原生 PPTX（json2pptx.py）")
    add_passthrough("export", "PPTD deck → 原生 PPTX（export_pptx.py）")
    add_passthrough("check", "质量审查：5 维度打分 + 硬约束（quality_check.py）")
    add_passthrough("render", "渲染级质量门禁（render_check.py，需 PowerPoint 或 LibreOffice）")
    add_passthrough("presets", "查设计预设：list / show <name>（presets.py）")
    add_passthrough("doctor", "环境预检：Python 依赖 / PowerPoint / LibreOffice")

    p = sub.add_parser("new", help="创建新 deck 脚手架")
    p.add_argument("name", help="项目目录名（如 my-deck）")
    p.add_argument("--theme", default="tech", help="设计预设名（默认 tech）")
    p.add_argument("--pages", type=int, default=8, help="页数（默认 8）")
    p.add_argument("--dir", default=None, help="创建位置（默认当前目录）")

    p = sub.add_parser("wizard", help="交互式向导：问答 → deck.json 骨架")
    p.add_argument("--output", "-o", default=None, help="输出 deck.json 路径（默认 deck.json）")
    p.add_argument("--answers", default=None, help="预填 answers.json（跳过问答）")
    p.add_argument("--defaults", action="store_true", help="全部取推荐默认值（非交互）")

    p = sub.add_parser("md2deck", help="Markdown 大纲 → deck.json")
    p.add_argument("input", help="Markdown 文件路径")
    p.add_argument("-o", "--output", default=None, help="输出 deck.json 路径")
    p.add_argument("--theme", default="tech", help="设计预设名（默认 tech）")

    p = sub.add_parser("validate", help="deck.json schema 校验（JSONPath 定位）")
    p.add_argument("input", help="deck.json 路径")
    p.add_argument("--strict", action="store_true", help="存在警告时退出码 1")

    p = sub.add_parser("pptx2json", help="PPTX → deck.json 反向导入")
    p.add_argument("input", help="deck.pptx 路径")
    p.add_argument("-o", "--output", default=None, help="输出 deck.json 路径")

    sub.add_parser("version", help="版本信息")
    return parser


# 直通子命令：参数原样转发给底层脚本，不经 argparse（避免 -o/--preset 等被拦截）
PASSTHROUGH = {
    "build": "json2pptx.py",
    "export": "export_pptx.py",
    "check": "quality_check.py",
    "render": "render_check.py",
    "presets": "presets.py",
}


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv or argv[0] in ("-h", "--help"):
        build_parser().print_help()
        return 0
    command = argv[0]
    if command in PASSTHROUGH:
        return _run(PASSTHROUGH[command], argv[1:])
    if command == "version":
        return cmd_version(None)

    parser = build_parser()
    args = parser.parse_args(argv)
    handlers = {
        "build": cmd_build,
        "export": cmd_export,
        "check": cmd_check,
        "render": cmd_render,
        "presets": cmd_presets,
        "doctor": cmd_doctor,
        "new": cmd_new,
        "wizard": cmd_wizard,
        "md2deck": cmd_md2deck,
        "validate": cmd_validate,
        "pptx2json": cmd_pptx2json,
    }
    handler = handlers.get(args.command)
    if handler is None:
        parser.print_help()
        return 2
    return handler(args)


if __name__ == "__main__":
    sys.exit(main())

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
# https://github.com/linhut/ppt-studio
# Licensed under the MIT License. See the LICENSE file for details.

"""ppt-studio 环境预检（ppt doctor）。

逐项检查并给出修复建议：
  runtime    pyyaml / python-pptx         —— 双引擎编译必需
  render     Pillow / numpy               —— 渲染级检测必需
  channel    PowerPoint(COM) / LibreOffice —— 渲染通道（至少其一可用即可出渲染报告）

退出码：必需项全部就绪返回 0，否则返回 1。
"""

import importlib.util
import shutil
import sys

REQUIRED = [
    ("pyyaml", "yaml", "编译引擎必需"),
    ("python-pptx", "pptx", "编译引擎必需"),
]
OPTIONAL = [
    ("Pillow", "PIL", "渲染级检查必需"),
    ("numpy", "numpy", "渲染级检查必需"),
]


def _module_ok(pkg_name):
    return importlib.util.find_spec(pkg_name) is not None


def _check_powerpoint():
    if sys.platform != "win32":
        return False, "仅 Windows 支持 PowerPoint COM 渲染"
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "PowerPoint.Application"):
            return True, "PowerPoint COM 已注册"
    except OSError:
        return False, "未检测到 PowerPoint（可安装 LibreOffice 兜底渲染）"


def _check_soffice():
    path = shutil.which("soffice") or shutil.which("soffice.exe")
    if path:
        return True, path
    return False, "未安装 LibreOffice（可选：soffice --convert-to png 兜底渲染）"


def main(argv=None):
    lines = []
    fail = False

    lines.append("ppt-studio 环境预检")
    lines.append("-" * 56)
    lines.append("Python %s (%s)" % (sys.version.split()[0], sys.platform))

    lines.append("\n[运行时依赖]")
    for pkg, mod, note in REQUIRED:
        ok = _module_ok(mod)
        fail = fail or not ok
        lines.append("  %-14s %s  %s" % (pkg, "OK " if ok else "MISS", note))
        if not ok:
            lines.append("    → 修复：pip install -r requirements.txt")

    lines.append("\n[渲染检测依赖]")
    for pkg, mod, note in OPTIONAL:
        ok = _module_ok(mod)
        lines.append("  %-14s %s  %s" % (pkg, "OK " if ok else "MISS", note))

    lines.append("\n[渲染通道]")
    so_ok, so_note = _check_soffice()
    pp_ok, pp_note = _check_powerpoint()
    lines.append("  %-14s %s  %s" % ("LibreOffice", "OK " if so_ok else "MISS", so_note))
    lines.append("  %-14s %s  %s" % ("PowerPoint", "OK " if pp_ok else "MISS", pp_note))
    if not (so_ok or pp_ok):
        lines.append("    → 至少安装其一，render_check 才能产出渲染级报告")

    lines.append("-" * 56)
    if fail:
        lines.append("结果：存在缺失的必需依赖（见上方 MISS 项）")
    elif so_ok or pp_ok:
        lines.append("结果：环境就绪 ✔（编译 + 渲染全链路可用）")
    else:
        lines.append("结果：环境就绪 ✔（编译可用；渲染通道缺失，仅跳过渲染级检查）")

    print("\n".join(lines))
    return 1 if fail else 0


if __name__ == "__main__":
    sys.exit(main())

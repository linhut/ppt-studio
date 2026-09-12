# -*- coding: utf-8 -*-
# (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
# https://github.com/linhut/ppt-studio
# Licensed under the MIT License. See the LICENSE file for details.

"""pytest 路径注入：确保 tests/ 能导入 scripts/ 下的模块。"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
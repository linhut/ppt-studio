#!/usr/bin/env sh
# ppt-studio unified CLI launcher (macOS / Linux)
# (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
# https://github.com/linhut/ppt-studio
# Licensed under the MIT License. See the LICENSE file for details.
DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
exec python3 "$DIR/scripts/ppt.py" "$@"

@echo off
rem ppt-studio unified CLI launcher (Windows)
rem (c) 2026 Jose AI (https://github.com/linhut/ppt-studio)
rem https://github.com/linhut/ppt-studio
rem Licensed under the MIT License. See the LICENSE file for details.
setlocal
set "ROOT=%~dp0"
where python >nul 2>nul
if %errorlevel%==0 (
  python "%ROOT%scripts\ppt.py" %*
) else (
  py -3 "%ROOT%scripts\ppt.py" %*
)
exit /b %errorlevel%

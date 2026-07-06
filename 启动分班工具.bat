@echo off
setlocal
chcp 65001 >nul
set "SCRIPT_DIR=%~dp0"
set "BUNDLED_PY=C:\Users\yunxi\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if exist "%BUNDLED_PY%" (
  "%BUNDLED_PY%" "%SCRIPT_DIR%class_splitter_gui.py"
  goto :end
)

py -3 "%SCRIPT_DIR%class_splitter_gui.py"
if errorlevel 1 python "%SCRIPT_DIR%class_splitter_gui.py"

:end
endlocal

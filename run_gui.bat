@echo off
set "PY=%~dp0venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"
"%PY%" "%~dp0run_gui.py"

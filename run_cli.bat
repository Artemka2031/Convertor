@echo off
set "VENV=%~dp0venv"
if exist "%VENV%\Scripts\python.exe" (
    "%VENV%\\Scripts\\python.exe" cli.py %*
) else (
    python cli.py %*
)
pause

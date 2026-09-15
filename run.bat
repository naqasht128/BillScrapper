@echo off
title PITC Bill Scrapper & MDI Analyzer
cd /d "%~dp0"
python main.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Application exited with an error. Press any key to close.
    pause >nul
)

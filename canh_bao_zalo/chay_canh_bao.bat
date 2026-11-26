@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Canh bao tre hen FoxOne -^> Zalo
:loop
"%~dp0venv\Scripts\python.exe" foxone_zalo.py
echo [%date% %time%] Script dung, khoi dong lai sau 30 giay... >> foxone_zalo.log
timeout /t 30 /nobreak >nul
goto loop

@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -3 -m pip install -r requirements.txt
if errorlevel 1 echo ERRO: instale Python 3 e marque Add Python to PATH.
pause

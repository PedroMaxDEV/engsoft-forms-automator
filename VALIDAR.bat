@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "DOC=%~1"
if not defined DOC set "DOC=%~dp0MODELO_UNIVERSAL.docx"
py -3 preencher_brave.py "%DOC%" --listar
set /p CAP=Informe capitulo a validar: 
py -3 preencher_brave.py "%DOC%" --cap "%CAP%" --validar
pause

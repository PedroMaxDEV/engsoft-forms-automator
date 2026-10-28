@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "DOC=%~1"
if not defined DOC set "DOC=%~dp0MODELO_UNIVERSAL.docx"
echo.
echo =====================================
echo  EngSoft Forms - escolha seu capitulo
 echo =====================================
py -3 preencher_brave.py "%DOC%" --listar
if errorlevel 1 goto erro
echo.
set /p CAP=Digite o NUMERO do capitulo para preencher: 
if not defined CAP goto erro
py -3 preencher_brave.py "%DOC%" --cap "%CAP%"
if errorlevel 1 goto erro
goto fim
:erro
echo Confira os erros acima. Nenhuma resposta foi enviada automaticamente.
:fim
pause

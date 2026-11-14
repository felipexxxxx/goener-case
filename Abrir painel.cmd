@echo off
cd /d "%~dp0"
echo Salve e feche o Excel antes de continuar.
pause
if not exist "Goener.exe" goto python
Goener.exe --panel
if not errorlevel 1 goto fim
echo O executavel nao iniciou ou nao concluiu. Tentando pelo Python 3.12.
:python
py -3.12 executar.py --panel
:fim
echo.
echo O relatorio salvo fica na pasta resultados.
pause

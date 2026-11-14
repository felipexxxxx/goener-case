@echo off
cd /d "%~dp0"
echo Salve e feche o Excel antes de continuar.
pause
if exist "Goener.exe" (
    Goener.exe --panel
) else (
    py -3.12 executar.py --panel
)
echo.
echo Confira o resultado acima. O relatorio fica na pasta resultados.
pause

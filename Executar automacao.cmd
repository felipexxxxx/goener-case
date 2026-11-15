@echo off
cd /d "%~dp0"
set GOENER_CMD=1
py -3.12 --version >nul 2>&1
if errorlevel 1 (
    echo Instale Python 3.12 e execute novamente.
    pause
    exit /b 1
)
py -3.12 executar.py %*
set GOENER_RESULT=%ERRORLEVEL%
if not "%GOENER_RESULT%"=="0" (
    echo.
    echo A execucao nao concluiu normalmente. Confira o erro e o caminho do log acima.
    pause
)
exit /b %GOENER_RESULT%

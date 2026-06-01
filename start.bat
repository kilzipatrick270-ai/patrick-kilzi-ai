@echo off
setlocal enabledelayedexpansion

echo ==================================================
echo   INSIGHTER AI - Demarrage du serveur
echo ==================================================
echo.

REM Verifier Python
python --version >nul 2>&1
if errorlevel 1 (
    echo XX Python3 n'est pas installe
    echo    Telecharge Python depuis https://www.python.org
    pause
    exit /b 1
)

for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo OK Python %PYTHON_VERSION% detecte
echo.

REM Verifier cle API
if "%ANTHROPIC_API_KEY%"=="" (
    echo AVERTISSEMENT : ANTHROPIC_API_KEY non definie
    echo.
    echo   Definis ta cle API avant de lancer :
    echo   set ANTHROPIC_API_KEY=sk-ant-...
    echo.
    set /p CONTINUE="Continuer sans cle ? (N): "
    if /i not "!CONTINUE!"=="Y" (
        exit /b 1
    )
) else (
    echo OK ANTHROPIC_API_KEY detecte
)

echo.
echo Verification des dependencies...

REM Installer dependencies
pip install -q anthropic flask python-pptx 2>nul
if errorlevel 1 (
    echo.
    echo Installation des dependencies...
    pip install anthropic flask python-pptx
)

echo OK Toutes les dependencies sont OK
echo.
echo ==================================================
echo   SERVEUR EN DEMARRAGE...
echo ==================================================
echo.
echo    Acces : http://localhost:5002
echo    Appuie sur CTRL+C pour arreter
echo.
echo ==================================================
echo.

python app.py

pause

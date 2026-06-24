@echo off
REM build_exe.bat — double-cliquez pour construire BulkArticleEditor.exe
setlocal
cd /d "%~dp0"

REM --- Recherche d'un Python utilisable ---------------------------------
REM On essaie, dans l'ordre : py -3.11, py -3, python, python3.
set "PYCMD="
py -3.11 -c "import sys" >nul 2>&1 && set "PYCMD=py -3.11"
if not defined PYCMD ( py -3 -c "import sys" >nul 2>&1 && set "PYCMD=py -3" )
if not defined PYCMD ( python -c "import sys" >nul 2>&1 && set "PYCMD=python" )
if not defined PYCMD ( python3 -c "import sys" >nul 2>&1 && set "PYCMD=python3" )

if not defined PYCMD (
    echo.
    echo [ERREUR] Aucun Python trouve sur ce poste.
    echo Installez Python 3.11 ^(64 bits recommande^) depuis :
    echo     https://www.python.org/downloads/release/python-3119/
    echo IMPORTANT : cochez "Add python.exe to PATH" pendant l'installation.
    echo.
    pause
    exit /b 1
)

echo == Python utilise : %PYCMD% ==
%PYCMD% -c "import sys; print('   ', sys.version)"

echo == Installation des dependances ==
%PYCMD% -m pip install -r requirements.txt -r requirements-build.txt
if errorlevel 1 (
    echo.
    echo [ERREUR] Echec de l'installation des dependances.
    pause
    exit /b 1
)

echo == Construction de l'executable ==
%PYCMD% -m PyInstaller --noconfirm --clean --onefile --windowed --name BulkArticleEditor bulk_article_editor.py
if errorlevel 1 (
    echo.
    echo [ERREUR] La construction a echoue ^(voir les messages ci-dessus^).
    pause
    exit /b 1
)

echo.
echo Termine. Executable : %~dp0dist\BulkArticleEditor.exe
pause

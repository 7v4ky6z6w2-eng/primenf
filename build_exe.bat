@echo off
REM build_exe.bat — double-cliquez pour construire BulkArticleEditor.exe
REM Concu pour fonctionner avec Python 3.14 (ou 3.11/3.12/3.13).
setlocal
cd /d "%~dp0"

REM --- Recherche d'un Python utilisable ---------------------------------
REM On prend le plus recent dispo : py -3.14, puis py -3, python, py -3.11...
set "PYCMD="
py -3.14 -c "import sys" >nul 2>&1 && set "PYCMD=py -3.14"
if not defined PYCMD ( py -3 -c "import sys" >nul 2>&1 && set "PYCMD=py -3" )
if not defined PYCMD ( python -c "import sys" >nul 2>&1 && set "PYCMD=python" )
if not defined PYCMD ( py -3.11 -c "import sys" >nul 2>&1 && set "PYCMD=py -3.11" )
if not defined PYCMD ( python3 -c "import sys" >nul 2>&1 && set "PYCMD=python3" )

if not defined PYCMD (
    echo.
    echo [ERREUR] Aucun Python trouve sur ce poste.
    echo Installez Python 3.14 ^(64 bits^) depuis https://www.python.org/downloads/
    echo IMPORTANT : cochez "Add python.exe to PATH" pendant l'installation.
    echo.
    pause
    exit /b 1
)

echo == Python utilise : %PYCMD% ==
%PYCMD% -c "import sys; print('   ', sys.version)"

echo == Mise a jour de pip ==
%PYCMD% -m pip install --upgrade pip

echo == Installation/maj des dependances ^(PyInstaller a jour pour 3.14^) ==
%PYCMD% -m pip install --upgrade -r requirements.txt -r requirements-build.txt
if errorlevel 1 (
    echo.
    echo [ERREUR] Echec de l'installation des dependances.
    pause
    exit /b 1
)

echo == Construction de l'executable ==
%PYCMD% -m PyInstaller --noconfirm --clean --onefile --windowed --name BulkArticleEditor ^
    --hidden-import win32ui --hidden-import win32print --hidden-import win32con ^
    bulk_article_editor.py
if errorlevel 1 (
    echo.
    echo [ERREUR] La construction a echoue.
    echo Si le message parle d'une version de Python non supportee, c'est que
    echo votre PyInstaller ne gere pas encore Python 3.14. Deux solutions :
    echo   1^) Reessayez apres  %PYCMD% -m pip install --upgrade pyinstaller
    echo   2^) Ou installez Python 3.12 ^(64 bits^) et relancez ce script.
    echo.
    pause
    exit /b 1
)

echo.
echo Termine. Executable : %~dp0dist\BulkArticleEditor.exe
pause

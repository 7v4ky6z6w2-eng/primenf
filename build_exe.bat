@echo off
setlocal enabledelayedexpansion
REM ============================================================
REM  Construction de l'executable Windows (.exe)
REM  A LANCER SUR WINDOWS (double-clic), avec Python installe.
REM ============================================================
REM  Ce script DETECTE automatiquement une version de Python
REM  installee (3.13, 3.12, 3.11, 3.10, 3.9...). Plus besoin
REM  d'avoir precisement Python 3.11.
REM ------------------------------------------------------------
REM  Pour FORCER une version (ex. 32 bits pour un Firebird 2.5
REM  32 bits), decommentez la ligne ci-dessous :
REM  set "PY=py -3.11-32"
REM ============================================================

REM --- Detection automatique d'un lanceur Python qui fonctionne ---
if not defined PY (
    for %%P in ("py -3.13" "py -3.12" "py -3.11" "py -3.10" "py -3.9" "py -3" "py" "python") do (
        if not defined PY (
            %%~P --version >nul 2>&1 && set "PY=%%~P"
        )
    )
)

if not defined PY (
    echo.
    echo *** Aucun Python trouve. ***
    echo Installez Python depuis https://www.python.org/downloads/
    echo en cochant "Add python.exe to PATH" pendant l'installation.
    echo Versions detectees sur ce poste :
    py -0 2>nul
    python --version 2>nul
    pause
    exit /b 1
)

echo Python utilise : %PY%
%PY% --version
%PY% -c "import struct; print('Architecture :', struct.calcsize('P')*8, 'bits')"

echo.
echo Installation des dependances...
%PY% -m pip install --upgrade pip
%PY% -m pip install pyinstaller -r requirements-gui.txt
if errorlevel 1 goto erreur

echo.
echo Construction de l'executable...
REM --collect-all : openpyxl/fdb sont importes dynamiquement, il faut les
REM forcer car PyInstaller ne les detecte pas tout seul.
%PY% -m PyInstaller --onefile --windowed --name ImportBonReception ^
    --collect-all openpyxl ^
    --collect-all et_xmlfile ^
    --collect-all fdb ^
    --add-data "import_bon_reception.py;." ^
    --add-data "nettoyer_articles.py;." ^
    --add-data "reparer_encodage.py;." ^
    import_bon_reception_gui.py
if errorlevel 1 goto erreur

echo.
echo ============================================================
echo  Termine ! L'executable est dans :  dist\ImportBonReception.exe
echo  Copiez ce SEUL fichier sur les autres PC.
echo ============================================================
pause
goto fin

:erreur
echo.
echo *** ECHEC. Voir les messages ci-dessus. ***
echo Astuce : "%PY% -m pip install pyinstaller" doit fonctionner.
pause

:fin

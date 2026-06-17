@echo off
REM ============================================================
REM  Construction de l'executable Windows (.exe)
REM  A LANCER SUR WINDOWS (double-clic), avec Python installe.
REM ============================================================
REM  IMPORTANT (bitness) : l'exe doit avoir la MEME architecture
REM  que le client Firebird (fbclient.dll) des PC cibles.
REM  Firebird 2.5 est souvent 32 bits -> utilisez un Python 32 bits.
REM ============================================================

REM Lanceur Python (modifiez ici si besoin, ex. py -3.11-32 pour du 32 bits)
set PY=py -3.11

echo Installation des dependances...
%PY% -m pip install --upgrade pip
%PY% -m pip install pyinstaller -r requirements-gui.txt
if errorlevel 1 goto erreur

echo.
echo Construction de l'executable...
%PY% -m PyInstaller --onefile --windowed --name ImportBonReception ^
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
echo *** ECHEC. Verifiez que "py -3.11" fonctionne (Python 3.11 installe). ***
pause

:fin

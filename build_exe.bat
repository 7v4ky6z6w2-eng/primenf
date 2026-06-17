@echo off
REM build_exe.bat — double-cliquez pour construire BulkArticleEditor.exe
cd /d "%~dp0"
py -3.11 -m pip install -r requirements.txt -r requirements-build.txt
py -3.11 -m PyInstaller --noconfirm --clean --onefile --windowed --name BulkArticleEditor bulk_article_editor.py
echo.
echo Executable : %~dp0dist\BulkArticleEditor.exe
pause

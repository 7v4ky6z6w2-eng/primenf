# ============================================================
#  Construction de l'executable Windows (.exe) - PowerShell
#  A lancer DANS le dossier du projet :
#      powershell -ExecutionPolicy Bypass -File build_exe.ps1
#  ou, si l'execution de scripts est autorisee :  .\build_exe.ps1
# ============================================================
#  Bitness : l'exe doit avoir la MEME architecture que le
#  client Firebird (fbclient.dll). Firebird 2.5 est souvent
#  32 bits -> utilisez un Python 32 bits (py -3.11-32).
# ============================================================

# Lanceur Python (changez en "-3.11-32" pour un build 32 bits)
$PY = "-3.11"

Write-Host "== Version de Python ==" -ForegroundColor Cyan
py $PY --version
if ($LASTEXITCODE -ne 0) { Write-Host "Python $PY introuvable." -ForegroundColor Red; exit 1 }
py $PY -c "import struct; print('Python', struct.calcsize('P')*8, 'bits')"

Write-Host "`n== Installation des dependances ==" -ForegroundColor Cyan
py $PY -m pip install --upgrade pip
py $PY -m pip install pyinstaller -r requirements-gui.txt
if ($LASTEXITCODE -ne 0) { Write-Host "Echec d'installation des dependances." -ForegroundColor Red; exit 1 }

Write-Host "`n== Construction de l'executable ==" -ForegroundColor Cyan
# --collect-all : openpyxl/fdb sont importes dynamiquement, il faut les forcer.
py $PY -m PyInstaller --onefile --windowed --name ImportBonReception `
    --collect-all openpyxl `
    --collect-all et_xmlfile `
    --collect-all fdb `
    --add-data "import_bon_reception.py;." `
    --add-data "nettoyer_articles.py;." `
    --add-data "reparer_encodage.py;." `
    import_bon_reception_gui.py
if ($LASTEXITCODE -ne 0) { Write-Host "Echec de la construction." -ForegroundColor Red; exit 1 }

Write-Host "`nTermine ! L'executable est dans : dist\ImportBonReception.exe" -ForegroundColor Green

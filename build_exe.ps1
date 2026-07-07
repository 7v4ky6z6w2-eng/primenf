# ============================================================
#  Construction de l'executable Windows (.exe) - PowerShell
#  A lancer DANS le dossier du projet :
#      powershell -ExecutionPolicy Bypass -File build_exe.ps1
#  ou, si l'execution de scripts est autorisee :  .\build_exe.ps1
# ============================================================
#  Ce script DETECTE automatiquement une version de Python
#  installee (3.13, 3.12, 3.11, 3.10, 3.9...). Plus besoin
#  d'avoir precisement Python 3.11.
#
#  Pour FORCER une version (ex. 32 bits pour un Firebird 2.5
#  32 bits), lancez par ex. :  .\build_exe.ps1 -Py "py -3.11-32"
# ============================================================

param([string]$Py = "")

function Find-Python {
    $candidates = @(
        @('py','-3.14'), @('py','-3.13'), @('py','-3.12'), @('py','-3.11'),
        @('py','-3.10'), @('py','-3.9'), @('py','-3'), @('py'), @('python')
    )
    foreach ($c in $candidates) {
        $exe  = $c[0]
        $rest = @($c[1..($c.Count - 1)])
        try {
            & $exe @rest --version *> $null 2>&1
            if ($LASTEXITCODE -eq 0) { return ,@($exe, $rest) }
        } catch { }
    }
    return $null
}

if ($Py -ne "") {
    $parts = $Py.Split(' ')
    $pyExe = $parts[0]; $pyArgs = @($parts[1..($parts.Count - 1)])
} else {
    $found = Find-Python
    if ($null -eq $found) {
        Write-Host "Aucun Python trouve." -ForegroundColor Red
        Write-Host "Installez Python depuis https://www.python.org/downloads/ (cochez 'Add to PATH')."
        Write-Host "Versions detectees :" -ForegroundColor Yellow
        try { py -0 } catch { }
        exit 1
    }
    $pyExe = $found[0]; $pyArgs = $found[1]
}

Write-Host "== Python utilise ==" -ForegroundColor Cyan
Write-Host ("  " + $pyExe + " " + ($pyArgs -join ' '))
& $pyExe @pyArgs --version
& $pyExe @pyArgs -c "import struct; print('Architecture :', struct.calcsize('P')*8, 'bits')"

Write-Host "`n== Installation des dependances ==" -ForegroundColor Cyan
& $pyExe @pyArgs -m pip install --upgrade pip
& $pyExe @pyArgs -m pip install pyinstaller -r requirements-gui.txt
if ($LASTEXITCODE -ne 0) { Write-Host "Echec d'installation des dependances." -ForegroundColor Red; exit 1 }

Write-Host "`n== Construction de l'executable ==" -ForegroundColor Cyan
# --collect-all : openpyxl/fdb sont importes dynamiquement, il faut les forcer.
& $pyExe @pyArgs -m PyInstaller --onefile --windowed --name ImportBonReception `
    --collect-all openpyxl `
    --collect-all et_xmlfile `
    --collect-all fdb `
    --collect-all pikepdf `
    --collect-all pdfplumber `
    --collect-all pdfminer `
    --collect-all fontTools `
    --add-data "import_bon_reception.py;." `
    --add-data "nettoyer_articles.py;." `
    --add-data "reparer_encodage.py;." `
    import_bon_reception_gui.py
if ($LASTEXITCODE -ne 0) { Write-Host "Echec de la construction." -ForegroundColor Red; exit 1 }

Write-Host "`nTermine ! L'executable est dans : dist\ImportBonReception.exe" -ForegroundColor Green

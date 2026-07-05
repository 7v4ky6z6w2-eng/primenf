# build_exe.ps1 — Construit BulkArticleEditor.exe (Windows, PowerShell)
# Lancement :  .\build_exe.ps1
$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

# --- Recherche d'un Python utilisable ---------------------------------
# On essaie, dans l'ordre : py -3.11, py -3, python, python3.
function Find-Python {
    $candidates = @(
        @("py", @("-3.14")),
        @("py", @("-3")),
        @("python", @()),
        @("py", @("-3.11")),
        @("python3", @())
    )
    foreach ($c in $candidates) {
        $exe = $c[0]; $pre = $c[1]
        if (Get-Command $exe -ErrorAction SilentlyContinue) {
            try {
                & $exe @pre -c "import sys" 2>$null
                if ($LASTEXITCODE -eq 0) { return ,@($exe, $pre) }
            } catch { }
        }
    }
    return $null
}

$py = Find-Python
if ($null -eq $py) {
    Write-Host ""
    Write-Host "[ERREUR] Aucun Python trouve sur ce poste." -ForegroundColor Red
    Write-Host "Installez Python 3.11 (64 bits recommande) depuis :" -ForegroundColor Yellow
    Write-Host "    https://www.python.org/downloads/release/python-3119/" -ForegroundColor Yellow
    Write-Host 'IMPORTANT : cochez "Add python.exe to PATH" pendant l''installation.' -ForegroundColor Yellow
    exit 1
}
$exe = $py[0]; $pre = $py[1]
Write-Host "== Python utilise : $exe $($pre -join ' ') ==" -ForegroundColor Cyan
& $exe @pre -c "import sys; print('   ', sys.version)"

Write-Host "== Mise a jour de pip ==" -ForegroundColor Cyan
& $exe @pre -m pip install --upgrade pip

Write-Host "== Installation/maj des dependances (PyInstaller a jour pour 3.14) ==" -ForegroundColor Cyan
& $exe @pre -m pip install --upgrade -r requirements.txt -r requirements-build.txt

Write-Host "== Construction de l'executable ==" -ForegroundColor Cyan
& $exe @pre -m PyInstaller --noconfirm --clean --onefile --windowed --name BulkArticleEditor `
    --hidden-import win32ui --hidden-import win32print --hidden-import win32con `
    --hidden-import win32gui `
    bulk_article_editor.py

Write-Host ""
Write-Host "Termine. L'executable est dans : $PSScriptRoot\dist\BulkArticleEditor.exe" -ForegroundColor Green
Write-Host "Copiez aussi config.example.json a cote (renomme config.json) si besoin." -ForegroundColor Yellow

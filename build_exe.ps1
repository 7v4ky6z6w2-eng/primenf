# build_exe.ps1 — Construit BulkArticleEditor.exe (Windows, PowerShell)
# Lancement :  py -3.11 -m pip install -r requirements.txt -r requirements-build.txt
#              .\build_exe.ps1
$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

Write-Host "== Installation des dependances ==" -ForegroundColor Cyan
py -3.11 -m pip install -r requirements.txt -r requirements-build.txt

Write-Host "== Construction de l'executable ==" -ForegroundColor Cyan
py -3.11 -m PyInstaller `
    --noconfirm --clean `
    --onefile --windowed `
    --name BulkArticleEditor `
    bulk_article_editor.py

Write-Host ""
Write-Host "Termine. L'executable est dans : $PSScriptRoot\dist\BulkArticleEditor.exe" -ForegroundColor Green
Write-Host "Copiez aussi config.example.json a cote (renomme config.json) si besoin." -ForegroundColor Yellow

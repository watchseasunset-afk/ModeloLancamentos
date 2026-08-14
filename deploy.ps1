param([string]$msg = "deploy automatico")

git add -A
git commit -m $msg
git push origin master:main

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERRO: git push falhou." -ForegroundColor Red
    exit 1
}

Write-Host "Push OK - Railway auto-deploya em ~1 min" -ForegroundColor Green

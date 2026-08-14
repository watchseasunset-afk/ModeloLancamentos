$BASE = "https://lancamentos-laterais.up.railway.app"

# ══════════════════════════════════════════════════════
# Limpar jogo anterior (PSG → done)
# ══════════════════════════════════════════════════════
Write-Host ">> A marcar PSG vs Villa como done..." -ForegroundColor DarkGray
Invoke-RestMethod -Method POST -Uri "$BASE/admin/set-status" `
  -ContentType "application/json" `
  -Body '{"token":"live2026","flash_mid":"GnSGwOrJ","status":"done"}' | ConvertTo-Json

# ══════════════════════════════════════════════════════
# JOGO 1 — Anderlecht vs PAOK  |  19:30 UTC  |  LE Qual
# ══════════════════════════════════════════════════════
Write-Host ">> Jogo 1: Anderlecht vs PAOK..." -ForegroundColor Cyan
Invoke-RestMethod -Method POST -Uri "$BASE/api/live/push" `
  -ContentType "application/json" `
  -Body '{"token":"live2026","game":{
    "flash_mid":"ANDPAOK13",
    "league":"LEQ",
    "home":"Anderlecht",
    "away":"PAOK",
    "kickoff":"2026-08-13T18:30:00+00:00",
    "lanc_baseline":35.0,
    "faltas_baseline":26.0,
    "status":"pending",
    "bet_ci":"358361500",
    "statscore_id":""
  }}' | ConvertTo-Json

# ══════════════════════════════════════════════════════
# JOGO 2 — Hearts vs Benfica  |  19:45 UTC  |  LE Qual
# ══════════════════════════════════════════════════════
Write-Host ">> Jogo 2: Hearts vs Benfica..." -ForegroundColor Cyan
Invoke-RestMethod -Method POST -Uri "$BASE/api/live/push" `
  -ContentType "application/json" `
  -Body '{"token":"live2026","game":{
    "flash_mid":"HRTBEN13",
    "league":"LEQ",
    "home":"Hearts",
    "away":"Benfica",
    "kickoff":"2026-08-13T18:45:00+00:00",
    "lanc_baseline":35.0,
    "faltas_baseline":26.0,
    "status":"pending",
    "bet_ci":"357962992",
    "statscore_id":""
  }}' | ConvertTo-Json

# ══════════════════════════════════════════════════════
# JOGO 3 — Dinamo Minsk vs SC Braga  |  18:00 UTC  |  Conference League Qual
# ══════════════════════════════════════════════════════
Write-Host ">> Jogo 3: Dinamo Minsk vs SC Braga..." -ForegroundColor Cyan
Invoke-RestMethod -Method POST -Uri "$BASE/api/live/push" `
  -ContentType "application/json" `
  -Body '{"token":"live2026","game":{
    "flash_mid":"Zjbslrb",
    "league":"LEQ",
    "home":"Dinamo Minsk",
    "away":"SC Braga",
    "kickoff":"2026-08-13T17:00:00+00:00",
    "lanc_baseline":35.0,
    "faltas_baseline":26.0,
    "status":"pending",
    "bet_ci":"358387744",
    "statscore_id":""
  }}' | ConvertTo-Json

# ══════════════════════════════════════════════════════
Write-Host ">> A forcar collect..." -ForegroundColor Cyan
Invoke-RestMethod -Method POST -Uri "$BASE/admin/collect-now" `
  -ContentType "application/json" `
  -Body '{"token":"live2026"}' | ConvertTo-Json

Write-Host "Feito!" -ForegroundColor Green
Write-Host ""
Write-Host "LEMBRAR: Adicionar statscore_id de cada jogo via Yonibet tracker quando ficarem live" -ForegroundColor Yellow
Write-Host "  Anderlecht vs PAOK    -> abrir Yonibet -> clicar no jogo -> ver URL do tracker iframe -> copiar eventId" -ForegroundColor Yellow
Write-Host "  Hearts vs Benfica     -> idem" -ForegroundColor Yellow
Write-Host "  Dinamo Minsk vs Braga -> idem (jogo 18:00 UTC)" -ForegroundColor Yellow

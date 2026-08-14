$BASE = "https://lancamentos-laterais.up.railway.app"

# ──────────────────────────────────────────────────────────────────────────
# Edita os statscore_id abaixo depois de os encontrares no Yonibet:
#   1. Abrir Yonibet → clicar no jogo → abrir DevTools (F12) → Network
#   2. Filtrar por "statscore" ou "tracker"
#   3. Copiar o eventId da URL do tracker
# ──────────────────────────────────────────────────────────────────────────

$SC_ANDPAOK = "PREENCHER"   # statscore_id Anderlecht vs PAOK
$SC_HRTBEN  = "PREENCHER"   # statscore_id Hearts vs Benfica
$SC_DINBRA  = "PREENCHER"   # statscore_id Dinamo Minsk vs SC Braga — encontrar no Yonibet quando o jogo entrar em live
$BET_DINBRA = "358387744"   # bet_ci 22bet Dinamo Minsk vs SC Braga (confirmado via URL)

Write-Host ">> A actualizar statscore_id Anderlecht vs PAOK..." -ForegroundColor Cyan
Invoke-RestMethod -Method POST -Uri "$BASE/api/live/push" `
  -ContentType "application/json" `
  -Body "{`"token`":`"live2026`",`"game`":{`"flash_mid`":`"ANDPAOK13`",`"league`":`"LEQ`",`"home`":`"Anderlecht`",`"away`":`"PAOK`",`"kickoff`":`"2026-08-13T19:30:00+00:00`",`"lanc_baseline`":35.0,`"faltas_baseline`":26.0,`"status`":`"pending`",`"bet_ci`":`"358361500`",`"statscore_id`":`"$SC_ANDPAOK`"}}" | ConvertTo-Json

Write-Host ">> A actualizar statscore_id Hearts vs Benfica..." -ForegroundColor Cyan
Invoke-RestMethod -Method POST -Uri "$BASE/api/live/push" `
  -ContentType "application/json" `
  -Body "{`"token`":`"live2026`",`"game`":{`"flash_mid`":`"HRTBEN13`",`"league`":`"LEQ`",`"home`":`"Hearts`",`"away`":`"Benfica`",`"kickoff`":`"2026-08-13T19:45:00+00:00`",`"lanc_baseline`":35.0,`"faltas_baseline`":26.0,`"status`":`"pending`",`"bet_ci`":`"357962992`",`"statscore_id`":`"$SC_HRTBEN`"}}" | ConvertTo-Json

Write-Host ">> A actualizar statscore_id Dinamo Minsk vs SC Braga..." -ForegroundColor Cyan
Invoke-RestMethod -Method POST -Uri "$BASE/api/live/push" `
  -ContentType "application/json" `
  -Body "{`"token`":`"live2026`",`"game`":{`"flash_mid`":`"DINBRA13`",`"league`":`"LEQ`",`"home`":`"Dinamo Minsk`",`"away`":`"SC Braga`",`"kickoff`":`"2026-08-13T18:00:00+00:00`",`"lanc_baseline`":35.0,`"faltas_baseline`":26.0,`"status`":`"pending`",`"bet_ci`":`"$BET_DINBRA`",`"statscore_id`":`"$SC_DINBRA`"}}" | ConvertTo-Json

Write-Host ">> A forcar collect..." -ForegroundColor Cyan
Invoke-RestMethod -Method POST -Uri "$BASE/admin/collect-now" `
  -ContentType "application/json" `
  -Body '{"token":"live2026"}' | ConvertTo-Json

Write-Host "Statscore IDs actualizados!" -ForegroundColor Green

$BASE   = "https://lancamentos-laterais.up.railway.app"
$TOKEN  = "live2026"
$FALTASDB = "$PSScriptRoot\..\Claude_Mod_Faltas\Faltas_206_2027\faltas_app_v5\faltas_app\faltas.db"

# ── 1. Verificar status antes do push ─────────────────────────────────────────
Write-Host ">> A verificar árbitros actuais no Railway..." -ForegroundColor DarkGray
try {
    $status = Invoke-RestMethod -Uri "$BASE/admin/arbitros-status" -Method GET
    Write-Host "   Railway tem $($status.total) árbitros em memória" -ForegroundColor DarkGray
} catch {
    Write-Host "   Erro ao verificar status: $_" -ForegroundColor Red
}

# ── 2. Ler referee_alerts do faltas.db local ──────────────────────────────────
Write-Host ">> A ler referee_alerts de $FALTASDB..." -ForegroundColor Cyan

$pyScript = @"
import sqlite3, json, sys
db_path = r'$FALTASDB'
try:
    con = sqlite3.connect(db_path)
    con.row_factory = sqlite3.Row
    rows = con.execute('SELECT * FROM referee_alerts ORDER BY detected_at DESC').fetchall()
    data = [{'home': r['home'], 'away': r['away'], 'referee': r['referee'], 'league': r['league'], 'match_time': r['match_time']} for r in rows]
    print(json.dumps(data))
    con.close()
except Exception as e:
    print(json.dumps({'error': str(e)}), file=sys.stderr)
    sys.exit(1)
"@

$jsonOut = python -c $pyScript 2>$null
if (-not $jsonOut) {
    Write-Host "Erro: sem dados do faltas.db ou Python nao encontrado." -ForegroundColor Red
    exit 1
}

$arbitros = $jsonOut | ConvertFrom-Json
Write-Host "   $($arbitros.Count) nomeacoes encontradas no faltas.db" -ForegroundColor Cyan

if ($arbitros.Count -eq 0) {
    Write-Host "Nenhuma nomeacao para enviar." -ForegroundColor Yellow
    exit 0
}

# Mostrar o que vai ser enviado
$arbitros | ForEach-Object {
    Write-Host "   $($_.home) vs $($_.away) - $($_.referee)" -ForegroundColor DarkGray
}

# ── 3. Push para Railway ──────────────────────────────────────────────────────
$body = @{ token = $TOKEN; arbitros = $arbitros } | ConvertTo-Json -Depth 5 -Compress
try {
    $resp = Invoke-RestMethod -Method POST -Uri "$BASE/admin/arbitros-push" `
        -ContentType "application/json" -Body $body
    Write-Host ""
    Write-Host "OK Push concluido: $($resp.added) nomecoes enviadas, total em Railway: $($resp.total)" -ForegroundColor Green
} catch {
    Write-Host "   Erro no push: $_" -ForegroundColor Red
    exit 1
}

# ── 4. Verificar resultado ────────────────────────────────────────────────────
try {
    $status2 = Invoke-RestMethod -Uri "$BASE/admin/arbitros-status" -Method GET
    Write-Host "   Railway agora tem $($status2.total) arbitros:" -ForegroundColor Cyan
    $status2.arbitros | ForEach-Object {
        Write-Host "     $($_.key) -> $($_.ref)" -ForegroundColor DarkGray
    }
} catch {}

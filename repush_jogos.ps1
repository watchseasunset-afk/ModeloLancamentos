$BASE  = "https://lancamentos-laterais.up.railway.app"
$TOKEN = "live2026"
$DB    = "$PSScriptRoot\lancamentos_lines.db"

# ── 1. Verificar status antes do push ─────────────────────────────────────────
Write-Host ">> A verificar status no Railway..." -ForegroundColor DarkGray
try {
    $status = Invoke-RestMethod -Uri "$BASE/admin/jogos-status" -Method GET
    Write-Host "   Railway tem $($status.total) jogos" -ForegroundColor DarkGray
} catch {
    Write-Host "   Erro ao verificar status: $_" -ForegroundColor Red
}

# ── 2. Ler jogos do DB local via Python ───────────────────────────────────────
Write-Host ">> A ler jogos locais de $DB..." -ForegroundColor Cyan

$pyScript = @"
import sqlite3, json, sys
db = r'$DB'
con = sqlite3.connect(db)
con.row_factory = sqlite3.Row
try:
    rows = con.execute('SELECT * FROM lancamentos_lines ORDER BY kickoff').fetchall()
    jogos = []
    for r in rows:
        d = dict(r)
        # all_lines pode ser string JSON
        try: d['all_lines'] = json.loads(d.get('all_lines') or '[]')
        except: d['all_lines'] = []
        jogos.append(d)
    print(json.dumps(jogos))
except Exception as e:
    print(json.dumps({'error': str(e)}), file=sys.stderr)
    sys.exit(1)
finally:
    con.close()
"@

$jogosJson = python -c $pyScript 2>$null
if (-not $jogosJson) {
    Write-Host "Erro: sem dados do DB local ou Python não encontrado." -ForegroundColor Red
    exit 1
}

$jogos = $jogosJson | ConvertFrom-Json
Write-Host "   $($jogos.Count) jogos encontrados localmente" -ForegroundColor Cyan

if ($jogos.Count -eq 0) {
    Write-Host "Nenhum jogo para enviar." -ForegroundColor Yellow
    exit 0
}

# ── 3. Push para Railway em lotes de 50 ──────────────────────────────────────
$loteSize = 50
$total_added = 0
$total_updated = 0

for ($i = 0; $i -lt $jogos.Count; $i += $loteSize) {
    $lote = $jogos[$i..([Math]::Min($i + $loteSize - 1, $jogos.Count - 1))]
    $body = @{ token = $TOKEN; jogos = $lote } | ConvertTo-Json -Depth 10 -Compress
    try {
        $resp = Invoke-RestMethod -Method POST -Uri "$BASE/admin/jogos-push" `
            -ContentType "application/json" -Body $body
        $total_added   += $resp.added
        $total_updated += $resp.updated
        Write-Host "   Lote $([Math]::Floor($i/$loteSize)+1): +$($resp.added) novos, ~$($resp.updated) actualizados" -ForegroundColor Green
    } catch {
        Write-Host "   Erro no lote: $_" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "✅ Push concluído: $total_added novos, $total_updated actualizados" -ForegroundColor Green
Write-Host ""

# ── 4. Verificar resultado ────────────────────────────────────────────────────
try {
    $status2 = Invoke-RestMethod -Uri "$BASE/admin/jogos-status" -Method GET
    Write-Host "   Railway agora tem $($status2.total) jogos" -ForegroundColor Cyan
    Write-Host "   Próximos jogos:" -ForegroundColor DarkGray
    $status2.jogos | Select-Object -First 5 | ForEach-Object {
        Write-Host "     $($_.home) vs $($_.away) - $($_.kickoff) [$($_.edge_signal)]" -ForegroundColor DarkGray
    }
} catch {}

"""
Colheita Automática de Resultados — Lançamentos Laterais
=========================================================
Corre após os jogos terminarem:
  python colheita_resultados.py           → processa todos os jogos pendentes
  python colheita_resultados.py --dry-run → mostra o que ia fazer sem gravar
  python colheita_resultados.py --horas 3 → só jogos com kickoff há > 3h (default: 2)

Fluxo:
  1. Lê lancamentos_lines.db → jogos sem resultado real
  2. Pesquisa match ID no Flashscore por nome das equipas + data
  3. Usa ninja API para buscar stats reais de lançamentos
  4. Atualiza DB com total real + GREEN/RED automático
"""
import sqlite3, json, re, time, sys, os, urllib.request, urllib.error
from datetime import datetime, timezone, timedelta
from pathlib import Path
from difflib import SequenceMatcher

# ── Configuração ──────────────────────────────────────────────────────────────
DB_PATH    = Path(__file__).parent / 'lancamentos_lines.db'
DELAY      = 1.5   # segundos entre pedidos
HORAS_MIN  = 2     # kickoff há mais de X horas para processar
DRY_RUN    = '--dry-run' in sys.argv

if '--horas' in sys.argv:
    idx = sys.argv.index('--horas')
    HORAS_MIN = int(sys.argv[idx+1])

# ── Headers Flashscore ────────────────────────────────────────────────────────
FS_HDR = {
    'X-Fsign':         'SW9D1eZo',
    'Referer':         'https://www.flashscore.com/',
    'Accept':          '*/*',
    'User-Agent':      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept-Language': 'pt-PT,pt;q=0.9,en;q=0.8',
}
SEARCH_HDR = {
    'Referer':         'https://www.flashscore.com/',
    'Accept':          'application/json, */*',
    'User-Agent':      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'x-requested-with':'XMLHttpRequest',
}
NINJA_URL  = 'https://global.flashscore.ninja/20/x/feed/df_st_1_{mid}'
SEARCH_URL = 'https://s.flashscore.com/search/?q={q}&l=1&s=2&f=1%3B1&pid=2&sid=1'

# Mapeamento de nomes 22bet → Flashscore (ajustar conforme necessário)
NOME_MAP = {
    # PL
    'Manchester Utd':    'Manchester United',
    'Nottingham':        'Nottm Forest',
    'Leeds':             'Leeds United',
    # PPL
    'Fc Porto':          'Porto',
    'Sporting Cp':       'Sporting CP',
    'Estrela Da Amadora':'Estrela Amadora',
    'Famalicao':         'Famalicão',
    'Moreirense':        'Moreirense FC',
    # BRA
    'Atletico Mg':       'Atlético Mineiro',
    'Flamengo Rj':       'Flamengo',
    'Botafogo Rj':       'Botafogo',
    'Ec Juventude':      'Juventude',
    'Bragantino':        'RB Bragantino',
    'Sao Paulo':         'São Paulo',
    'Sport Recife':      'Sport Recife',
    'Ceara':             'Ceará',
    'Gremio':            'Grêmio',
}

def normalizar(nome):
    return NOME_MAP.get(nome, nome)

def _req(url, headers, timeout=12):
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        raw = r.read()
        try:
            return raw.decode('utf-8')
        except:
            return raw.decode('latin-1', errors='replace')

# ── Pesquisa Flashscore ───────────────────────────────────────────────────────
STAT_RE = re.compile(r'SG÷([^¬]+)¬SH÷([^¬]*)¬SI÷([^¬~]*)')

def parse_stat(val):
    if not val: return None
    val = str(val).strip()
    m = re.search(r'\((\d+)/(\d+)\)', val)
    if m: return float(m.group(2))
    try: return float(val.replace('%','').strip())
    except: return None

def similarity(a, b):
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()

def search_flashscore(home, away, kickoff_iso):
    """
    Pesquisa Flashscore por home + away.
    Devolve match_id (str 8 chars) ou None.
    """
    home_fs = normalizar(home)
    away_fs = normalizar(away)

    # Tentar kickoff date para filtrar
    try:
        ko_dt = datetime.fromisoformat(kickoff_iso)
        ko_date = ko_dt.strftime('%Y-%m-%d')
    except:
        ko_date = None

    # Pesquisa 1: home vs away
    for query in [f'{home_fs} {away_fs}', home_fs, away_fs]:
        try:
            url = SEARCH_URL.format(q=urllib.request.quote(query))
            raw = _req(url, SEARCH_HDR)
            mid = _parse_search(raw, home_fs, away_fs, ko_date)
            if mid:
                return mid
            time.sleep(0.5)
        except Exception as e:
            print(f'    [search] Erro ({query}): {e}')

    return None

def _parse_search(raw, home, away, ko_date):
    """
    O Flashscore devolve formato texto: AA÷...¬...
    Procura um jogo que corresponda ao home+away+data.
    """
    # Formato: blocos separados por ~
    # Cada bloco tem: AA÷{id}¬AB÷{status}¬...¬CX÷{home}¬CY÷{away}¬...
    # Alternativa JSON se der
    try:
        data = json.loads(raw)
        if isinstance(data, dict):
            results = data.get('results', []) or data.get('d', {}).get('results', [])
            for r in results:
                h = r.get('homeTeam', {}).get('shortName', '') or r.get('homeParticipant', {}).get('shortName', '')
                a = r.get('awayTeam', {}).get('shortName', '') or r.get('awayParticipant', {}).get('shortName', '')
                mid = r.get('id', '') or r.get('eventId', '')
                if similarity(h, home) > 0.6 and similarity(a, away) > 0.6:
                    return mid[:8] if mid else None
    except:
        pass

    # Formato texto (AA÷ encoding)
    # Exemplo: AA÷AbCd1234¬...¬CX÷Team Home¬CY÷Team Away¬...
    pattern = re.compile(
        r'AA÷([A-Za-z0-9]{8})¬[^~]*?(?:CX|T3)÷([^¬]+)¬(?:CY|T4)÷([^¬~]+)'
    )
    best_score = 0
    best_mid   = None

    for m in pattern.finditer(raw):
        mid_cand = m.group(1)
        h_cand   = m.group(2).strip()
        a_cand   = m.group(3).strip()
        score = similarity(h_cand, home) + similarity(a_cand, away)
        if score > best_score and score > 1.0:
            best_score = score
            best_mid   = mid_cand

    return best_mid

def buscar_lancamentos(mid):
    """
    Vai buscar stats do jogo ao Flashscore ninja API.
    Devolve (total_casa, total_fora) ou (None, None).
    """
    try:
        url  = NINJA_URL.format(mid=mid)
        raw  = _req(url, FS_HDR)
        seen = set()
        lanc_casa = lanc_fora = None
        for m in STAT_RE.finditer(raw):
            name = m.group(1).strip()
            if name == 'Lançamentos' and name not in seen:
                seen.add(name)
                lanc_casa = parse_stat(m.group(2))
                lanc_fora = parse_stat(m.group(3))
                break
        return lanc_casa, lanc_fora
    except Exception as e:
        print(f'    [ninja] Erro {mid}: {e}')
        return None, None

# ── DB ────────────────────────────────────────────────────────────────────────
def get_pending(con, horas_min):
    """Jogos com kickoff passado há > horas_min sem resultado real."""
    limite = (datetime.now(timezone.utc) - timedelta(hours=horas_min)).isoformat()
    rows = con.execute('''
        SELECT id, match_id, league, home, away, kickoff,
               line, modelo_pred, edge_signal
        FROM lancamentos_lines
        WHERE (resultado_real IS NULL OR resultado_real = '')
          AND kickoff < ?
          AND kickoff != ''
        ORDER BY kickoff DESC
    ''', (limite,)).fetchall()
    return rows

def update_resultado(con, row_id, lanc_casa, lanc_fora, flash_mid):
    total = lanc_casa + lanc_fora
    now   = datetime.now(timezone.utc).isoformat()
    # Ler linha e edge_signal para calcular GREEN/RED
    r = con.execute('SELECT line, edge_signal FROM lancamentos_lines WHERE id=?', (row_id,)).fetchone()
    linha = r[0] if r else None
    signal = (r[1] or '').upper() if r else ''
    resultado = _calc_resultado(total, linha, signal)

    con.execute('''UPDATE lancamentos_lines
        SET resultado_real=?, lanc_real_casa=?, lanc_real_fora=?,
            resultado=?, flash_mid=?, updated_at=?
        WHERE id=?
    ''', (total, lanc_casa, lanc_fora, resultado, flash_mid, now, row_id))
    con.commit()
    return total, resultado

def _calc_resultado(total, linha, signal):
    if total is None or linha is None: return None
    if total == linha: return 'Void'
    is_over  = 'OVER'  in signal
    is_under = 'UNDER' in signal
    if not is_over and not is_under: return 'Sem sinal'
    if is_over:  return 'GREEN' if total > linha else 'RED'
    if is_under: return 'GREEN' if total < linha else 'RED'

def garantir_colunas(con):
    """Adiciona colunas novas se não existirem."""
    cols_existentes = {r[1] for r in con.execute('PRAGMA table_info(lancamentos_lines)')}
    novas = [
        ('resultado_real',  'REAL'),
        ('lanc_real_casa',  'REAL'),
        ('lanc_real_fora',  'REAL'),
        ('resultado',       'TEXT'),
        ('flash_mid',       'TEXT'),
    ]
    for col, tipo in novas:
        if col not in cols_existentes:
            con.execute(f'ALTER TABLE lancamentos_lines ADD COLUMN {col} {tipo}')
    con.commit()

# ── Principal ─────────────────────────────────────────────────────────────────
def run():
    if not DB_PATH.exists():
        print(f'DB não encontrada: {DB_PATH}')
        print('Corre primeiro o monitor_lancamentos.py para criar a DB.')
        return

    con = sqlite3.connect(DB_PATH)
    garantir_colunas(con)

    pendentes = get_pending(con, HORAS_MIN)
    print(f'\n{"="*60}')
    print(f'  Colheita de Resultados — {datetime.now().strftime("%d/%m/%Y %H:%M")}')
    print(f'  Jogos pendentes (kickoff > {HORAS_MIN}h atrás): {len(pendentes)}')
    print(f'  Modo: {"DRY RUN" if DRY_RUN else "REAL"}')
    print(f'{"="*60}')

    if not pendentes:
        print('  Nada a processar.\n')
        con.close()
        return

    ok = erro = manual = 0

    for row in pendentes:
        row_id, match_id, league, home, away, kickoff, linha, modelo_pred, edge = row
        try:
            ko_str = datetime.fromisoformat(kickoff).strftime('%d/%m %H:%M')
        except:
            ko_str = kickoff

        print(f'\n  [{league}] {home} vs {away}  ({ko_str})')
        print(f'    Linha: {linha}  |  Modelo: {modelo_pred}  |  Edge: {edge or "—"}')

        # 1. Pesquisar Flashscore
        print(f'    A pesquisar Flashscore...', end=' ')
        flash_mid = search_flashscore(home, away, kickoff)

        if not flash_mid:
            print(f'❌ Match ID não encontrado')
            print(f'    → Podes inserir manualmente: python colheita_resultados.py --mid {match_id} --flash XXXXXXXX')
            manual += 1
            continue

        print(f'OK ({flash_mid})')
        time.sleep(DELAY)

        # 2. Buscar stats
        print(f'    A buscar stats...', end=' ')
        lanc_casa, lanc_fora = buscar_lancamentos(flash_mid)

        if lanc_casa is None or lanc_fora is None:
            print(f'❌ Stats não disponíveis (jogo ainda não terminou?)')
            erro += 1
            continue

        total = lanc_casa + lanc_fora
        resultado_calc = _calc_resultado(total, linha, (edge or '').upper())
        emoji = '✅' if resultado_calc == 'GREEN' else '❌' if resultado_calc == 'RED' else '⚪'
        print(f'OK → {lanc_casa} + {lanc_fora} = {total}  |  Linha: {linha}  →  {emoji} {resultado_calc}')

        if not DRY_RUN:
            update_resultado(con, row_id, lanc_casa, lanc_fora, flash_mid)
            print(f'    ✓ DB atualizada')
        else:
            print(f'    [DRY RUN] Não gravado')

        ok += 1
        time.sleep(DELAY)

    print(f'\n{"─"*60}')
    print(f'  Resumo: {ok} OK | {erro} sem stats | {manual} sem match ID')
    if DRY_RUN:
        print(f'  [DRY RUN] Nenhuma alteração gravada')
    print()
    con.close()

# ── Inserção manual de flash_mid ──────────────────────────────────────────────
def run_manual():
    """
    Uso: python colheita_resultados.py --mid {22bet_CI} --flash {flashscore_mid}
    Vai buscar as stats e actualiza esse jogo específico.
    """
    args = sys.argv
    idx_mid   = args.index('--mid')   if '--mid'   in args else None
    idx_flash = args.index('--flash') if '--flash' in args else None
    if not idx_mid or not idx_flash:
        return False

    match_id  = args[idx_mid + 1]
    flash_mid = args[idx_flash + 1]

    con = sqlite3.connect(DB_PATH)
    garantir_colunas(con)

    row = con.execute('''SELECT id, league, home, away, line, edge_signal
                         FROM lancamentos_lines WHERE match_id=?''', (match_id,)).fetchone()
    if not row:
        print(f'Jogo {match_id} não encontrado na DB.')
        con.close()
        return True

    row_id, league, home, away, linha, edge = row
    print(f'[{league}] {home} vs {away}  |  Linha: {linha}')
    print(f'A buscar stats para {flash_mid}...')

    lanc_casa, lanc_fora = buscar_lancamentos(flash_mid)
    if lanc_casa is None:
        print('Stats não disponíveis.')
        con.close()
        return True

    total, resultado = update_resultado(con, row_id, lanc_casa, lanc_fora, flash_mid)
    emoji = '✅' if resultado == 'GREEN' else '❌' if resultado == 'RED' else '⚪'
    print(f'Total real: {total} ({lanc_casa} + {lanc_fora})  →  {emoji} {resultado}')
    print('DB atualizada.')
    con.close()
    return True

# ── CLI ───────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    if '--mid' in sys.argv and '--flash' in sys.argv:
        run_manual()
    else:
        run()

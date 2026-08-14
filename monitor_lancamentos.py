"""
Monitor de Lançamentos Laterais — 22bet
========================================
Deteta quando a 22bet abre linhas de Lançamentos (TI=55) para jogos
da Primeira Liga, Premier League e Brasileirão Série A.

Guarda linha + odds em SQLite e envia alerta Telegram.
Também compara com o modelo de previsão e sinaliza edge.

Corre em loop (via app.py scheduler) ou manualmente:
    python monitor_lancamentos.py

TI descoberto via DevTools:
  TI=55  → Lançamentos Laterais (Over/Under, G=17)
  Linha  → E[0][*].P  (over)
  Odds   → E[0][*].C (over) / E[1][*].C (under)
  CE=1   → linha central
"""
import os, json, sqlite3, urllib.request, urllib.error, time, sys
from datetime import datetime, timezone
from pathlib import Path

# ── Configuração ──────────────────────────────────────────────────────────────
TELEGRAM_TOKEN   = os.environ.get('TELEGRAM_ALERTS_TOKEN', '')
TELEGRAM_CHAT_ID = os.environ.get('TELEGRAM_ALERTS_CHAT_ID', '')

THROW_IN_TI = 55   # TI do mercado Lançamentos Laterais na 22bet

# Ligas a monitorizar — apenas as que têm mercado de lançamentos
# Para confirmar o ID da La Liga: vai a 22bet → La Liga → inspeciona o URL ou usa --debug
LEAGUES = {
    'PPL': {'name': '🇵🇹 Primeira Liga',   'league_id': int(os.environ.get('LID_PPL', '3007689'))},
    'PL':  {'name': '🏴󠁧󠁢󠁥󠁮󠁧󠁿 Premier League', 'league_id': int(os.environ.get('LID_PL',  '88637'))},
    'BRA': {'name': '🇧🇷 Brasileirão A',   'league_id': int(os.environ.get('LID_BRA', '1268397'))},
    'ESP': {'name': '🇪🇸 La Liga',          'league_id': int(os.environ.get('LID_ESP', '127733'))},
}

# ── 22bet API ─────────────────────────────────────────────────────────────────
_BASE = 'https://22bet4me.com/service-api/LineFeed'
_HDR  = {
    'User-Agent':      ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                        'AppleWebKit/537.36 (KHTML, like Gecko) '
                        'Chrome/120.0.0.0 Safari/537.36'),
    'Accept':          'application/json, text/plain, */*',
    'Accept-Language': 'pt-PT,pt;q=0.9',
    'Referer':         'https://22bet4me.com/line/football/',
}

def _api(url, timeout=15):
    req = urllib.request.Request(url, headers=_HDR)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read())

def get_league_games(league_id):
    """Lista jogos das próximas 48h para a liga."""
    url = (f'{_BASE}/Get1x2_VZip'
           f'?sports=1&champs={league_id}&count=50&lng=pt_PT'
           f'&tf=172800&tz=1&mode=4&country=148&partner=151&getEmpty=true&gr=151')
    try:
        data  = _api(url)
        games = data.get('Value', [])
        return games if isinstance(games, list) else []
    except Exception as e:
        print(f'[22BET] Erro liga {league_id}: {e}')
        return []

def get_game_sg(game_ci):
    """Devolve array SG do jogo (lista de mercados/sub-jogos)."""
    url = (f'{_BASE}/GetGameZip'
           f'?id={game_ci}&lng=pt_PT&tzo=1&cfview=0&isSubGames=true'
           f'&GroupEvents=true&countevents=250&country=148&fcountry=148'
           f'&isNewBuilder=true&partner=151&grMode=4')
    try:
        data = _api(url)
        return data.get('Value', {}).get('SG', [])
    except Exception as e:
        print(f'[22BET] Erro GetGameZip {game_ci}: {e}')
        return []

def find_throw_in_entry(sg):
    """
    Procura Lançamentos (TI=55) no SG.
    Prefere entrada sem P (jogo completo); fallback para qualquer TI=55.
    """
    # 1ª tentativa: jogo completo (sem período)
    entry = next(
        (e for e in sg if e.get('TI') == THROW_IN_TI and not e.get('P')),
        None
    )
    if entry:
        return entry
    # Fallback: qualquer entrada TI=55 (pode ter P=0 ou P='')
    return next((e for e in sg if e.get('TI') == THROW_IN_TI), None)

def get_throw_in_odds(ti_ci):
    """
    Lê as odds Over/Under de Lançamentos para o sub-jogo CI.
    G=17 = grupo Total Over/Under.
    Devolve dict: {line, over, under, all_lines}
    """
    url = (f'{_BASE}/GetGameZip'
           f'?id={ti_ci}&lng=pt&tzo=1&cfview=0&isSubGames=true'
           f'&GroupEvents=true&countevents=250&country=148&fcountry=148'
           f'&isNewBuilder=true&partner=151&grMode=4')
    try:
        data = _api(url)
        ge   = data.get('Value', {}).get('GE', [])
    except Exception as e:
        print(f'[22BET] Erro odds lançamentos {ti_ci}: {e}')
        return {}

    grp = next((g for g in ge if g.get('G') == 17), ge[0] if ge else None)
    if not grp:
        return {}

    over_list  = grp.get('E', [[]])[0] if len(grp.get('E', [])) > 0 else []
    under_list = grp.get('E', [[], []])[1] if len(grp.get('E', [])) > 1 else []
    if not over_list:
        return {}

    mid        = len(over_list) // 2
    main_over  = next((x for x in over_list  if x.get('CE') == 1), over_list[mid])
    main_under = next((x for x in under_list if x.get('CE') == 1),
                      under_list[mid] if mid < len(under_list) else None)

    all_lines = []
    for o, u in zip(over_list[:8], under_list[:8]):
        p = o.get('P')
        if p is not None:
            all_lines.append({
                'line':  p,
                'over':  o.get('C'),
                'under': u.get('C') if u else None,
            })

    return {
        'line':      main_over.get('P'),
        'over':      main_over.get('C'),
        'under':     main_under.get('C') if main_under else None,
        'all_lines': all_lines,
    }

# ── Base de dados ─────────────────────────────────────────────────────────────
_DATA_DIR = os.environ.get('DATA_DIR', str(Path(__file__).parent))
DB_PATH   = Path(_DATA_DIR) / 'lancamentos_lines.db'

def init_db():
    con = sqlite3.connect(DB_PATH)
    con.execute('''CREATE TABLE IF NOT EXISTS lancamentos_lines (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        match_id    TEXT NOT NULL,
        league      TEXT NOT NULL,
        home        TEXT NOT NULL,
        away        TEXT NOT NULL,
        kickoff     TEXT NOT NULL,
        line        REAL,
        over_odds   REAL,
        under_odds  REAL,
        all_lines   TEXT,
        modelo_pred REAL,
        edge_signal TEXT,
        detected_at TEXT NOT NULL,
        updated_at  TEXT
    )''')
    con.execute('''CREATE UNIQUE INDEX IF NOT EXISTS idx_match
                   ON lancamentos_lines(match_id, league)''')
    con.commit()
    return con

def is_known(con, match_id, league):
    r = con.execute(
        'SELECT id, line FROM lancamentos_lines WHERE match_id=? AND league=?',
        (match_id, league)
    ).fetchone()
    return r

def save_line(con, match_id, league, home, away, kickoff_iso,
              line, over_odds, under_odds, all_lines,
              modelo_pred=None, edge_signal=None):
    now = datetime.now(timezone.utc).isoformat()
    con.execute('''INSERT INTO lancamentos_lines
        (match_id, league, home, away, kickoff, line, over_odds, under_odds,
         all_lines, modelo_pred, edge_signal, detected_at)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)''',
        (match_id, league, home, away, kickoff_iso,
         line, over_odds, under_odds, json.dumps(all_lines),
         modelo_pred, edge_signal, now))
    con.commit()

def update_line(con, match_id, league, line, over_odds, under_odds, all_lines, edge_signal=None):
    now = datetime.now(timezone.utc).isoformat()
    con.execute('''UPDATE lancamentos_lines
        SET line=?, over_odds=?, under_odds=?, all_lines=?, edge_signal=?, updated_at=?
        WHERE match_id=? AND league=?''',
        (line, over_odds, under_odds, json.dumps(all_lines), edge_signal, now, match_id, league))
    con.commit()

# ── Modelo de previsão (multi-liga) ─────────────────────────────────────────
_modelos_cache = {}   # liga → pacote

def _get_modelo(liga):
    """Carrega (e cacheia) o modelo para uma liga específica."""
    global _modelos_cache
    if liga in _modelos_cache:
        return _modelos_cache[liga]
    try:
        sys.path.insert(0, str(Path(__file__).parent))
        from modelo_previsao import carregar_pacote
        pacote = carregar_pacote(liga)
        _modelos_cache[liga] = pacote
        return pacote
    except Exception as e:
        print(f'[MODELO] Não disponível [{liga}]: {e}')
        _modelos_cache[liga] = None
        return None

# Compat — carregar_modelo() pré-carrega todos os modelos
def carregar_modelo():
    for liga in ('PL', 'PPL', 'BRA', 'ESP'):
        _get_modelo(liga)
    return _modelos_cache  # devolve dict

def prever_jogo(pacote_ou_dict, home, away, liga='PL'):
    """Tenta prever lançamentos. Devolve (pred, edge_signal) ou (None, None)."""
    pacote = _get_modelo(liga)
    if not pacote:
        return None, None
    try:
        sys.path.insert(0, str(Path(__file__).parent))
        from modelo_previsao import prever_jogo as _prever
        r = _prever(home, away, liga=liga, pacote=pacote)
        if r:
            return r['previsao'], None
        return None, None
    except Exception as e:
        print(f'[MODELO] Erro previsão [{liga}] {home} vs {away}: {e}')
        return None, None

def calcular_edge(pred, linha, std_resid=6.5):
    """Devolve sinal de edge baseado na diferença previsão vs linha."""
    if pred is None or linha is None:
        return None
    diff = pred - linha
    if diff > 4:
        return f'OVER_FORTE (+{diff:.1f})'
    elif diff > 2:
        return f'OVER_MODERADO (+{diff:.1f})'
    elif diff < -4:
        return f'UNDER_FORTE ({diff:.1f})'
    elif diff < -2:
        return f'UNDER_MODERADO ({diff:.1f})'
    return None

# ── Telegram ──────────────────────────────────────────────────────────────────
def send_telegram(text):
    if not TELEGRAM_TOKEN or not TELEGRAM_CHAT_ID:
        print(f'[ALERT-DRY] {text}')
        return
    try:
        body = json.dumps({'chat_id': TELEGRAM_CHAT_ID, 'text': text,
                           'parse_mode': 'HTML'}).encode()
        req = urllib.request.Request(
            f'https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage',
            data=body,
            headers={'Content-Type': 'application/json'},
        )
        urllib.request.urlopen(req, timeout=8)
    except Exception as e:
        print(f'[TELEGRAM] Erro: {e}')

def build_alert(league_name, home, away, kickoff_iso, odds_info,
                modelo_pred=None, edge_signal=None):
    """Formata mensagem de alerta."""
    try:
        ko_dt = datetime.fromisoformat(kickoff_iso)
        ko_str = ko_dt.strftime('%d/%m %H:%M')
    except:
        ko_str = kickoff_iso

    linha = odds_info.get('line')
    over  = odds_info.get('over')
    under = odds_info.get('under')

    lines = [
        f"🎯 <b>LINHA ABERTA · Lançamentos</b>  ·  {league_name}",
        f"⚽ {home}  vs  {away}",
        f"🕐 {ko_str}",
        f"📊 Linha: <b>{linha}</b>  |  Over: {over}  |  Under: {under}",
    ]
    if modelo_pred:
        lines.append(f"🤖 Modelo prevê: <b>{modelo_pred:.1f}</b>")
    if edge_signal:
        emoji = '🟢' if 'FORTE' in edge_signal else '🟡'
        lines.append(f"{emoji} <b>EDGE: {edge_signal}</b>")

    return '\n'.join(lines)

# ── Ciclo principal ────────────────────────────────────────────────────────────
def run_once(verbose=True):
    con = init_db()
    carregar_modelo()   # pré-carrega todos os modelos no cache
    now = datetime.now(timezone.utc)
    alerts = 0

    for code, cfg in LEAGUES.items():
        league_name = cfg['name']
        league_id   = cfg['league_id']

        if verbose:
            print(f'\n[{code}] A verificar {league_name} (LID={league_id})...')

        games = get_league_games(league_id)
        if verbose:
            print(f'  → {len(games)} jogos encontrados')

        for g in games:
            ci   = g.get('CI')
            home = g.get('O1E') or g.get('O1') or '?'
            away = g.get('O2E') or g.get('O2') or '?'
            ko_ts = g.get('S')

            if not ci:
                continue

            # Kickoff
            try:
                ko_dt  = datetime.fromtimestamp(ko_ts, tz=timezone.utc) if ko_ts else None
                ko_iso = ko_dt.isoformat() if ko_dt else ''
                ko_str = ko_dt.strftime('%d/%m %H:%M') if ko_dt else '?'
            except:
                ko_iso = ko_str = str(ko_ts or '')

            if verbose:
                print(f'  Jogo: {home} vs {away}  ({ko_str})')

            # Ver mercados do jogo
            sg    = get_game_sg(ci)
            entry = find_throw_in_entry(sg)

            if not entry:
                if verbose:
                    tis_encontrados = sorted({e.get('TI') for e in sg if e.get('TI')})
                    print(f'    → Sem lançamentos (TI=55). TIs disponíveis: {tis_encontrados}')
                continue

            ti_ci = entry.get('CI')
            if not ti_ci:
                continue

            # Ler odds
            odds = get_throw_in_odds(ti_ci)
            if not odds or odds.get('line') is None:
                if verbose:
                    print(f'    → Linha vazia')
                continue

            current_line = odds['line']
            if verbose:
                print(f'    ✓ Linha: {current_line}  Over: {odds["over"]}  Under: {odds["under"]}')

            # Previsão do modelo (usa o modelo da liga correcta)
            modelo_pred, _ = prever_jogo(None, home, away, liga=code)
            std_resid = (_modelos_cache.get(code) or {}).get('std_resid', 6.5)
            edge_signal = calcular_edge(modelo_pred, current_line, std_resid)
            if verbose and modelo_pred:
                print(f'    🤖 Modelo: {modelo_pred:.1f}  Edge: {edge_signal or "—"}')

            # Verificar se já está na DB
            existing = is_known(con, str(ci), code)

            if not existing:
                # Nova linha — guardar e alertar
                save_line(con, str(ci), code, home, away, ko_iso,
                          current_line, odds['over'], odds['under'],
                          odds.get('all_lines', []), modelo_pred, edge_signal)

                msg = build_alert(league_name, home, away, ko_iso, odds,
                                  modelo_pred, edge_signal)
                send_telegram(msg)
                print(f'    [NOVO] Alerta enviado!')
                alerts += 1

            elif existing[1] != current_line:
                # Linha mudou — atualizar
                old_line = existing[1]
                update_line(con, str(ci), code, current_line,
                            odds['over'], odds['under'],
                            odds.get('all_lines', []), edge_signal)
                if verbose:
                    print(f'    [UPDATE] Linha: {old_line} → {current_line}')
            else:
                if verbose:
                    print(f'    [KNOWN] Sem alteração')

            time.sleep(0.5)

    con.close()
    return alerts

def run_loop(interval_minutes=15):
    print(f'Monitor Lançamentos a correr — intervalo {interval_minutes}min')
    print(f'Ligas: {", ".join(LEAGUES.keys())}  |  TI=55')
    print(f'DB: {DB_PATH}\n')
    while True:
        try:
            alerts = run_once()
            print(f'\n[{datetime.now().strftime("%H:%M")}] Ciclo completo. {alerts} alertas.')
        except Exception as e:
            print(f'[ERRO] {e}')
        time.sleep(interval_minutes * 60)

# ── CLI ───────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    if '--loop' in sys.argv:
        mins = int(next((sys.argv[i+1] for i,a in enumerate(sys.argv)
                         if a == '--loop' and i+1 < len(sys.argv)), 15))
        run_loop(mins)
    elif '--debug' in sys.argv:
        # Modo debug: mostra resposta bruta da API para a primeira liga
        code = 'PPL'
        lid  = LEAGUES[code]['league_id']
        print(f'[DEBUG] Liga: {code}  LID={lid}')
        games = get_league_games(lid)
        print(f'[DEBUG] Jogos encontrados: {len(games)}')
        if not games:
            print('[DEBUG] API devolveu lista vazia.')
            sys.exit()
        # Inspecionar primeiro jogo em detalhe
        g    = games[0]
        ci   = g.get('CI')
        home = g.get('O1E') or g.get('O1')
        away = g.get('O2E') or g.get('O2')
        print(f'\n[DEBUG] {home} vs {away} (CI={ci})')
        sg = get_game_sg(ci)
        print(f'[DEBUG] SG entries: {len(sg)}')
        # Mostrar TODOS os SG entries com TI
        for e in sg:
            ti = e.get('TI')
            if ti:
                e_ci = e.get('CI')
                print(f'\n  TI={ti}  CI={e_ci}  P={e.get("P","")!r}  N={e.get("N","")!r}')
                # Ir buscar o nome do mercado (GE)
                url = (f'{_BASE}/GetGameZip'
                       f'?id={e_ci}&lng=pt&tzo=1&cfview=0&isSubGames=true'
                       f'&GroupEvents=true&countevents=250&country=148&fcountry=148'
                       f'&isNewBuilder=true&partner=151&grMode=4')
                try:
                    import urllib.request as _ur
                    req = _ur.Request(url, headers=_HDR)
                    with _ur.urlopen(req, timeout=10) as r:
                        d = json.loads(r.read())
                    ge = d.get('Value', {}).get('GE', [])
                    print(f'  GE groups: {len(ge)}')
                    for grp in ge[:5]:
                        g_id = grp.get('G')
                        g_n  = grp.get('GN', '')
                        evts = len(grp.get('E', [[]])[0]) if grp.get('E') else 0
                        print(f'    G={g_id}  GN={g_n!r}  events={evts}')
                        # Mostrar primeira linha/odd
                        ev = (grp.get('E') or [[]])[0]
                        if ev:
                            print(f'    Exemplo: P={ev[0].get("P")}  C={ev[0].get("C")}  CE={ev[0].get("CE")}')
                except Exception as ex:
                    print(f'  Erro: {ex}')
    else:
        print('=== Monitor Lançamentos — execução única ===')
        try:
            alerts = run_once(verbose=True)
            print(f'\nTotal alertas: {alerts}')
            print(f'DB: {DB_PATH}')
        except Exception as e:
            import traceback
            print(f'\n[ERRO FATAL] {e}')
            traceback.print_exc()

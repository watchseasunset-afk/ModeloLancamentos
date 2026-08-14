"""
Live Collector — Lançamentos + Faltas
======================================
Recolhe stats em tempo real durante os jogos do fim de semana.
Guarda snapshots de 10 em 10 minutos com previsão intercalar.

Uso:
    python live_collector.py --setup       # encontra jogos do fds e regista
    python live_collector.py --collect     # corre ciclo de colecta (1 vez)
    python live_collector.py --loop        # corre em loop até não haver jogos live
    python live_collector.py --analyze     # análise pós-jogo
    python live_collector.py --status      # mostra estado dos jogos
"""
import os, json, re, sqlite3, sys, time, urllib.request, urllib.error
from datetime import datetime, timezone, timedelta
from pathlib import Path
from difflib import SequenceMatcher

# ── Configuração ──────────────────────────────────────────────────────────────
BASE     = Path(__file__).parent
DB_PATH  = BASE / 'live_stats.db'
INTERVAL = 600   # segundos entre ciclos (10 min)

# Médias históricas por liga (faltas/jogo e lançamentos/jogo)
# La Liga 2023/24: ~23 faltas, ~40 lançamentos (fonte: FBref/Understat)
FALTAS_BASE = {'PPL': 27.5, 'PL': 21.6, 'BRA': 27.7, 'ESP': 23.0}
LANC_BASE   = {'PPL': 39.1, 'PL': 35.8, 'BRA': 36.7, 'ESP': 40.5}

# Ligas activas
# LID_ESP: confirmar via 22bet → La Liga → inspecionar URL ou usar --find-league
LEAGUES = {
    'PPL': {'name': 'Primeira Liga',  'league_id': int(os.environ.get('LID_PPL', '3007689'))},
    'BRA': {'name': 'Brasileirão A',  'league_id': int(os.environ.get('LID_BRA', '1268397'))},
    'ESP': {'name': 'La Liga',        'league_id': int(os.environ.get('LID_ESP', '127733'))},
}

# ── Flashscore ────────────────────────────────────────────────────────────────
FS_HDR = {
    'X-Fsign':    'SW9D1eZo',
    'Referer':    'https://www.flashscore.com/',
    'Accept':     '*/*',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
}
SEARCH_HDR = {
    'Referer':          'https://www.flashscore.com/',
    'Accept':           'application/json, */*',
    'User-Agent':       'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'x-requested-with': 'XMLHttpRequest',
}
NINJA_URL  = 'https://global.flashscore.ninja/20/x/feed/df_st_1_{mid}'
SEARCH_URL = 'https://s.flashscore.com/search/?q={q}&l=1&s=2&f=1%3B1&pid=2&sid=1'
STAT_RE    = re.compile(r'SG÷([^¬]+)¬SH÷([^¬]*)¬SI÷([^¬~]*)')

# Mapeamento nomes 22bet → Flashscore
NOME_MAP = {
    # PPL
    'Fc Porto': 'Porto', 'Sporting Cp': 'Sporting CP',
    'Estrela Da Amadora': 'Estrela Amadora',
    'Famalicao': 'Famalicão', 'Moreirense': 'Moreirense FC',
    # BRA
    'Atletico Mg': 'Atlético Mineiro', 'Flamengo Rj': 'Flamengo',
    'Botafogo Rj': 'Botafogo', 'Ec Juventude': 'Juventude',
    'Bragantino': 'RB Bragantino', 'Sao Paulo': 'São Paulo',
    'Gremio': 'Grêmio', 'Ceara': 'Ceará',
    # La Liga
    'Atletico Madrid': 'Atlético Madrid',
    'Atletico De Madrid': 'Atlético Madrid',
    'Real Betis': 'Real Betis',
    'Rayo Vallecano': 'Rayo Vallecano',
    'Real Sociedad': 'Real Sociedad',
    'Athletic Bilbao': 'Athletic Club',
    'Athletic Club Bilbao': 'Athletic Club',
    'Deportivo Alaves': 'Alavés',
    'Alaves': 'Alavés',
    'Cd Leganes': 'Leganés',
    'Leganes': 'Leganés',
    'Rc Celta': 'Celta Vigo',
    'Celta De Vigo': 'Celta Vigo',
    'Espanyol': 'Espanyol',
    'Rcd Espanyol': 'Espanyol',
    'Cd Espanyol': 'Espanyol',
    'Getafe Cf': 'Getafe',
    'Real Valladolid': 'Valladolid',
    'Sevilla Fc': 'Sevilla',
    'Valencia Cf': 'Valencia',
    'Villarreal Cf': 'Villarreal',
    'Osasuna': 'Osasuna',
    'Ca Osasuna': 'Osasuna',
    # La Liga 2026/27 — equipas promovidas
    'Levante': 'Levante',
    'Levante Ud': 'Levante',
    'Racing Santander': 'Racing Santander',
    'Racing De Santander': 'Racing Santander',
    'Deportivo La Coruna': 'Dep. A Coruna',
    'Rc Deportivo': 'Dep. A Coruna',
    'Deportivo De La Coruna': 'Dep. A Coruna',
    'Elche': 'Elche',
    'Elche Cf': 'Elche',
    'Malaga': 'Málaga',
    'Malaga Cf': 'Málaga',
    'Cd Malaga': 'Málaga',
}

# ── 22bet API ─────────────────────────────────────────────────────────────────
_BASE_22 = 'https://22bet4me.com/service-api/LineFeed'
_HDR_22  = {
    'User-Agent':      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept':          'application/json, text/plain, */*',
    'Accept-Language': 'pt-PT,pt;q=0.9',
    'Referer':         'https://22bet4me.com/line/football/',
}

def _req(url, headers, timeout=12):
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        raw = r.read()
        try:    return raw.decode('utf-8')
        except: return raw.decode('latin-1', errors='replace')

def _api22(url, timeout=15):
    req = urllib.request.Request(url, headers=_HDR_22)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read())

def get_league_games(league_id):
    url = (f'{_BASE_22}/Get1x2_VZip'
           f'?sports=1&champs={league_id}&count=50&lng=pt_PT'
           f'&tf=259200&tz=1&mode=4&country=148&partner=151&getEmpty=true&gr=151')
    try:
        data = _api22(url)
        return data.get('Value', []) or []
    except Exception as e:
        print(f'[22BET] Erro {league_id}: {e}')
        return []

# ── Flashscore helpers ────────────────────────────────────────────────────────
def similarity(a, b):
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()

def normalizar(nome):
    return NOME_MAP.get(nome, nome)

def search_flashscore(home, away):
    home_fs = normalizar(home)
    away_fs = normalizar(away)
    for query in [f'{home_fs} {away_fs}', home_fs, away_fs]:
        try:
            url = SEARCH_URL.format(q=urllib.request.quote(query))
            raw = _req(url, SEARCH_HDR)
            mid = _parse_search(raw, home_fs, away_fs)
            if mid:
                return mid
            time.sleep(0.5)
        except Exception as e:
            print(f'  [search] Erro: {e}')
    return None

def _parse_search(raw, home, away):
    try:
        data = json.loads(raw)
        results = data.get('results', [])
        for r in results:
            h = r.get('homeTeam', {}).get('shortName', '')
            a = r.get('awayTeam', {}).get('shortName', '')
            mid = r.get('id', '')
            if similarity(h, home) > 0.6 and similarity(a, away) > 0.6:
                return mid[:8] if mid else None
    except:
        pass
    pattern = re.compile(r'AA÷([A-Za-z0-9]{8})¬[^~]*?(?:CX|T3)÷([^¬]+)¬(?:CY|T4)÷([^¬~]+)')
    best_score, best_mid = 0, None
    for m in pattern.finditer(raw):
        h_cand = m.group(2).strip()
        a_cand = m.group(3).strip()
        score = similarity(h_cand, home) + similarity(a_cand, away)
        if score > best_score and score > 1.0:
            best_score = score
            best_mid   = m.group(1)
    return best_mid

def fetch_live_stats(flash_mid):
    """
    Busca stats ao vivo do Flashscore ninja API.
    Devolve dict com lanc_casa, lanc_fora, faltas_casa, faltas_fora, minuto_est.
    """
    try:
        url = NINJA_URL.format(mid=flash_mid)
        raw = _req(url, FS_HDR)
        stats = {}
        for m in STAT_RE.finditer(raw):
            name = m.group(1).strip()
            val_h = _parse_stat(m.group(2))
            val_a = _parse_stat(m.group(3))
            if name == 'Lançamentos':
                stats['lanc_casa'] = val_h
                stats['lanc_fora'] = val_a
            elif name == 'Faltas':
                stats['faltas_casa'] = val_h
                stats['faltas_fora'] = val_a
        return stats
    except Exception as e:
        print(f'  [ninja] Erro {flash_mid}: {e}')
        return {}

def _parse_stat(val):
    if not val: return None
    val = str(val).strip()
    m = re.search(r'\((\d+)/(\d+)\)', val)
    if m: return float(m.group(2))
    try: return float(val.replace('%','').strip())
    except: return None

# ── Modelo pré-jogo ───────────────────────────────────────────────────────────
_modelos_cache = {}

def get_modelo_pred(liga, home, away):
    """Previsão pré-jogo de lançamentos usando modelo treinado."""
    try:
        import pickle
        if liga not in _modelos_cache:
            f = BASE / 'data' / f'modelo_lancamentos_{liga}.pkl'
            if not f.exists():
                f = BASE / 'data' / 'modelo_lancamentos.pkl'
            with open(f, 'rb') as fh:
                _modelos_cache[liga] = pickle.load(fh)
        pkg = _modelos_cache[liga]
        model   = pkg['model']
        scaler  = pkg['scaler']
        feats   = pkg['feat_cols']
        medias  = pkg.get('medias', {})
        medias_g= pkg.get('medias_global', {})
        hf      = pkg.get('home_feats', [])
        af      = pkg.get('away_feats', [])

        h_med = medias.get(home, {})
        a_med = medias.get(away, {})

        row = {}
        for f_name in feats:
            orig = f_name[2:]  # remove H_ ou A_
            if f_name.startswith('H_'):
                row[f_name] = h_med.get(orig, medias_g.get('total', 37) / 2)
            else:
                row[f_name] = a_med.get(orig, medias_g.get('total', 37) / 2)

        import numpy as np
        X = np.array([[row[f] for f in feats]])
        Xs = scaler.transform(X)
        pred = float(model.predict(Xs)[0])
        return round(pred, 1)
    except Exception as e:
        print(f'  [modelo] Erro {liga} {home} vs {away}: {e}')
        return LANC_BASE.get(liga, 37.0)

# ── Previsão intercalar ───────────────────────────────────────────────────────
def calcular_previsao(minuto, total_atual, baseline, max_min=90):
    """
    Previsão ponderada: combina extrapolação live + baseline pré-jogo.
    Quanto mais avançado o jogo, mais peso tem a extrapolação.
    """
    if minuto <= 0:
        return baseline
    peso_live = min(minuto / max_min, 1.0)
    extrap = (total_atual / minuto) * max_min
    return round(peso_live * extrap + (1 - peso_live) * baseline, 1)

# ── Base de dados ─────────────────────────────────────────────────────────────
def init_db():
    con = sqlite3.connect(DB_PATH)
    con.execute('''CREATE TABLE IF NOT EXISTS live_games (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        flash_mid       TEXT UNIQUE,
        league          TEXT,
        home            TEXT,
        away            TEXT,
        kickoff         TEXT,
        lanc_baseline   REAL,
        faltas_baseline REAL,
        status          TEXT DEFAULT 'pending',
        final_lanc      INTEGER,
        final_faltas    INTEGER,
        added_at        TEXT
    )''')
    con.execute('''CREATE TABLE IF NOT EXISTS live_snapshots (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        flash_mid       TEXT,
        league          TEXT,
        home            TEXT,
        away            TEXT,
        kickoff         TEXT,
        minuto_est      INTEGER,
        lanc_casa       INTEGER,
        lanc_fora       INTEGER,
        lanc_total      INTEGER,
        faltas_casa     INTEGER,
        faltas_fora     INTEGER,
        faltas_total    INTEGER,
        lanc_extrap     REAL,
        faltas_extrap   REAL,
        lanc_baseline   REAL,
        faltas_baseline REAL,
        lanc_pred       REAL,
        faltas_pred     REAL,
        captured_at     TEXT
    )''')
    con.commit()
    return con

# ── Setup: encontrar jogos do fds ─────────────────────────────────────────────
def setup_games():
    """Encontra jogos para as próximas 72h e regista na DB."""
    con = init_db()
    now = datetime.now(timezone.utc)
    total_added = 0

    for liga_key, liga_info in LEAGUES.items():
        print(f'\n[SETUP] {liga_info["name"]} ...')
        games = get_league_games(liga_info['league_id'])
        print(f'  {len(games)} jogos encontrados na 22bet')

        for g in games:
            home = g.get('O1', '')
            away = g.get('O2', '')
            ts   = g.get('S', 0)
            if not home or not away or not ts:
                continue

            ko = datetime.fromtimestamp(ts, tz=timezone.utc)
            ko_iso = ko.isoformat()

            # Só jogos futuros ou das próximas 72h
            if ko < now - timedelta(hours=2):
                continue

            print(f'  → {home} vs {away} ({ko.strftime("%d/%m %H:%M")})')

            # Procurar flash_mid no Flashscore
            flash_mid = search_flashscore(home, away)
            if not flash_mid:
                print(f'    ⚠ flash_mid não encontrado')
                continue

            print(f'    flash_mid: {flash_mid}')

            # Previsão pré-jogo
            lanc_base   = get_modelo_pred(liga_key, home, away)
            faltas_base = FALTAS_BASE.get(liga_key, 27.0)

            # Guardar na DB
            try:
                con.execute('''INSERT OR IGNORE INTO live_games
                    (flash_mid, league, home, away, kickoff,
                     lanc_baseline, faltas_baseline, status, added_at)
                    VALUES (?,?,?,?,?,?,?,'pending',?)''',
                    (flash_mid, liga_key, home, away, ko_iso,
                     lanc_base, faltas_base, now.isoformat()))
                con.commit()
                total_added += 1
                print(f'    ✓ Registado (lanc_base={lanc_base}, faltas_base={faltas_base})')
            except Exception as e:
                print(f'    ⚠ DB erro: {e}')

            time.sleep(1)

    con.close()
    print(f'\n[SETUP] Total registado: {total_added} jogos')

# ── Colecta: um ciclo de snapshots ───────────────────────────────────────────
def collect_once():
    """Corre um ciclo: para cada jogo live, recolhe stats e guarda snapshot."""
    con = init_db()
    now = datetime.now(timezone.utc)

    # Jogos que estão live agora (kickoff passado há < 115 min)
    live = con.execute('''
        SELECT flash_mid, league, home, away, kickoff,
               lanc_baseline, faltas_baseline
        FROM live_games
        WHERE status IN ('pending','live')
          AND kickoff <= ?
          AND kickoff >= ?
    ''', (now.isoformat(),
          (now - timedelta(minutes=115)).isoformat())).fetchall()

    if not live:
        print('[COLLECT] Sem jogos live neste momento.')
        con.close()
        return 0

    snapshots = 0
    for row in live:
        flash_mid, liga, home, away, ko_iso, lanc_base, faltas_base = row

        # Calcular minuto estimado
        try:
            ko_dt  = datetime.fromisoformat(ko_iso)
            minuto = int((now - ko_dt).total_seconds() / 60)
            minuto = max(0, min(minuto, 95))
        except:
            minuto = 45

        print(f'[COLLECT] {home} vs {away} (~{minuto}\') ...')

        # Marcar como live
        con.execute("UPDATE live_games SET status='live' WHERE flash_mid=?", (flash_mid,))
        con.commit()

        # Buscar stats
        stats = fetch_live_stats(flash_mid)
        if not stats:
            print(f'  ⚠ Sem stats disponíveis')
            continue

        lanc_casa   = stats.get('lanc_casa')
        lanc_fora   = stats.get('lanc_fora')
        faltas_casa = stats.get('faltas_casa')
        faltas_fora = stats.get('faltas_fora')

        lanc_total   = (lanc_casa or 0) + (lanc_fora or 0) if lanc_casa is not None else None
        faltas_total = (faltas_casa or 0) + (faltas_fora or 0) if faltas_casa is not None else None

        # Previsões intercalares
        lanc_pred   = calcular_previsao(minuto, lanc_total or 0, lanc_base or LANC_BASE.get(liga, 37))
        faltas_pred = calcular_previsao(minuto, faltas_total or 0, faltas_base or FALTAS_BASE.get(liga, 27))

        lanc_extrap   = round((lanc_total / minuto * 90), 1) if (lanc_total and minuto > 0) else None
        faltas_extrap = round((faltas_total / minuto * 90), 1) if (faltas_total and minuto > 0) else None

        print(f'  Lançamentos: {lanc_casa}+{lanc_fora}={lanc_total} → pred={lanc_pred}')
        print(f'  Faltas:      {faltas_casa}+{faltas_fora}={faltas_total} → pred={faltas_pred}')

        # Guardar snapshot
        con.execute('''INSERT INTO live_snapshots
            (flash_mid, league, home, away, kickoff, minuto_est,
             lanc_casa, lanc_fora, lanc_total,
             faltas_casa, faltas_fora, faltas_total,
             lanc_extrap, faltas_extrap,
             lanc_baseline, faltas_baseline,
             lanc_pred, faltas_pred, captured_at)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (flash_mid, liga, home, away, ko_iso, minuto,
             lanc_casa, lanc_fora, lanc_total,
             faltas_casa, faltas_fora, faltas_total,
             lanc_extrap, faltas_extrap,
             lanc_base, faltas_base,
             lanc_pred, faltas_pred,
             now.isoformat()))
        con.commit()
        snapshots += 1
        time.sleep(1.5)

    # Marcar jogos terminados (> 115 min após kickoff)
    terminados = con.execute('''
        UPDATE live_games SET status='done'
        WHERE status='live' AND kickoff <= ?
    ''', ((now - timedelta(minutes=115)).isoformat(),)).rowcount
    con.commit()

    if terminados:
        # Para os jogos terminados, guardar stats finais
        _finalizar_jogos(con)

    con.close()
    return snapshots

def _finalizar_jogos(con):
    """Para jogos terminados, guardar o total final do último snapshot."""
    done = con.execute('''
        SELECT DISTINCT flash_mid, league, home, away
        FROM live_games WHERE status='done' AND final_lanc IS NULL
    ''').fetchall()
    for flash_mid, liga, home, away in done:
        last = con.execute('''
            SELECT lanc_total, faltas_total FROM live_snapshots
            WHERE flash_mid=? ORDER BY minuto_est DESC LIMIT 1
        ''', (flash_mid,)).fetchone()
        if last:
            con.execute('''UPDATE live_games SET final_lanc=?, final_faltas=?
                WHERE flash_mid=?''', (last[0], last[1], flash_mid))
            con.commit()
            print(f'[DONE] {home} vs {away}: lanc={last[0]}, faltas={last[1]}')

# ── Loop principal ────────────────────────────────────────────────────────────
def run_loop():
    """Corre em loop de 10 em 10 minutos até não haver jogos activos."""
    print('[LOOP] A iniciar colecta live...')
    idle_cycles = 0
    while True:
        n = collect_once()
        if n == 0:
            idle_cycles += 1
            print(f'[LOOP] Sem actividade ({idle_cycles}x). A aguardar {INTERVAL//60} min...')
            if idle_cycles >= 6:  # 1h sem jogos → parar
                print('[LOOP] Sem jogos há 1h. A terminar.')
                break
        else:
            idle_cycles = 0
        time.sleep(INTERVAL)

# ── Análise pós-jogo ──────────────────────────────────────────────────────────
def analyze():
    """Analisa precisão das previsões intercalares vs resultado final."""
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row

    games = con.execute('''
        SELECT * FROM live_games WHERE status='done' AND final_lanc IS NOT NULL
    ''').fetchall()

    if not games:
        print('[ANALYZE] Sem jogos terminados ainda.')
        con.close()
        return

    print(f'\n{"="*60}')
    print(f'  ANÁLISE LIVE — {len(games)} jogos')
    print(f'{"="*60}')

    for g in games:
        snaps = con.execute('''
            SELECT * FROM live_snapshots WHERE flash_mid=? ORDER BY minuto_est
        ''', (g['flash_mid'],)).fetchall()

        if not snaps:
            continue

        final_lanc   = g['final_lanc']
        final_faltas = g['final_faltas']

        print(f'\n{g["home"]} vs {g["away"]} ({g["league"]})')
        print(f'  Final: {final_lanc} lançamentos | {final_faltas} faltas')
        print(f'  Baseline pré-jogo: lanc={g["lanc_baseline"]} | faltas={g["faltas_baseline"]}')
        print(f'\n  {"Min":>4} | {"Lanc":>6} | {"Pred":>6} | {"Erro":>6} | {"Faltas":>6} | {"Pred":>6} | {"Erro":>6}')
        print(f'  {"-"*55}')

        for s in snaps:
            err_l = round(s['lanc_pred'] - final_lanc, 1) if s['lanc_pred'] and final_lanc else '?'
            err_f = round(s['faltas_pred'] - final_faltas, 1) if s['faltas_pred'] and final_faltas else '?'
            print(f'  {s["minuto_est"]:>4} | {s["lanc_total"] or "?":>6} | {s["lanc_pred"] or "?":>6} | {err_l:>6} | {s["faltas_total"] or "?":>6} | {s["faltas_pred"] or "?":>6} | {err_f:>6}')

    con.close()

# ── Status ────────────────────────────────────────────────────────────────────
def status():
    if not DB_PATH.exists():
        print('DB não existe. Corre --setup primeiro.')
        return
    con = sqlite3.connect(DB_PATH)
    games = con.execute('SELECT * FROM live_games ORDER BY kickoff').fetchall()
    snaps = con.execute('SELECT COUNT(*) FROM live_snapshots').fetchone()[0]
    print(f'\n{"="*55}')
    print(f'  Live Collector — Status')
    print(f'{"="*55}')
    print(f'  Snapshots totais: {snaps}')
    print(f'\n  {"Liga":>5} | {"Casa":>20} | {"Fora":>20} | {"KO":>12} | {"Status":>8} | {"Lanc Final":>10}')
    for g in games:
        ko = g[4][:16] if g[4] else '?'
        print(f'  {g[2]:>5} | {g[3]:>20} | {g[4]:>20} | {ko:>12} | {g[8]:>8} | {g[9] or "—":>10}')
    con.close()

# ── CLI ───────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    if '--setup' in sys.argv:
        setup_games()
    elif '--collect' in sys.argv:
        collect_once()
    elif '--loop' in sys.argv:
        run_loop()
    elif '--analyze' in sys.argv:
        analyze()
    elif '--status' in sys.argv:
        status()
    else:
        print(__doc__)

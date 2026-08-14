"""
Lançamentos Laterais — App Local
==================================
Corre em http://localhost:5000
  - Dashboard com todos os jogos da DB
  - Monitor corre automaticamente de 2 em 2 horas
  - Botão para correr manualmente
  - Logs em tempo real

Iniciar:
    python app.py
"""
import json, sqlite3, threading, logging, sys, urllib.request, urllib.parse
from datetime import datetime, timezone, timedelta
from pathlib import Path
from flask import Flask, render_template_string, jsonify, Response, send_file, request

# ── APScheduler ───────────────────────────────────────────────────────────────
from apscheduler.schedulers.background import BackgroundScheduler

# ── Paths ─────────────────────────────────────────────────────────────────────
import os
BASE     = Path(__file__).parent
DATA_DIR = Path(os.environ.get('DATA_DIR', str(BASE)))
DB_PATH  = DATA_DIR / 'lancamentos_lines.db'
DATA_DIR.mkdir(parents=True, exist_ok=True)

LIVE_DB_PATH        = DATA_DIR / 'live_stats.db'
LIVE_TOKEN          = os.environ.get('RAILWAY_TOKEN', 'live2026')
FALTAS_MODEL_PATH   = Path(os.environ.get(
    'FALTAS_MODEL_PATH',
    r'C:\Claude_Mod_Faltas\Faltas_206_2027\faltas_app_v5\faltas_app\model_data.json'
))
FALTAS_CACHE_PATH   = DATA_DIR / 'faltas_model_cache.json'

# Liga name mapping: model_data.json keys → códigos internos
_LIGA_MAP = {'Tugao': 'PPL', 'Brazil': 'BRA', 'Spain': 'ESP'}

# ── Médias equipas (carregado do Excel no arranque) ───────────────────────────
import unicodedata, re as _re

def _norm_nome(s):
    """Normaliza nome de equipa para matching: minúsculas, sem acentos, sem espaços duplos."""
    if not s: return ''
    s = unicodedata.normalize('NFD', str(s))
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    return _re.sub(r'\s+', ' ', s.lower().strip())

_MEDIAS:    dict = {}   # {norm_nome: {liga, equipa, faltas_media, jogos, lanc_media, posse_media}}
_ARBITROS:  dict = {}   # {'PPL': {norm_nome: {nome, jogos, media_faltas}}, 'ESP': {...}, 'BRA': {...}}

# Ligas disponíveis em model_data.json → código interno
_JSON_LIGA_MAP = {'Tugao': 'PPL', 'Spain': 'ESP', 'Brazil': 'BRA', 'UK': 'UK', 'Italia': 'ITA', 'France': 'FRA'}

def _load_model_json():
    """Tenta carregar model_data.json do projecto Faltas (master).
    Prioridade: Faltas project (master) → bundled copy → DATA_DIR.
    Em Railway só existe a bundled copy."""
    paths = [
        FALTAS_MODEL_PATH,          # master — projecto Faltas local
        BASE / 'model_data.json',   # bundled — incluído no deploy Railway
        DATA_DIR / 'model_data.json',
    ]
    for p in paths:
        try:
            if Path(p).exists():
                with open(p, encoding='utf-8') as f:
                    data = json.load(f)
                print(f'[MODEL_DATA] carregado de {p}')
                return data
        except Exception as e:
            print(f'[MODEL_DATA] erro em {p}: {e}')
    print('[MODEL_DATA] model_data.json nao encontrado em nenhum caminho')
    return None

def _load_medias():
    """Carrega médias de faltas de todas as equipas a partir de model_data.json (projecto Faltas)."""
    global _MEDIAS
    md = _load_model_json()
    if not md:
        print('[MEDIAS] sem model_data.json — medias de faltas indisponíveis')
        return
    tmp = {}
    for liga_key, liga_code in _JSON_LIGA_MAP.items():
        teams = md.get('teams', {}).get(liga_key, [])
        for t in teams:
            nome = t.get('name', '')
            if not nome:
                continue
            gh  = t.get('games_home', 0) or 0
            ga  = t.get('games_away', 0) or 0
            total = gh + ga
            ffh = t.get('ffh', 0) or 0   # faltas feitas em casa
            ffa = t.get('ffa', 0) or 0   # faltas feitas fora
            faltas_media = round((ffh * gh + ffa * ga) / total, 2) if total > 0 else 0.0
            key = _norm_nome(nome)
            tmp[key] = {
                'liga':         liga_code,
                'equipa':       nome,
                'jogos':        total,
                'faltas_media': faltas_media,
                'ffh':          round(ffh, 2),
                'fsh':          round(float(t.get('fsh', 0) or 0), 2),
                'ffa':          round(ffa, 2),
                'fsa':          round(float(t.get('fsa', 0) or 0), 2),
                'lanc_media':   0.0,   # não disponível em model_data.json
                'posse_media':  0.0,   # não disponível em model_data.json
            }
    _MEDIAS = tmp
    print(f'[MEDIAS] {len(_MEDIAS)} equipas carregadas de model_data.json')

def _load_arbitros():
    """Carrega médias de árbitros a partir de model_data.json (projecto Faltas, leitura apenas)."""
    global _ARBITROS
    md = _load_model_json()
    if not md:
        print('[ARBITROS] sem model_data.json — árbitros indisponíveis')
        return
    refs_data = md.get('refs', {})
    liga_map  = {'PPL': 'Tugao', 'ESP': 'Spain', 'BRA': 'Brazil'}
    tmp = {}
    for liga_code, liga_key in liga_map.items():
        refs = refs_data.get(liga_key, [])
        arbs = {}
        for r in refs:
            nome = r.get('name', '')
            if not nome:
                continue
            arbs[_norm_nome(nome)] = {
                'nome':         nome,
                'jogos':        r.get('n_games', 0),
                'media_faltas': r.get('avg_total', 0),
                'last_totals':  r.get('last_totals', []),
            }
        tmp[liga_code] = arbs
        print(f'[ARBITROS] {liga_code}: {len(arbs)} árbitros de model_data.json')
    _ARBITROS = tmp

_load_medias()   # carregar no arranque
_load_arbitros() # carregar árbitros no arranque

# Aliases Excel Lancamentos → chave canónica em _MEDIAS (model_data.json)
_LANC_ALIASES = {
    'afs':                   'avs',         # Excel usa 'Afs', model_data usa 'AVS'
    'fc porto':              'porto',
    'sporting cp':           'sp lisbon',
    'estrela da amadora':    'estrela',
    'braga':                 'sp braga',
    'vitoria guimaraes':     'guimaraes',
    'vitoria sc':            'guimaraes',
    'cs maritimo':           'maritimo',
    'atletico mg':           'atletico mg',
    'botafogo rj':           'botafogo rj',
    'flamengo rj':           'flamengo rj',
    'fluminense fc':         'fluminense',
    'atletico-mg':           'atletico mg',
    'ahtletico pr':          'ahtletico pr',
    'rayo vallecano':        'vallecano',
    'atletico madrid':       'ath madrid',
    'athletic bilbao':       'ath bilbao',
    'athletic club':         'ath bilbao',
    'real betis':            'betis',
    'real sociedad':         'sociedad',
    'celta vigo':            'celta',
}
_LANC_XLSX_PATHS = [
    BASE / 'data' / 'medias_equipas_MULTI_LIGA.xlsx',
    DATA_DIR / 'medias_equipas_MULTI_LIGA.xlsx',
    BASE / 'medias_equipas_MULTI_LIGA.xlsx',
]

def _load_lanc_medias():
    """Enriquece _MEDIAS com lanc_media e posse_media do Excel do projecto Lancamentos."""
    global _MEDIAS
    for p in _LANC_XLSX_PATHS:
        if not Path(p).exists():
            continue
        try:
            import importlib
            pd = importlib.import_module('pandas')
            # ler todas as sheets "Geral" e concatenar (evita ler só a 1ª sheet)
            _xl = pd.ExcelFile(str(p))
            _dfs = [pd.read_excel(str(p), sheet_name=s)
                    for s in _xl.sheet_names if 'Geral' in s]
            df = pd.concat(_dfs, ignore_index=True) if _dfs else pd.read_excel(str(p))
            matched, skipped = 0, 0
            for _, row in df.iterrows():
                raw_key = _norm_nome(str(row.get('Equipa', '')))
                lanc    = float(row.get('Lançamentos_media', 0) or 0)
                _pv = row.get('Posse de bola_media', 0)
                posse   = 0.0 if (_pv is None or (isinstance(_pv, float) and _pv != _pv)) else float(_pv or 0)
                # resolver chave canonical
                canon = _LANC_ALIASES.get(raw_key, raw_key)
                if canon in _MEDIAS:
                    _MEDIAS[canon]['lanc_media']  = round(lanc, 2)
                    _MEDIAS[canon]['posse_media'] = round(posse, 1)
                    matched += 1
                else:
                    # fallback: substring/word match contra _MEDIAS existente
                    words = {w for w in raw_key.split() if len(w) > 3}
                    found = None
                    for k in _MEDIAS:
                        k_words = {w for w in k.split() if len(w) > 3}
                        if words & k_words:
                            found = k; break
                    if found:
                        _MEDIAS[found]['lanc_media']  = round(lanc, 2)
                        _MEDIAS[found]['posse_media'] = round(posse, 1)
                        matched += 1
                    else:
                        skipped += 1
            print(f'[LANC] {matched} equipas enriquecidas com lanc_media ({skipped} sem match) de {Path(p).name}')
            return
        except Exception as e:
            print(f'[LANC] erro a carregar {p}: {e}')
    print('[LANC] medias_equipas_MULTI_LIGA.xlsx nao encontrado')

_load_lanc_medias()  # enriquecer _MEDIAS com lançamentos após faltas

def _get_arbitro_stats(liga_code, nome_arbitro):
    """Devolve stats do árbitro ou None."""
    if not nome_arbitro or not _ARBITROS:
        return None
    liga_arbs = _ARBITROS.get(liga_code, {})
    if not liga_arbs:
        return None
    key = _norm_nome(nome_arbitro)
    # 1. exact
    if key in liga_arbs:
        return liga_arbs[key]
    # 2. substring
    for k, v in liga_arbs.items():
        if key in k or k in key:
            return v
    # 3. word-based: qualquer palavra significativa (len>2) em comum
    key_words = {w for w in key.split() if len(w) > 2}
    if key_words:
        for k, v in liga_arbs.items():
            k_words = {w for w in k.split() if len(w) > 2}
            if key_words & k_words:
                return v
    return None

_TSDB_ALIASES = {
    # TheSportsDB norm → model_data.json canonical norm (Tugao/PPL)
    'fc porto':                 'porto',
    'sporting cp':              'sp lisbon',
    'sporting lisbon':          'sp lisbon',
    'sporting clube de portugal':'sp lisbon',
    'sc braga':                 'sp braga',
    'sporting braga':           'sp braga',
    'vitoria sc':               'guimaraes',
    'vitoria de guimaraes':     'guimaraes',
    'vitoria guimaraes':        'guimaraes',
    'maritimo':                 'maritimo',
    'cs maritimo':              'maritimo',
    'academico de viseu':       'academico viseu',
    'academico viseu fc':       'academico viseu',
    'cd nacional':              'nacional',
    'estrela da amadora':       'estrela',
    'estrela amadora':          'estrela',
    'cf estrela da amadora':    'estrela',
    'afs':                      'avs',         # Flashscore usa 'Afs' para AVS
    'avs futebol':              'avs',
    'avs futebol sad':          'avs',
    'casa pia ac':              'casa pia',
    'gil vicente fc':           'gil vicente',
    'fc famalicao':             'famalicao',
    'rio ave fc':               'rio ave',
    'fc arouca':                'arouca',
    'gd estoril praia':         'estoril',
    'estoril praia':            'estoril',
    'moreirense fc':            'moreirense',
    'cf alverca':               'alverca',
    'cd tondela':               'tondela',
    # TheSportsDB norm → model_data.json canonical norm (Spain)
    'atletico madrid':          'ath madrid',
    'atletico de madrid':       'ath madrid',
    'club atletico de madrid':  'ath madrid',
    'athletic club':            'ath bilbao',
    'athletic bilbao':          'ath bilbao',
    'real betis':               'betis',
    'real betis balompie':      'betis',
    'rcd espanyol':             'espanol',
    'espanyol':                 'espanol',
    'rayo vallecano':           'vallecano',
    'real sociedad':            'sociedad',
    'real sociedad de futbol':  'sociedad',
    'celta vigo':               'celta',
    'rc celta':                 'celta',
    'deportivo alaves':         'alaves',
    'deportivo de la coruna':   'la coruna',
    'racing santander':         'santander',
    'rcd mallorca':             'mallorca',
    'cd mallorca':              'mallorca',
    'valencia cf':              'valencia',
    'girona fc':                'girona',  # may not be in model
    'getafe cf':                'getafe',
    'cd leganes':               'leganes',
    # TheSportsDB norm → model_data.json canonical norm (Brazil)
    'atletico mineiro':         'atletico mg',
    'atletico-mineiro':         'atletico mg',
    'clube atletico mineiro':   'atletico mg',
    'athletico paranaense':     'ahtletico pr',
    'atletico paranaense':      'ahtletico pr',
    'ca paranaense':            'ahtletico pr',
    'fluminense fc':            'fluminense',
    'botafogo fr':              'botafogo rj',
    'botafogo':                 'botafogo rj',
    'cruzeiro ec':              'cruzeiro',
    'sport club corinthians':   'corinthians',
    'sc corinthians':           'corinthians',
    'gremio fbpa':              'gremio',
    'gremio foot-ball porto alegrense': 'gremio',
    'santos fc':                'santos',
    'sao paulo fc':             'sao paulo',
    'red bull bragantino':      'bragantino',
    'rb bragantino':            'bragantino',
    'clube de regatas vasco':   'vasco',
    'cr vasco da gama':         'vasco',
    'ec bahia':                 'bahia',
    'sport club internacional': 'internacional',
    'se palmeiras':             'palmeiras',
    'flamengo':                 'flamengo rj',
    'cr flamengo':              'flamengo rj',
}

def _get_stats(nome):
    """Devolve stats da equipa ou None."""
    if not _MEDIAS: return None
    key = _norm_nome(nome)
    # 1. exact match
    if key in _MEDIAS: return _MEDIAS[key]
    # 2. alias manual
    alias = _TSDB_ALIASES.get(key)
    if alias and alias in _MEDIAS: return _MEDIAS[alias]
    # 3. substring (ex: "porto" in "fc porto")
    for k, v in _MEDIAS.items():
        if key in k or k in key: return v
    # 4. word-overlap: partilhar >= 2 palavras significativas (len>2)
    key_words = {w for w in key.split() if len(w) > 2}
    best_score, best_val = 0, None
    for k, v in _MEDIAS.items():
        k_words = {w for w in k.split() if len(w) > 2}
        common = key_words & k_words
        if len(common) >= 2 and len(common) > best_score:
            best_score, best_val = len(common), v
    if best_val: return best_val
    # 5. single distinctive word (len>4) match
    for w in key_words:
        if len(w) > 4:
            for k, v in _MEDIAS.items():
                if w in k.split(): return v
    return None

app = Flask(__name__)
log_buffer = []          # guarda últimas 200 linhas de log
_running   = False       # lock para evitar runs simultâneos

# ── Telegram ──────────────────────────────────────────────────────────────────
TG_TOKEN   = os.environ.get('TELEGRAM_TOKEN', '8934065469:AAEHwZB1z7_lpSg1JTKwc66_4nqXpEsG1Sw')
TG_CHAT_ID = os.environ.get('TELEGRAM_CHAT_ID', '552636700')

# Cache de alertas: {(flash_mid, mercado): ultimo_sinal_enviado}
_alert_cache: dict = {}

def send_telegram(msg):
    try:
        url  = f'https://api.telegram.org/bot{TG_TOKEN}/sendMessage'
        data = urllib.parse.urlencode({'chat_id': TG_CHAT_ID, 'text': msg, 'parse_mode': 'HTML'}).encode()
        urllib.request.urlopen(url, data=data, timeout=10)
    except Exception as e:
        logger.warning(f'[TG] Falha ao enviar: {e}')

# ── Logging ───────────────────────────────────────────────────────────────────
class BufferHandler(logging.Handler):
    def emit(self, record):
        msg = self.format(record)
        log_buffer.append(msg)
        if len(log_buffer) > 200:
            log_buffer.pop(0)

logging.basicConfig(level=logging.INFO,
    format='%(asctime)s  %(message)s',
    datefmt='%H:%M:%S',
    handlers=[logging.StreamHandler(sys.stdout), BufferHandler()])
logger = logging.getLogger('lancamentos')

# Redirecionar print() para logger
import builtins
_orig_print = builtins.print
def _log_print(*args, **kwargs):
    msg = ' '.join(str(a) for a in args)
    log_buffer.append(msg)
    if len(log_buffer) > 200:
        log_buffer.pop(0)
    _orig_print(*args, **kwargs)
builtins.print = _log_print

# ── Monitor ───────────────────────────────────────────────────────────────────
def correr_monitor():
    global _running
    if _running:
        logger.info('[MONITOR] Já a correr — a aguardar...')
        return
    _running = True
    try:
        logger.info('[MONITOR] ── Início ──────────────────────────')
        sys.path.insert(0, str(BASE))

        # IDs já existentes antes do run
        ids_antes = set()
        if DB_PATH.exists():
            con = sqlite3.connect(DB_PATH)
            ids_antes = {r[0] for r in con.execute('SELECT id FROM lancamentos_lines')}
            con.close()

        import monitor_lancamentos as m
        alerts = m.run_once(verbose=True)
        logger.info(f'[MONITOR] Concluído. {alerts} alertas.')

        # Verificar novos jogos e notificar via Telegram
        EDGE_THRESHOLD = 4.0
        if DB_PATH.exists():
            con = sqlite3.connect(DB_PATH)
            con.row_factory = sqlite3.Row
            rows = con.execute('SELECT * FROM lancamentos_lines').fetchall()
            con.close()
            for r in rows:
                if r['id'] not in ids_antes:
                    sinal  = r['edge_signal'] or '—'
                    linha  = r['line']
                    pred   = r['modelo_pred']
                    odd_o  = f"{r['over_odds']:.2f}" if r['over_odds'] else '—'
                    odd_u  = f"{r['under_odds']:.2f}" if r['under_odds'] else '—'

                    # Calcular edge_score
                    edge_score = abs(pred - linha) if (pred and linha) else 0
                    high_value = edge_score >= EDGE_THRESHOLD

                    if high_value:
                        emoji = '🔥'
                        tag   = f'<b>HIGH VALUE</b> (edge: {edge_score:.1f})'
                    else:
                        emoji = '📋'
                        tag   = f'Registo (edge: {edge_score:.1f})'

                    msg = (
                        f"{emoji} <b>NOVO LANÇAMENTO</b> — {tag}\n"
                        f"🏆 {r['league']}\n"
                        f"🏠 {r['home']} vs {r['away']}\n"
                        f"🕐 {r['kickoff']}\n"
                        f"📊 Linha: <b>{linha}</b> | Modelo: <b>{pred:.1f}</b> | Signal: <b>{sinal}</b>\n"
                        f"💰 Over: {odd_o} | Under: {odd_u}"
                    )
                    send_telegram(msg)
                    nivel = 'HIGH VALUE' if high_value else 'registo'
                    logger.info(f'[TG] {nivel}: {r["home"]} vs {r["away"]} (edge={edge_score:.1f})')

        # Regenerar dashboard
        import gerar_dashboard as gd
        gd.gerar()
        logger.info('[MONITOR] Dashboard actualizado.')
    except Exception as e:
        import traceback
        logger.error(f'[MONITOR] ERRO: {e}')
        logger.error(traceback.format_exc())
    finally:
        _running = False

# ── DB helpers ────────────────────────────────────────────────────────────────
def ler_jogos():
    if not DB_PATH.exists():
        return []
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    cols = {r[1] for r in con.execute('PRAGMA table_info(lancamentos_lines)')}
    opcionals = ['resultado_real','lanc_real_casa','lanc_real_fora','resultado','flash_mid']
    sel_extra = ', '.join(c for c in opcionals if c in cols)
    sel = 'id,league,home,away,kickoff,line,over_odds,under_odds,modelo_pred,edge_signal,detected_at'
    if sel_extra:
        sel += ', ' + sel_extra
    rows = con.execute(f'SELECT {sel} FROM lancamentos_lines ORDER BY kickoff DESC').fetchall()
    con.close()
    result = []
    for r in rows:
        d = dict(r)
        for c in opcionals:
            d.setdefault(c, None)
        result.append(d)
    return result

def stats(jogos):
    fechados = [j for j in jogos if j.get('resultado')]
    greens   = sum(1 for j in fechados if j['resultado'] == 'GREEN')
    reds     = sum(1 for j in fechados if j['resultado'] == 'RED')
    win_rate = round(greens / len(fechados) * 100) if fechados else None
    return {'total': len(jogos), 'fechados': len(fechados),
            'greens': greens, 'reds': reds, 'win_rate': win_rate}

def fmt_dt(s):
    if not s: return '—'
    try:
        dt = datetime.fromisoformat(s)
        return dt.strftime('%d/%m %H:%M')
    except:
        return s

# ── Shell com sidebar ─────────────────────────────────────────────────────────
SHELL = """<!DOCTYPE html>
<html lang="pt">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Lançamentos Laterais</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;overflow:hidden;background:#0e1b2e;font-family:'Segoe UI',sans-serif}
.layout{display:flex;height:100vh}
/* Sidebar */
.sidebar{width:200px;min-width:200px;background:#0e1b2e;border-right:1px solid #1e2d42;
         display:flex;flex-direction:column;padding:0}
.sidebar-logo{padding:20px 16px 16px;border-bottom:1px solid #1e2d42}
.sidebar-logo .title{font-size:.95rem;font-weight:700;color:#fff;line-height:1.2}
.sidebar-logo .sub{font-size:.65rem;color:#555;margin-top:2px}
.nav{flex:1;padding:12px 8px}
.nav-item{display:flex;align-items:center;gap:10px;padding:10px 12px;border-radius:8px;
          cursor:pointer;color:#888;font-size:.82rem;font-weight:500;margin-bottom:2px;
          transition:all .15s;text-decoration:none;border:none;background:none;width:100%}
.nav-item:hover{background:#1a1d27;color:#ccc}
.nav-item.active{background:#1565c0;color:#fff}
.nav-item .icon{font-size:1rem;width:20px;text-align:center}
.sidebar-footer{padding:12px 16px;border-top:1px solid #1e2d42;font-size:.65rem;color:#333}
/* Content */
.content{flex:1;overflow:hidden}
iframe{width:100%;height:100%;border:none;display:block}
</style>
</head>
<body>
<div class="layout">
  <div class="sidebar">
    <div class="sidebar-logo">
      <div class="title">⚽ Lançamentos</div>
      <div class="sub">Laterais · Monitor</div>
    </div>
    <nav class="nav">
      <button class="nav-item active" onclick="navTo('/dashboard','btn-dash')" id="btn-dash">
        <span class="icon">📊</span> Dashboard
      </button>
      <button class="nav-item" onclick="navTo('/calc','btn-calc')" id="btn-calc">
        <span class="icon">🧮</span> Calculadora
      </button>
      <button class="nav-item" onclick="navTo('/registo','btn-reg')" id="btn-reg">
        <span class="icon">📋</span> Registo
      </button>
      <button class="nav-item" onclick="navTo('/live','btn-live')" id="btn-live">
        <span class="icon">🟢</span> Live Data
      </button>
      <button class="nav-item" onclick="navTo('/jogos','btn-jogos')" id="btn-jogos">
        <span class="icon">📅</span> Jogos FDS
      </button>
      <button class="nav-item" onclick="navTo('/live/historico','btn-hist')" id="btn-hist">
        <span class="icon">📊</span> Histórico
      </button>
    </nav>
    <div class="sidebar-footer">lancamentos-laterais.up.railway.app</div>
  </div>
  <div class="content">
    <iframe id="frame" src="/dashboard"></iframe>
  </div>
</div>
<script>
function navTo(url, btnId){
  document.getElementById('frame').src = url;
  document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));
  document.getElementById(btnId).classList.add('active');
}
// Restaurar nav ativo pelo hash
const h = location.hash.replace('#','');
if(h) navTo('/'+h, 'btn-'+h.split('/')[0]);
</script>
</body>
</html>"""

# ── Dashboard template ────────────────────────────────────────────────────────
TEMPLATE = """<!DOCTYPE html>
<html lang="pt">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Dashboard</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',sans-serif;background:#0e1b2e;color:#e0e0e0;padding:20px}
.header{display:flex;align-items:center;justify-content:space-between;margin-bottom:20px}
h1{font-size:1.3rem;color:#fff}
.sub{font-size:.75rem;color:#555;margin-top:2px}
.btn{padding:8px 16px;border:none;border-radius:6px;cursor:pointer;font-size:.82rem;font-weight:600}
.btn-run{background:#1565c0;color:#fff}
.btn-run:hover{background:#1976d2}
.btn-run:disabled{background:#333;color:#666;cursor:not-allowed}
.stats{display:flex;gap:10px;margin-bottom:20px;flex-wrap:wrap}
.stat{background:#1a1d27;border:1px solid #2a2d3a;border-radius:8px;padding:12px 18px;min-width:100px}
.stat .lbl{font-size:.68rem;color:#888;text-transform:uppercase;letter-spacing:.05em}
.stat .val{font-size:1.5rem;font-weight:700;margin-top:2px;color:#fff}
.stat.g .val{color:#4caf50}.stat.r .val{color:#f44336}.stat.b .val{color:#64b5f6}
table{width:100%;border-collapse:collapse;font-size:.8rem}
thead th{background:#1a1d27;color:#aaa;font-weight:600;padding:9px 7px;text-align:left;
         border-bottom:1px solid #2a2d3a;white-space:nowrap}
tbody tr{border-bottom:1px solid #1a1d27;transition:background .12s}
tbody tr:hover{background:#1a1d27}
td{padding:8px 7px;vertical-align:middle}
td.num{text-align:right;font-variant-numeric:tabular-nums}
.sig{color:#90caf9;font-size:.72rem;font-weight:600}
.res{font-weight:700;font-size:.78rem}
.res.GREEN{color:#4caf50}.res.RED{color:#f44336}.res.Void{color:#888}
tr.rg{background:rgba(76,175,80,.04)}tr.rr{background:rgba(244,67,54,.04)}
.hv{background:#ff6f00;color:#fff;font-size:.65rem;font-weight:700;padding:2px 6px;
    border-radius:4px;display:inline-block;margin-left:4px}
.edge-score{font-size:.7rem;color:#888}
.log-box{margin-top:24px;background:#0b1525;border:1px solid #1e2d42;border-radius:8px;
         padding:12px;max-height:260px;overflow-y:auto;font-family:monospace;font-size:.75rem;color:#aaa}
.log-box .err{color:#f44336}.log-box .ok{color:#4caf50}
.status{display:inline-block;width:8px;height:8px;border-radius:50%;
        background:{% if running %}#4caf50{% else %}#555{% endif %};margin-right:6px}
.refresh-bar{font-size:.7rem;color:#555;margin-top:4px}
</style>
</head>
<body>
<div class="header">
  <div>
    <h1>Monitor</h1>
    <div class="sub">
      <span class="status"></span>
      {% if running %}A correr monitor...{% else %}Última execução: {{ ultima_exec }}{% endif %}
      &nbsp;·&nbsp; Próxima: {{ proxima_exec }}
    </div>
    <div class="refresh-bar">Auto-refresh em <span id="cnt">60</span>s</div>
  </div>
  <div style="display:flex;gap:8px">
    <button class="btn btn-run" id="btnRun" onclick="correrAgora()"
      {% if running %}disabled{% endif %}>
      ▶ Correr Agora
    </button>
  </div>
</div>

<div class="stats">
  <div class="stat"><div class="lbl">Total</div><div class="val">{{ st.total }}</div></div>
  <div class="stat"><div class="lbl">Fechados</div><div class="val">{{ st.fechados }}</div></div>
  <div class="stat g"><div class="lbl">GREEN</div><div class="val">{{ st.greens }}</div></div>
  <div class="stat r"><div class="lbl">RED</div><div class="val">{{ st.reds }}</div></div>
  <div class="stat b"><div class="lbl">Win Rate</div><div class="val">{% if st.win_rate is not none %}{{ st.win_rate }}%{% else %}—{% endif %}</div></div>
</div>

{% if jogos %}
<table>
<thead>
  <tr><th>Liga</th><th>Casa</th><th>Fora</th><th>KO</th>
      <th>Linha</th><th>Modelo</th><th>Edge</th><th>Signal</th>
      <th>Odd O</th><th>Odd U</th><th>Real</th><th>Resultado</th></tr>
</thead>
<tbody>
{% for j in jogos %}
<tr class="{% if j.resultado=='GREEN' %}rg{% elif j.resultado=='RED' %}rr{% endif %}">
  <td>{{ j.league }}</td>
  <td>{{ j.home }}</td>
  <td>{{ j.away }}</td>
  <td>{{ j.kickoff|fmt }}</td>
  <td class="num">{{ j.line or '—' }}</td>
  <td class="num">{{ '%.1f'|format(j.modelo_pred) if j.modelo_pred else '—' }}</td>
  <td class="num edge-score">{% if j.modelo_pred and j.line %}{% set es = (j.modelo_pred - j.line)|abs %}{{ '%.1f'|format(es) }}{% if es >= 4 %}<span class="hv">HV</span>{% endif %}{% else %}—{% endif %}</td>
  <td class="sig">{{ j.edge_signal or '—' }}</td>
  <td class="num">{{ '%.2f'|format(j.over_odds) if j.over_odds else '—' }}</td>
  <td class="num">{{ '%.2f'|format(j.under_odds) if j.under_odds else '—' }}</td>
  <td class="num">{{ j.resultado_real|int if j.resultado_real else '—' }}</td>
  <td class="res {{ j.resultado or '' }}">{{ j.resultado or 'Pendente' }}</td>
</tr>
{% endfor %}
</tbody>
</table>
{% else %}
<p style="color:#555;margin-top:20px">Sem jogos na DB ainda. Clica em "Correr Agora" para verificar.</p>
{% endif %}

<div class="log-box" id="logBox">
{% for line in logs %}
<div class="{% if 'ERRO' in line or 'ERROR' in line %}err{% elif 'NOVO' in line or 'GREEN' in line %}ok{% endif %}">{{ line }}</div>
{% endfor %}
</div>

<script>
let cnt = 60;
const el = document.getElementById('cnt');
setInterval(()=>{ cnt--; if(cnt<=0){location.reload()}else{el.textContent=cnt} },1000);
const lb = document.getElementById('logBox');
if(lb) lb.scrollTop = lb.scrollHeight;
function correrAgora(){
  document.getElementById('btnRun').disabled = true;
  document.getElementById('btnRun').textContent = '⏳ A correr...';
  fetch('/run').then(r=>r.json()).then(()=>{ setTimeout(()=>location.reload(), 2000); });
}
</script>
</body>
</html>"""

# ── Jinja filter ──────────────────────────────────────────────────────────────
@app.template_filter('fmt')
def fmt_filter(s):
    return fmt_dt(s)

# ── Rotas ─────────────────────────────────────────────────────────────────────
ultima_exec  = 'Nunca'
proxima_exec = '—'

@app.route('/')
def index():
    return render_template_string(SHELL)

@app.route('/dashboard')
def dashboard():
    jogos = ler_jogos()
    return render_template_string(TEMPLATE,
        jogos       = jogos,
        st          = stats(jogos),
        logs        = log_buffer[-80:],
        running     = _running,
        ultima_exec = ultima_exec,
        proxima_exec= proxima_exec,
    )

@app.route('/calc')
def calc():
    return send_file(BASE / 'lançamentos_calc.html')

@app.route('/registo')
def registo():
    return send_file(BASE / 'lançamentos_registo.html')

@app.route('/run')
def run_now():
    t = threading.Thread(target=correr_monitor, daemon=True)
    t.start()
    return jsonify({'ok': True})

@app.route('/status')
def status():
    return jsonify({'running': _running, 'logs': log_buffer[-20:]})

# ── Scheduler ─────────────────────────────────────────────────────────────────
def job_monitor():
    global ultima_exec, proxima_exec
    ultima_exec = datetime.now().strftime('%d/%m %H:%M')
    correr_monitor()

scheduler = BackgroundScheduler(timezone='Europe/Lisbon')
scheduler.add_job(job_monitor, 'interval', hours=2, id='monitor',
                  next_run_time=datetime.now())   # corre imediatamente no arranque

scheduler.start()

# Calcular próxima execução
def actualizar_proxima():
    global proxima_exec
    try:
        job = scheduler.get_job('monitor')
        if job and job.next_run_time:
            proxima_exec = job.next_run_time.strftime('%d/%m %H:%M')
    except:
        pass

import atexit
atexit.register(lambda: scheduler.shutdown())

# ── Live Collector (corre no Railway) ─────────────────────────────────────────
import re as _re

_NINJA_URL = 'https://global.flashscore.ninja/20/x/feed/df_st_1_{mid}'
_FS_HDR = {
    'X-Fsign':    'SW9D1eZo',
    'Referer':    'https://www.flashscore.com/',
    'Accept':     '*/*',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
}
_STAT_RE     = _re.compile(r'SG÷([^¬]+)¬SH÷([^¬]*)¬SI÷([^¬~]*)')
_LANC_BASE   = {'PPL': 39.1, 'PL': 35.8, 'BRA': 36.7, 'ESP': 40.5, 'SCOPA': 37.0}
_FALTAS_BASE = {'PPL': 26.78, 'PL': 22.25, 'BRA': 26.76, 'ESP': 25.41, 'SCOPA': 22.0}  # médias reais do modelo

# ── Modelo de Faltas (extraído de model_data.json — só leitura) ───────────────
# Regressão linear: pred = intercept + ffh*home_ffh + fsh*home_fsh + ffa*away_ffa + fsa*away_fsa
# ffh=faltas feitas em casa, fsh=faltas sofridas em casa, ffa=feitas fora, fsa=sofridas fora
_FALTAS_COEFS = {
    'PPL': dict(intercept=-19.989, ffh=0.8794, fsh=0.9549, ffa=0.7740, fsa=0.8836),
    'BRA': dict(intercept=-26.000, ffh=1.0106, fsh=0.9728, ffa=0.9165, fsa=1.0432),
    'ESP': dict(intercept=-15.967, ffh=0.8549, fsh=0.6702, ffa=0.7849, fsa=0.9559),
}
_FALTAS_TEAMS = {
    'PPL': {
        'Porto':           dict(ffh=15.17, fsh=13.00, ffa=14.20, fsa=11.60),
        'Sp Lisbon':       dict(ffh=12.20, fsh=13.00, ffa=11.17, fsa=13.50),
        'Benfica':         dict(ffh=12.00, fsh=14.67, ffa=11.00, fsa=15.60),
        'Sp Braga':        dict(ffh=12.67, fsh=11.67, ffa=13.33, fsa=11.50),
        'Gil Vicente':     dict(ffh=12.67, fsh=13.83, ffa=14.20, fsa=13.60),
        'Famalicao':       dict(ffh=13.80, fsh=14.20, ffa=11.83, fsa=16.50),
        'Estoril':         dict(ffh=13.17, fsh=12.33, ffa=13.60, fsa=15.60),
        'Moreirense':      dict(ffh=12.86, fsh=14.43, ffa=13.00, fsa=15.60),
        'Guimaraes':       dict(ffh=13.17, fsh=11.83, ffa=13.50, fsa=15.33),
        'Alverca':         dict(ffh=11.60, fsh=12.80, ffa=10.14, fsa=13.71),
        'Rio Ave':         dict(ffh=10.00, fsh= 9.60, ffa=12.17, fsa=14.50),
        'Arouca':          dict(ffh=14.40, fsh=12.00, ffa=12.67, fsa=10.33),
        'Estrela':         dict(ffh=15.33, fsh=15.17, ffa=11.17, fsa=15.50),
        'Santa Clara':     dict(ffh=16.50, fsh=14.17, ffa=12.20, fsa=14.80),
        'Casa Pia':        dict(ffh=18.17, fsh=12.67, ffa=15.67, fsa=13.50),
        'Nacional':        dict(ffh=14.17, fsh=13.00, ffa=15.67, fsa=13.33),
        'Maritimo':        dict(ffh=14.83, fsh=13.00, ffa=15.20, fsa=12.40),
        'Academico Viseu': dict(ffh=14.40, fsh=13.00, ffa=14.17, fsa=13.50),
        'Tondela':         dict(ffh=14.40, fsh=12.80, ffa=13.40, fsa=15.20),
        'AVS':             dict(ffh=14.20, fsh= 9.00, ffa=13.33, fsa= 9.00),
    },
    'BRA': {
        'Ahtletico PR':  dict(ffh=14.27, fsh=13.09, ffa=14.09, fsa=11.64),
        'Atletico MG':   dict(ffh=10.44, fsh=15.00, ffa=11.58, fsa=15.42),
        'Bahia':         dict(ffh=15.50, fsh=15.92, ffa=13.00, fsa=13.00),
        'Botafogo RJ':   dict(ffh=14.60, fsh=12.20, ffa=13.64, fsa=13.18),
        'Bragantino':    dict(ffh=15.00, fsh=12.00, ffa=15.30, fsa=11.70),
        'Chapecoense':   dict(ffh=11.50, fsh= 8.50, ffa=13.64, fsa=12.64),
        'Corinthians':   dict(ffh=15.27, fsh=15.73, ffa=13.64, fsa=16.73),
        'Coritiba':      dict(ffh=11.18, fsh=14.18, ffa=10.27, fsa=12.09),
        'Cruzeiro':      dict(ffh=13.83, fsh=15.58, ffa=15.30, fsa=16.70),
        'Flamengo RJ':   dict(ffh=12.90, fsh=12.20, ffa=12.36, fsa=12.45),
        'Fluminense':    dict(ffh=12.18, fsh=11.18, ffa=14.09, fsa=13.27),
        'Gremio':        dict(ffh=13.82, fsh=12.91, ffa=12.70, fsa=14.10),
        'Internacional': dict(ffh=16.00, fsh=14.09, ffa=15.09, fsa=14.55),
        'Mirassol':      dict(ffh=12.91, fsh=11.18, ffa=12.70, fsa=12.10),
        'Palmeiras':     dict(ffh=12.09, fsh=14.82, ffa=15.73, fsa=11.64),
        'Remo':          dict(ffh=12.27, fsh=12.18, ffa=10.73, fsa=12.55),
        'Santos':        dict(ffh=14.17, fsh=14.08, ffa=12.44, fsa=11.22),
        'Sao Paulo':     dict(ffh=13.56, fsh=14.89, ffa=15.92, fsa=13.00),
        'Vasco':         dict(ffh=12.00, fsh=12.73, ffa=12.00, fsa=13.82),
        'Vitoria':       dict(ffh=13.50, fsh=14.30, ffa=13.00, fsa=15.64),
    },
    'ESP': {
        'Alaves':      dict(ffh=11.40, fsh=13.20, ffa=15.40, fsa=13.20),
        'Ath Bilbao':  dict(ffh=12.00, fsh= 8.60, ffa=14.00, fsa= 8.40),
        'Ath Madrid':  dict(ffh=10.40, fsh=11.20, ffa=12.50, fsa= 7.83),
        'Barcelona':   dict(ffh=12.00, fsh=11.20, ffa= 9.00, fsa=12.80),
        'Betis':       dict(ffh=11.60, fsh=10.00, ffa=10.60, fsa=10.80),
        'Celta':       dict(ffh=10.80, fsh=14.80, ffa=14.20, fsa=13.00),
        'Elche':       dict(ffh=14.20, fsh=13.40, ffa=14.20, fsa=17.60),
        'Espanol':     dict(ffh=13.00, fsh=16.40, ffa=16.20, fsa=11.60),
        'Getafe':      dict(ffh=13.40, fsh=15.80, ffa=17.67, fsa=12.83),
        'Malaga':      dict(ffh=12.00, fsh=18.20, ffa= 9.20, fsa=13.20),
        'Levante':     dict(ffh=13.60, fsh=17.20, ffa=10.80, fsa=11.00),
        'Santander':   dict(ffh=13.00, fsh=11.40, ffa=16.60, fsa=12.60),
        'Osasuna':     dict(ffh=12.83, fsh=11.17, ffa=14.40, fsa=12.60),
        'La Coruna':   dict(ffh=12.60, fsh=14.20, ffa=11.40, fsa=10.40),
        'Real Madrid': dict(ffh= 7.60, fsh=14.80, ffa=11.60, fsa=16.20),
        'Sevilla':     dict(ffh=14.20, fsh=11.60, ffa=14.00, fsa=12.20),
        'Sociedad':    dict(ffh=12.67, fsh=14.33, ffa=14.00, fsa=13.40),
        'Valencia':    dict(ffh=13.00, fsh=13.80, ffa=10.80, fsa=11.80),
        'Vallecano':   dict(ffh=13.00, fsh=11.00, ffa=15.60, fsa=10.80),
        'Villarreal':  dict(ffh=10.20, fsh=11.00, ffa=10.20, fsa=12.20),
    },
}
# Aliases: nomes usados no setup_weekend.py → nomes no modelo de faltas
_FALTAS_ALIAS = {
    # PPL
    'Sporting CP': 'Sp Lisbon', 'Vitoria Guimaraes': 'Guimaraes',
    'FC Porto': 'Porto', 'SC Braga': 'Sp Braga',
    # BRA
    'Athletico-PR': 'Ahtletico PR', 'Atletico-MG': 'Atletico MG',
    'Chapecoense-SC': 'Chapecoense',
    # ESP
    'Rayo Vallecano': 'Vallecano', 'Racing Santander': 'Santander',
    'Espanyol': 'Espanol', 'Celta Vigo': 'Celta',
    'Dep. A Coruna': 'La Coruna', 'Atl. Madrid': 'Ath Madrid',
    'Real Sociedad': 'Sociedad',
}

# Médias históricas de árbitros (avg_total = média faltas totais por jogo)
# Fonte: model_data.json › refs (Tugao/Brazil/Spain) — nomes em lowercase para match
_FALTAS_REFS = {
    'PPL': {
        'fonseca miguel': 27.4167, 'jose bessa': 28.4545, 'sergio guelho': 27.5455,
        'gustavo correia': 27.8182, 'hélder carvalho': 28.0909, 'joão gonçalves': 23.6364,
        'luís godinho': 26.5455, 'david silva': 30.0909, 'pedro ramalho': 29.2,
        'ricardo baixinho': 30.6, 'a rodrigues': 25.8, 'bruno costa': 25.3,
        'hélder malheiro': 32.2, 'joão pinheiro': 24.6, 'carlos macedo': 25.9,
        'antónio nobre': 27.2, 'miguel nogueira': 27.1, 'cláudio pereira': 27.8,
        'vasilica': 25.7, 'fábio veríssimo': 25.8, 'filipe l': 26.7778,
        'torres m': 25.4444, 'andré narciso': 32.0, 'diogo rosa': 29.1429,
    },
    'BRA': {
        'lacerda': 23.5185, 'klein': 24.8846, 'pereira de lima': 29.32,
        'arleu': 26.8696, 'abatti': 29.6364, 'fernandes lima': 30.4091,
        'rodrigues souza': 26.3182, 'claus': 22.9091, 'daronco': 28.7619,
        'zanovelli': 30.65, 'w sampaio': 27.3, 'p sampaio': 26.15,
        'gobi': 27.4737, 'casagrande': 27.2632, 'delgado': 25.8421,
        'stefano': 31.7222, 'w nascimento': 26.2941, 'pinheiro': 27.6667,
        'torezin': 27.4286, 'fernando salles filho': 29.2857, 'machado': 23.3077,
        'alves batista': 24.3333, 'vasconcelos': 25.4545, 'henrique': 24.9,
        'policarpo': 31.0, 'd oliveira': 21.7, 'bauemann': 26.8,
        'marcelo de lima': 26.9, 'vuaden': 24.5, 'v augusto': 24.125,
        'cruz': 29.1667, 'serafim ribeiro': 23.3333, 'ferreira j': 23.6667,
        'belence': 23.25, 'lima barbosa': 28.0, 'jefferson': 26.25,
    },
    'ESP': {
        'josé luis guzmán mansilla': 23.3, 'ricardo de burgos bengochea': 23.1,
        'víctor garcía verdura': 25.8, 'jesús gil manzano': 27.7,
        'iosu galech apezteguía': 29.2, 'francisco hernández maeso': 28.1,
        'alejandro hernández hernández': 25.9, 'guillermo cuadra fernández': 22.7,
        'miguel ortiz arias': 22.2, 'césar soto grado': 21.7,
        'alejandro quintero gonzález': 28.4, 'mateo busquets ferrer': 24.1,
        'isidro díaz de mera': 23.8, 'miguel sesma espinosa': 27.7,
        'josé sánchez martínez': 27.0, 'juan martínez munuera': 28.7,
        'javier alberola rojas': 25.6, 'alejandro muñiz ruiz': 23.4,
        'josé munuera montero': 24.2, 'adrián cordero vega': 21.2,
    },
}

def _apply_faltas_model(teams, refs, base, coefs):
    """Actualiza os globais do modelo de faltas em memória."""
    global _FALTAS_TEAMS, _FALTAS_REFS, _FALTAS_BASE, _FALTAS_COEFS
    _FALTAS_TEAMS = teams
    _FALTAS_REFS  = refs
    _FALTAS_BASE  = base
    _FALTAS_COEFS = coefs

def _load_faltas_model_from_json(path):
    """Lê model_data.json e extrai teams, refs, base, coefs para PPL/BRA/ESP."""
    with open(path, encoding='utf-8') as f:
        data = json.load(f)
    new_teams, new_refs, new_base, new_coefs = {}, {}, {}, {}
    for json_liga, mod_liga in _LIGA_MAP.items():
        # Equipas: lista de dicts com campo 'name'
        raw_teams = data.get('teams', {}).get(json_liga, [])
        new_teams[mod_liga] = {
            t['name']: dict(ffh=float(t.get('ffh',0)), fsh=float(t.get('fsh',0)),
                            ffa=float(t.get('ffa',0)), fsa=float(t.get('fsa',0)))
            for t in raw_teams if t.get('name')
        }
        # Árbitros (apenas com jogos)
        new_refs[mod_liga] = {
            r['name']: r['avg_total']
            for r in data.get('refs', {}).get(json_liga, [])
            if r.get('n_games', 0) > 0
        }
        # Médias liga e coeficientes
        new_base[mod_liga]  = data.get('league_avg', {}).get(json_liga, 26.0)
        new_coefs[mod_liga] = data.get('coefs', {}).get(json_liga, {})
    return new_teams, new_refs, new_base, new_coefs

# ── Carrega cache ao arrancar (se existir) ────────────────────────────────────
if FALTAS_CACHE_PATH.exists():
    try:
        with open(FALTAS_CACHE_PATH, encoding='utf-8') as _f:
            _cache = json.load(_f)
        _apply_faltas_model(_cache['teams'], _cache['refs'], _cache['base'], _cache['coefs'])
        logger.info(f'[FALTAS] cache carregado: {_cache.get("updated_at","?")}')
    except Exception as _e:
        logger.warning(f'[FALTAS] falha ao carregar cache: {_e}')

def _faltas_pregame_pred(home, away, liga, referee=None):
    """Previsão pré-jogo de faltas usando o modelo real (regressão linear por liga).
    Fórmula de produção: 0.4*reg_pred + 0.4*media_simples + 0.2*ref_pred
    Devolve float ou None se equipa não encontrada."""
    coefs = _FALTAS_COEFS.get(liga)
    teams = _FALTAS_TEAMS.get(liga, {})
    if not coefs or not teams:
        return None
    h_key = _FALTAS_ALIAS.get(home, home)
    a_key = _FALTAS_ALIAS.get(away, away)
    h = teams.get(h_key)
    a = teams.get(a_key)
    if not h or not a:
        logger.debug(f'[FALTAS] equipa não encontrada: {h_key!r} ou {a_key!r} ({liga})')
        return None
    ffh, fsh = h['ffh'], h['fsh']
    ffa, fsa = a['ffa'], a['fsa']
    reg_pred     = (coefs['intercept']
                    + coefs['ffh'] * ffh + coefs['fsh'] * fsh
                    + coefs['ffa'] * ffa + coefs['fsa'] * fsa)
    media_simples = ffh + ffa
    league_avg    = _FALTAS_BASE.get(liga, 26.0)
    ref_pred      = league_avg
    if referee:
        ref_lower = referee.strip().lower()
        refs_liga = _FALTAS_REFS.get(liga, {})
        # Tentativa 1: match exacto
        if ref_lower in refs_liga:
            ref_pred = refs_liga[ref_lower]
        else:
            # Tentativa 2: match parcial (apelido)
            for name, avg in refs_liga.items():
                if ref_lower in name or name in ref_lower:
                    ref_pred = avg
                    break
        logger.debug(f'[FALTAS-REF] {referee!r} → ref_pred={ref_pred} ({liga})')
    combined = round(0.4 * reg_pred + 0.4 * media_simples + 0.2 * ref_pred, 1)
    return max(combined, 10.0)

def _fetch_ninja(flash_mid):
    try:
        import urllib.request as _ur
        req = _ur.Request(_NINJA_URL.format(mid=flash_mid), headers=_FS_HDR)
        with _ur.urlopen(req, timeout=12) as r:
            raw = r.read().decode('utf-8', errors='replace')
        stats = {}
        def pv(v):
            v = str(v).strip()
            mm = _re.search(r'\((\d+)/(\d+)\)', v)
            if mm: return float(mm.group(2))
            try: return float(v.replace('%',''))
            except: return None
        _LANC_NAMES  = {'Lançamentos', 'Throw-ins', 'Throw Ins', 'Lanzamientos'}
        _FALT_NAMES  = {'Faltas', 'Fouls', 'Fouls committed', 'Faltas cometidas'}
        for m in _STAT_RE.finditer(raw):
            name = m.group(1).strip()
            if name in _LANC_NAMES:
                stats['lanc_casa'] = pv(m.group(2))
                stats['lanc_fora'] = pv(m.group(3))
            elif name in _FALT_NAMES:
                stats['faltas_casa'] = pv(m.group(2))
                stats['faltas_fora'] = pv(m.group(3))
        if stats:
            logger.debug(f'[NINJA] {flash_mid}: {stats}')
        return stats
    except Exception as e:
        logger.warning(f'[NINJA] {flash_mid}: {e}')
        return {}

def _pred(minuto, total, baseline):
    """Previsão blendada ritmo+baseline.
    Usa expoente 1.5 para reduzir o peso do ritmo instantâneo nos primeiros minutos
    (evita sobreestimação causada por ritmos altos no início do jogo).
    """
    if minuto <= 0: return baseline
    pace_proj = total / minuto * 90
    pace_w    = min(minuto / 90, 1.0) ** 1.5   # cresce mais devagar que linear
    return round(pace_w * pace_proj + (1 - pace_w) * baseline, 1)

# ── Statscore live stats ───────────────────────────────────────────────────────
_STATSCORE_BASE = 'https://events-d.pc.statscore.com/get_pushes/{sid}?messageId=0&auth={auth}&poll=false'
_STATSCORE_AUTH_DEFAULT = '9352141a'   # auth estático do Yonibet/Betby

def _fetch_statscore(statscore_id, auth=None):
    """Busca stats Statscore: faltas (stat id 22) e lançamentos laterais (stat id 32)."""
    if not statscore_id:
        return None
    auth = auth or _STATSCORE_AUTH_DEFAULT
    try:
        url = _STATSCORE_BASE.format(sid=statscore_id, auth=auth)
        req = urllib.request.Request(url, headers={'Referer': 'https://start5.sptpub.com/'})
        with urllib.request.urlopen(req, timeout=15) as r:
            data = json.loads(r.read())
        ev_msgs = [m for m in data if m.get('type') == 'event']
        if not ev_msgs:
            return None
        last_ev = ev_msgs[-1]['data']['event']
        parts = last_ev.get('participants', [])
        if len(parts) < 2:
            return None
        def get_stat(part, sid):
            for s in part.get('stats', []):
                if s['id'] == sid and s['value'] is not None:
                    try: return int(round(float(s['value'])))
                    except: pass
            return None
        p0, p1 = parts[0], parts[1]
        # status: 'inprogress', 'halftime', 'finished', etc.
        sc_status = str(last_ev.get('status', '')).lower()
        # minuto real segundo Statscore
        sc_minute = None
        try:
            sc_minute = int(last_ev.get('minute') or last_ev.get('time_of_play') or 0) or None
        except: pass
        return {
            'sc_faltas_casa': get_stat(p0, 22),
            'sc_faltas_fora': get_stat(p1, 22),
            'sc_lanc_casa':   get_stat(p0, 32),
            'sc_lanc_fora':   get_stat(p1, 32),
            'sc_status':      sc_status,
            'sc_minute':      sc_minute,
        }
    except Exception as e:
        logger.warning(f'[STATSCORE] {statscore_id}: {e}')
        return None

# ── SofaScore live stats ───────────────────────────────────────────────────────
_SOFA_HDR = {
    'User-Agent':      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept':          'application/json, text/plain, */*',
    'Accept-Language': 'pt-PT,pt;q=0.9,en;q=0.8',
    'Accept-Encoding': 'gzip, deflate, br',
    'Referer':         'https://www.sofascore.app/',
    'Origin':          'https://www.sofascore.app',
    'Cache-Control':   'no-cache',
    'Pragma':          'no-cache',
    'Sec-Fetch-Dest':  'empty',
    'Sec-Fetch-Mode':  'cors',
    'Sec-Fetch-Site':  'same-origin',
}
_sofa_index_cache: dict = {}   # {flash_mid: sofa_event_id}
_sofa_index_ts:    float = 0.0  # timestamp da última actualização

_sofa_name_index: dict = {}  # {(norm_home, norm_away): sofa_event_id}

def _norm_sofa(name):
    import unicodedata
    return unicodedata.normalize('NFD', name.lower()).encode('ascii', 'ignore').decode().strip()

def _sofa_fetch(url):
    """Fetch SofaScore URL — tenta directo, fallback via allorigins proxy se 403."""
    req = urllib.request.Request(url, headers=_SOFA_HDR)
    try:
        with urllib.request.urlopen(req, timeout=10) as r:
            raw = r.read()
            # descomprimir gzip se necessário
            if r.info().get('Content-Encoding') == 'gzip':
                import gzip
                raw = gzip.decompress(raw)
            return json.loads(raw)
    except Exception as e:
        if '403' not in str(e) and '429' not in str(e):
            raise
        # Fallback: allorigins proxy
        proxy_url = 'https://api.allorigins.win/get?url=' + urllib.parse.quote(url, safe='')
        req2 = urllib.request.Request(proxy_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req2, timeout=15) as r2:
            outer = json.loads(r2.read())
            return json.loads(outer.get('contents', '{}'))

def _refresh_sofascore_index():
    """Busca todos os jogos live do SofaScore e indexa por customId (=flash_mid) e por nomes."""
    global _sofa_index_cache, _sofa_index_ts, _sofa_name_index
    import time
    if time.time() - _sofa_index_ts < 25:   # cache válida por 25s (< ciclo de 30s)
        return _sofa_index_cache
    try:
        url = 'https://api.sofascore.app/api/v1/sport/football/events/live'
        data = _sofa_fetch(url)
        idx = {}
        name_idx = {}
        for ev in data.get('events', []):
            cid = ev.get('customId')   # = flash_mid do Flashscore
            eid = ev['id']
            if cid:
                idx[cid] = eid
            # índice por nome (fallback para pushes manuais)
            h = _norm_sofa(ev.get('homeTeam', {}).get('name', ''))
            a = _norm_sofa(ev.get('awayTeam', {}).get('name', ''))
            if h and a:
                name_idx[(h, a)] = eid
        _sofa_index_cache = idx
        _sofa_name_index  = name_idx
        _sofa_index_ts = time.time()
        logger.debug(f'[SOFA] índice actualizado: {len(idx)} jogos live')
    except Exception as e:
        logger.warning(f'[SOFA] falha ao actualizar índice: {e}')
    return _sofa_index_cache

def _fetch_sofascore(flash_mid, home=None, away=None):
    """Busca stats SofaScore para o jogo — por flash_mid (customId) ou por nome de equipa."""
    idx = _refresh_sofascore_index()
    sofa_id = idx.get(flash_mid)
    # Fallback: match por nome (para pushes manuais com flash_mid personalizado)
    if not sofa_id and home and away:
        h = _norm_sofa(home); a = _norm_sofa(away)
        sofa_id = _sofa_name_index.get((h, a))
        if not sofa_id:
            # Tentar match parcial (ex: "SC Braga" vs "Sporting Braga")
            for (sh, sa), eid in _sofa_name_index.items():
                hwords = {w for w in h.split() if len(w) > 3}
                awords = {w for w in a.split() if len(w) > 3}
                if hwords & {w for w in sh.split() if len(w) > 3} and \
                   awords & {w for w in sa.split() if len(w) > 3}:
                    sofa_id = eid
                    logger.debug(f'[SOFA] match parcial: {home} vs {away} → {sh} vs {sa}')
                    break
    if not sofa_id:
        return None
    try:
        url = f'https://api.sofascore.app/api/v1/event/{sofa_id}/statistics'
        data = _sofa_fetch(url)
        lc = lf = fc = ff = None
        for period in data.get('statistics', []):
            if period.get('period') != 'ALL':
                continue
            for group in period.get('groups', []):
                for item in group.get('statisticsItems', []):
                    name = item.get('name', '')
                    if name == 'Throw-ins':
                        try: lc = int(item['home']); lf = int(item['away'])
                        except: pass
                    elif name == 'Fouls':
                        try: fc = int(item['home']); ff = int(item['away'])
                        except: pass
        if lc is None and fc is None:
            return None
        logger.info(f'[SOFA] {flash_mid}: Lanc={lc}-{lf} Falt={fc}-{ff}')
        return {'sofa_lc': lc, 'sofa_lf': lf, 'sofa_fc': fc, 'sofa_ff': ff}
    except Exception as e:
        logger.warning(f'[SOFA] {flash_mid}: {e}')
        return None

_BET_BASE  = 'https://22bet4me.com/service-api/LiveFeed'
_BET_HDR   = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept':     'application/json, text/plain, */*',
    'Referer':    'https://22bet4me.com/line/football/',
}
_LANC_TI   = 55    # TI confirmado: Lançamentos Laterais (G=17 Over/Under, linhas ~29-35)
_FALTAS_TI = 14    # TI confirmado: Faltas Totais (G=17 Over/Under, linhas ~20-27)

def _bet_api(url):
    import urllib.request as _ur
    req = _ur.Request(url, headers=_BET_HDR)
    with _ur.urlopen(req, timeout=12) as r:
        return json.loads(r.read())

def _parse_ou_odds(ge):
    """Extrai {line, over, under} do array GE de um sub-jogo 22bet."""
    grp = next((g for g in ge if g.get('G') == 17), ge[0] if ge else None)
    if not grp:
        return None
    over_list  = grp.get('E', [[]])[0]  if len(grp.get('E', [])) > 0 else []
    under_list = grp.get('E', [[], []])[1] if len(grp.get('E', [])) > 1 else []
    if not over_list:
        return None
    mid        = len(over_list) // 2
    main_over  = next((x for x in over_list  if x.get('CE') == 1), over_list[mid])
    main_under = next((x for x in under_list if x.get('CE') == 1),
                      under_list[mid] if mid < len(under_list) else None)
    return {
        'line':  main_over.get('P'),
        'over':  main_over.get('C'),
        'under': main_under.get('C') if main_under else None,
    }

def _fetch_live_bet_lines(bet_ci, tis=None):
    """Busca linhas ao vivo para múltiplos TIs de um jogo 22bet.
    Devolve dict {ti: {line, over, under}}. Uma chamada API, múltiplos mercados.
    """
    if tis is None:
        tis = [_LANC_TI]
    active_tis = [t for t in tis if t and t > 0]
    if not active_tis:
        return {}
    try:
        url1 = (f'{_BET_BASE}/GetGameZip?id={bet_ci}&lng=pt_PT&tzo=1&cfview=0'
                f'&isSubGames=true&GroupEvents=true&countevents=250'
                f'&country=148&fcountry=148&isNewBuilder=true&partner=151&grMode=4')
        resp1 = _bet_api(url1)
        value1 = (resp1 or {}).get('Value') if resp1 else None
        if value1 is None:
            # pre-match CI usado em feed live → Value=null
            logger.debug(f'[22BET] GetGameZip {bet_ci}: Value=null (CI pré-jogo?)')
            return None   # sentinel: distingue "null" de "sem linhas"
        sg = (value1 or {}).get('SG', [])
    except Exception as e:
        logger.warning(f'[22BET] GetGameZip {bet_ci}: {e}')
        return {}
    result = {}
    for ti in active_tis:
        entry = next((e for e in sg if e.get('TI') == ti and not e.get('P')), None)
        if not entry:
            entry = next((e for e in sg if e.get('TI') == ti), None)
        if not entry:
            logger.debug(f'[22BET] TI={ti} não encontrado para CI={bet_ci}')
            continue
        ti_ci = entry.get('N') or entry.get('CI')   # LiveFeed usa N, LineFeed usava CI
        if not ti_ci:
            continue
        try:
            url2 = (f'{_BET_BASE}/GetGameZip?id={ti_ci}&lng=pt&tzo=1&cfview=0'
                    f'&isSubGames=true&GroupEvents=true&countevents=250'
                    f'&country=148&fcountry=148&isNewBuilder=true&partner=151&grMode=4')
            ge   = ((_bet_api(url2) or {}).get('Value') or {}).get('GE', [])
            odds = _parse_ou_odds(ge)
            if odds:
                result[ti] = odds
        except Exception as e:
            logger.warning(f'[22BET] odds TI={ti} CI={ti_ci}: {e}')
    return result

# ── 22bet live CI lookup ─────────────────────────────────────────────────────
# Cache de índice live: {(norm_home, norm_away): live_ci}
_22BET_LIVE_IDX:    dict  = {}
_22BET_LIVE_IDX_TS: float = 0.0
# Liga codes → LIs conhecidos no feed live (adicionados automaticamente ao descobrir)
_22BET_KNOWN_LIS: list = [3007689, 118593, 2252762]   # PPL, Europa League, Conference League
# Cache negativo: games onde já falhámos lookup (evitar re-scan frequente)
_22BET_NO_CI: dict = {}   # {flash_mid: timestamp}

# ── HT snapshot: guarda totais do 1º tempo para corrigir reset do intervalo ──
_HT_TOTALS: dict = {}   # {flash_mid: {'lanc': int, 'faltas': int}}

def _seed_ht_totals():
    """No arranque, semeia _HT_TOTALS a partir dos snapshots guardados em DB
    para jogos que já passaram o intervalo."""
    if not LIVE_DB_PATH.exists(): return
    try:
        con = sqlite3.connect(LIVE_DB_PATH); con.row_factory = sqlite3.Row
        rows = con.execute('''
            SELECT flash_mid,
                   MAX(CASE WHEN minuto_est IN (-1,44,45) THEN lanc_total  END) as lt,
                   MAX(CASE WHEN minuto_est IN (-1,44,45) THEN faltas_total END) as ft
            FROM live_snapshots
            GROUP BY flash_mid
            HAVING lt IS NOT NULL AND ft IS NOT NULL
        ''').fetchall()
        con.close()
        for r in rows:
            fmid = r['flash_mid']
            if fmid not in _HT_TOTALS:
                _HT_TOTALS[fmid] = {'lanc': r['lt'], 'faltas': r['ft']}
                logger.info(f'[HT-SEED] {fmid}: L={r["lt"]} F={r["ft"]}')
    except Exception as e:
        logger.debug(f'[HT-SEED] erro: {e}')

def _norm_bet(s: str) -> str:
    import unicodedata
    return unicodedata.normalize('NFD', (s or '').lower()).encode('ascii','ignore').decode().strip()

def _22bet_league_games(li: int) -> list:
    """Busca jogos live de uma liga (por LI). Filtra only games da liga pedida."""
    url = (f'{_BET_BASE}/GetSportsShortZip?sports=1&champs={li}'
           f'&lng=pt_PT&gr=151&partner=151&virtualSports=true')
    try:
        data = _bet_api(url)
        games = []
        for c in (data.get('Value') or []):
            for l in (c.get('L') or []):
                if l.get('LI') != li:
                    continue   # skip context leagues returned by API
                for g in (l.get('G') or []):
                    if g.get('O1') and g.get('O2') and g.get('I'):
                        games.append({'ci': g['I'], 'o1': g['O1'], 'o2': g['O2'], 'li': li})
        return games
    except Exception as e:
        logger.debug(f'[22BET] league {li} games: {e}')
        return []

def _refresh_22bet_live_idx():
    """Reconstrói índice {(norm_home, norm_away): live_ci} com TTL 60s."""
    global _22BET_LIVE_IDX, _22BET_LIVE_IDX_TS, _22BET_KNOWN_LIS
    import time as _time
    if _time.time() - _22BET_LIVE_IDX_TS < 60:
        return _22BET_LIVE_IDX
    idx = {}
    # 1) Fetch known leagues
    lis_to_check = list(_22BET_KNOWN_LIS)
    # 2) Enrich with live index (up to 10 extra LIs with GC>0)
    try:
        url_idx = f'{_BET_BASE}/GetSportsShortZip?sports=1&lng=pt_PT&gr=151&partner=151&virtualSports=true'
        idx_data = _bet_api(url_idx)
        extra = 0
        for c in (idx_data.get('Value') or []):
            for l in (c.get('L') or []):
                li = l.get('LI')
                if li and l.get('GC', 0) > 0 and li not in lis_to_check:
                    lis_to_check.append(li)
                    extra += 1
                    if extra >= 10:
                        break
            if extra >= 10:
                break
    except Exception as e:
        logger.debug(f'[22BET] live idx index: {e}')
    # 3) Fetch each league individually
    for li in lis_to_check:
        for g in _22bet_league_games(li):
            h = _norm_bet(g['o1'])
            a = _norm_bet(g['o2'])
            if h and a:
                idx[(h, a)] = g['ci']
                # Guardar LI descoberto para proximas calls
                if li not in _22BET_KNOWN_LIS:
                    _22BET_KNOWN_LIS.append(li)
    _22BET_LIVE_IDX    = idx
    _22BET_LIVE_IDX_TS = _time.time()
    logger.info(f'[22BET] live idx: {len(idx)} jogos em {len(lis_to_check)} ligas')
    return _22BET_LIVE_IDX

def _find_22bet_live_ci(home: str, away: str) -> int | None:
    """Procura o live CI de um jogo por nomes das equipas. Devolve CI ou None."""
    idx = _refresh_22bet_live_idx()
    h = _norm_bet(home)
    a = _norm_bet(away)
    ci = idx.get((h, a))
    if not ci:
        # fuzzy word match
        hw = {w for w in h.split() if len(w) > 3}
        aw = {w for w in a.split() if len(w) > 3}
        for (sh, sa), eid in idx.items():
            if hw & {w for w in sh.split() if len(w) > 3} and \
               aw & {w for w in sa.split() if len(w) > 3}:
                ci = eid
                break
    return ci

_LANC_ALERT_THRESHOLD  = 4.0   # delta mínimo lançamentos para acionar Telegram
_FALT_ALERT_THRESHOLD  = 2.5   # delta mínimo faltas para acionar Telegram

def collect_live_once():
    """Job APScheduler: recolhe stats de jogos activos a cada 10 min.
    Inclui jogos até 2h antes do kickoff (pré-jogo) para mostrar linhas 22bet antecipadamente."""
    now = datetime.now(timezone.utc)
    con = init_live_db()
    window_start = (now - timedelta(minutes=130)).isoformat()
    window_end   = (now + timedelta(hours=3)).isoformat()   # pré-jogo até 3h antes
    # Limpar jogos terminados (kickoff > 130 min passados) — sempre, antes do return
    con.execute("UPDATE live_games SET status='done' WHERE status='live' AND kickoff <= ?",
                (window_start,))
    con.commit()

    rows = con.execute('''SELECT flash_mid,league,home,away,kickoff,lanc_baseline,faltas_baseline,bet_ci,referee,statscore_id
        FROM live_games WHERE status IN ('pending','live')
        AND kickoff <= ? AND kickoff >= ?''',
        (window_end, window_start)).fetchall()
    if not rows:
        con.close(); return
    logger.info(f'[LIVE] {len(rows)} jogos na janela (incl. pré-jogo)')
    for flash_mid, liga, home, away, ko_iso, lb, fb, bet_ci, ref_name, statscore_id in rows:
        try:
            ko_dt      = datetime.fromisoformat(ko_iso)
            secs_to    = (ko_dt - now).total_seconds()
            pre_game   = secs_to > 0
            elapsed_m  = 0 if pre_game else int(-secs_to / 60)
            # heurística HT: 46-62 min de relógio = intervalo
            is_ht      = 46 <= elapsed_m <= 62
            if pre_game:
                minuto = 0
            elif is_ht:
                minuto = 45   # congela previsões no 45'
            elif elapsed_m > 62:
                # 2ª parte — desconta ~17 min de intervalo
                minuto = max(45, min(elapsed_m - 17, 95))
            else:
                minuto = max(0, min(elapsed_m, 45))
        except:
            pre_game = False; is_ht = False; minuto = 45

        if not pre_game:
            con.execute("UPDATE live_games SET status='live' WHERE flash_mid=?", (flash_mid,))

        # ── Statscore stats (tempo real, prioridade sobre Flashscore) ─────────
        sc_lc = sc_lf = sc_fc = sc_ff = None
        if not pre_game and statscore_id:
            sc = _fetch_statscore(statscore_id)
            if sc:
                sc_lc = sc.get('sc_lanc_casa')
                sc_lf = sc.get('sc_lanc_fora')
                sc_fc = sc.get('sc_faltas_casa')
                sc_ff = sc.get('sc_faltas_fora')
                # Statscore conhece HT com certeza
                if sc.get('sc_status') == 'halftime':
                    is_ht = True; minuto = 45
                elif sc.get('sc_minute'):
                    minuto = sc['sc_minute']; is_ht = False
                logger.info(f'[SC] {home}: L={sc_lc}-{sc_lf} F={sc_fc}-{sc_ff} status={sc.get("sc_status")} min={sc.get("sc_minute")}')

        # ── Flashscore stats (fallback se sem Statscore/SofaScore) ───────────────
        lc = lf = fc = ff = lt = ft = le = fe = None
        if not pre_game:
            if sc_lc is not None or sc_lf is not None:
                # 1ª prioridade: Statscore (tempo real)
                lc, lf = sc_lc, sc_lf
                fc, ff = sc_fc, sc_ff
            else:
                # 2ª prioridade: SofaScore (near real-time, via flash_mid ou nome)
                sofa = _fetch_sofascore(flash_mid, home=home, away=away)
                if sofa:
                    lc, lf = sofa.get('sofa_lc'), sofa.get('sofa_lf')
                    fc, ff = sofa.get('sofa_fc'), sofa.get('sofa_ff')
                else:
                    # 3ª prioridade: Ninja/Flashscore (fallback com delay)
                    stats = _fetch_ninja(flash_mid)
                    if not stats: continue
                    lc,lf = stats.get('lanc_casa'), stats.get('lanc_fora')
                    fc,ff = stats.get('faltas_casa'), stats.get('faltas_fora')
            lt = (lc or 0)+(lf or 0) if lc is not None else None
            ft = (fc or 0)+(ff or 0) if fc is not None else None

            # ── HT snapshot: guardar totais no intervalo ──────────────────────
            if is_ht and lt is not None and ft is not None:
                if flash_mid not in _HT_TOTALS or _HT_TOTALS[flash_mid]['faltas'] < ft:
                    _HT_TOTALS[flash_mid] = {'lanc': lt, 'faltas': ft}
                    logger.info(f'[HT] {flash_mid}: snapshot L={lt} F={ft}')

            # ── HT reset fix: somar 1º tempo ao contador da 2ª parte ──────────
            if minuto > 45 and flash_mid in _HT_TOTALS:
                ht = _HT_TOTALS[flash_mid]
                if lt is not None and lt < ht['lanc']:
                    logger.debug(f'[HT-FIX] {flash_mid}: lt {lt}+{ht["lanc"]}')
                    lt += ht['lanc']
                if ft is not None and ft < ht['faltas']:
                    logger.debug(f'[HT-FIX] {flash_mid}: ft {ft}+{ht["faltas"]}')
                    ft += ht['faltas']

            le = round(lt/minuto*90,1) if (lt and minuto>0) else None
            fe = round(ft/minuto*90,1) if (ft and minuto>0) else None

        lb2 = lb or _LANC_BASE.get(liga, 37)
        fb2 = _faltas_pregame_pred(home, away, liga, referee=ref_name) or fb or _FALTAS_BASE.get(liga, 27)
        lp = _pred(minuto, lt or 0, lb2)
        fp = _pred(minuto, ft or 0, fb2)

        # ── 22bet linha ao vivo ───────────────────────────────────────────────
        # Lookup bet_ci em lancamentos_lines.db se ainda não temos
        if not bet_ci:
            try:
                ldb = sqlite3.connect(DB_PATH)
                r = ldb.execute(
                    '''SELECT match_id FROM lancamentos_lines
                       WHERE league=? AND home LIKE ? AND away LIKE ? LIMIT 1''',
                    (liga, f'{home[:10]}%', f'{away[:10]}%')
                ).fetchone()
                ldb.close()
                if r:
                    bet_ci = r[0]
                    con.execute('UPDATE live_games SET bet_ci=? WHERE flash_mid=?', (bet_ci, flash_mid))
                    logger.info(f'[22BET] bet_ci={bet_ci} encontrado para {home} vs {away}')
            except Exception as e:
                logger.warning(f'[22BET] lookup bet_ci: {e}')

        live_line = live_over = live_under = live_signal = None
        fl_line   = fl_over   = fl_under   = fl_signal  = None
        if bet_ci:
            import time as _t
            bets = _fetch_live_bet_lines(bet_ci, tis=[_LANC_TI, _FALTAS_TI])
            # bets=None → CI pré-jogo usado em feed live (Value=null)
            # Procurar live CI automaticamente (throttle: 1x por 5min por jogo)
            if bets is None and not pre_game:
                _last = _22BET_NO_CI.get(flash_mid, 0)
                if _t.time() - _last > 300:
                    new_ci = _find_22bet_live_ci(home, away)
                    if new_ci and str(new_ci) != str(bet_ci):
                        logger.info(f'[22BET] live CI encontrado: {new_ci} (era {bet_ci}) para {home} vs {away}')
                        bet_ci = str(new_ci)
                        con.execute('UPDATE live_games SET bet_ci=? WHERE flash_mid=?', (bet_ci, flash_mid))
                        con.commit()
                        bets = _fetch_live_bet_lines(bet_ci, tis=[_LANC_TI, _FALTAS_TI])
                    else:
                        logger.warning(f'[22BET] live CI não encontrado para {home} vs {away}')
                        _22BET_NO_CI[flash_mid] = _t.time()
                bets = bets or {}
            bets = bets or {}
            lanc_bet   = bets.get(_LANC_TI, {})
            faltas_bet = bets.get(_FALTAS_TI, {}) if _FALTAS_TI else {}
            live_line  = lanc_bet.get('line')
            live_over  = lanc_bet.get('over')
            live_under = lanc_bet.get('under')
            fl_line    = faltas_bet.get('line')
            fl_over    = faltas_bet.get('over')
            fl_under   = faltas_bet.get('under')

        # ── Sinal + Alerta Telegram — Lançamentos ────────────────────────────
        # Cache guarda só a direção (OVER/UNDER) — não o delta exacto.
        # Assim não re-alerta a cada variação mínima de linha/stats.
        if live_line and lp:
            delta = round(lp - live_line, 1)
            _ck = (flash_mid, 'lanc')
            if delta >= _LANC_ALERT_THRESHOLD:
                live_signal = f'OVER+{delta}'
                if _alert_cache.get(_ck) != 'OVER':
                    _alert_cache[_ck] = 'OVER'
                    send_telegram(
                        f'📲 <b>LIVE ALERTA — Lançamentos OVER</b>\n'
                        f'{home} vs {away} ({liga}) ~{minuto}\'\n'
                        f'Pred: <b>{lp}</b> | Linha: <b>{live_line}</b> | Δ <b>+{delta}</b>\n'
                        f'Odds Over: {live_over}'
                    )
            elif delta <= -_LANC_ALERT_THRESHOLD:
                live_signal = f'UNDER{delta}'
                if _alert_cache.get(_ck) != 'UNDER':
                    _alert_cache[_ck] = 'UNDER'
                    send_telegram(
                        f'📲 <b>LIVE ALERTA — Lançamentos UNDER</b>\n'
                        f'{home} vs {away} ({liga}) ~{minuto}\'\n'
                        f'Pred: <b>{lp}</b> | Linha: <b>{live_line}</b> | Δ <b>{delta}</b>\n'
                        f'Odds Under: {live_under}'
                    )
            else:
                live_signal = f'NEU{delta:+.1f}'
                # sinal neutro — limpa cache para re-alertar se voltar a cruzar
                _alert_cache.pop(_ck, None)

        # ── Sinal + Alerta Telegram — Faltas ─────────────────────────────────
        if fl_line and fp:
            fdelta = round(fp - fl_line, 1)
            _ck = (flash_mid, 'falt')
            if fdelta >= _FALT_ALERT_THRESHOLD:
                fl_signal = f'OVER+{fdelta}'
                if _alert_cache.get(_ck) != 'OVER':
                    _alert_cache[_ck] = 'OVER'
                    send_telegram(
                        f'📲 <b>LIVE ALERTA — Faltas OVER</b>\n'
                        f'{home} vs {away} ({liga}) ~{minuto}\'\n'
                        f'Pred: <b>{fp}</b> | Linha: <b>{fl_line}</b> | Δ <b>+{fdelta}</b>\n'
                        f'Odds Over: {fl_over}'
                    )
            elif fdelta <= -_FALT_ALERT_THRESHOLD:
                fl_signal = f'UNDER{fdelta}'
                if _alert_cache.get(_ck) != 'UNDER':
                    _alert_cache[_ck] = 'UNDER'
                    send_telegram(
                        f'📲 <b>LIVE ALERTA — Faltas UNDER</b>\n'
                        f'{home} vs {away} ({liga}) ~{minuto}\'\n'
                        f'Pred: <b>{fp}</b> | Linha: <b>{fl_line}</b> | Δ <b>{fdelta}</b>\n'
                        f'Odds Under: {fl_under}'
                    )
            else:
                fl_signal = f'NEU{fdelta:+.1f}'
                _alert_cache.pop(_ck, None)

        # minuto_est: -1 = HT, 0-95 = minuto real
        minuto_snap = -1 if is_ht else minuto
        con.execute('''INSERT INTO live_snapshots
            (flash_mid,league,home,away,kickoff,minuto_est,
             lanc_casa,lanc_fora,lanc_total,faltas_casa,faltas_fora,faltas_total,
             lanc_extrap,faltas_extrap,lanc_pred,faltas_pred,
             live_line,live_over_odds,live_under_odds,live_signal,
             fl_live_line,fl_live_over_odds,fl_live_under_odds,fl_live_signal,
             sc_lanc_casa,sc_lanc_fora,sc_faltas_casa,sc_faltas_fora,
             captured_at)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
            (flash_mid,liga,home,away,ko_iso,minuto_snap,lc,lf,lt,fc,ff,ft,le,fe,lp,fp,
             live_line,live_over,live_under,live_signal,
             fl_line,fl_over,fl_under,fl_signal,
             sc_lc,sc_lf,sc_fc,sc_ff,
             now.isoformat()))
        logger.info(f'[LIVE] {home} vs {away} ~{minuto}\' '
                    f'Lanc={lt}→{lp} L22={live_line}/{live_signal} '
                    f'Falt={ft}→{fp} F22={fl_line}/{fl_signal}')
    con.commit(); con.close()

# Job live — adicionado depois de collect_live_once estar definida (v2)
scheduler.add_job(collect_live_once, 'interval', seconds=30, id='live_collect')

# ── Auto-loader diário de jogos (Flashscore) ──────────────────────────────────
_AUTO_LOADER_LEAGUES = [
    {'code': 'PPL', 'url': 'https://www.flashscore.pt/futebol/portugal/liga-portugal-betclic/'},
    {'code': 'ESP', 'url': 'https://www.flashscore.com/football/spain/laliga/'},
    {'code': 'BRA', 'url': 'https://www.flashscore.com/football/brazil/serie-a/'},
]

def auto_load_daily_games():
    """Detecta automaticamente os jogos do dia via Flashscore e popula live_games DB.
    Corre no arranque e às 09:00 UTC. Não sobrescreve registos existentes (INSERT OR IGNORE).
    AB=3 = por começar, AB=1 = live, AB=2 = terminado (ignorado).
    """
    import time as _t
    GBOT_UA = 'Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)'
    now         = _t.time()
    today_start = now - 7200     # até 2h atrás (jogo já começou mas app reiniciou)
    today_end   = now + 86400    # próximas 24h

    # Cria a DB se não existir (após redeploy Railway o ficheiro é apagado)
    con = init_live_db()
    total = 0

    for league in _AUTO_LOADER_LEAGUES:
        code = league['code']
        url  = league['url']
        try:
            html = _arb_fetch(url, ua=GBOT_UA)
            if not html or len(html) < 5000:
                logger.warning(f'[AUTO] {code}: feed insuficiente ({len(html or "")} bytes)')
                continue

            found = 0
            for rec in _FS_MATCH_RE.finditer(html):
                mid    = rec.group(1)
                fields = {kv.group(1): kv.group(2).strip()
                          for kv in _FS_FIELD_RE.finditer(rec.group(2))}

                if fields.get('AB', '') == '2':
                    continue  # terminado — ignorar

                ts = int(fields.get('AD', '0') or '0')
                if not ts or not (today_start <= ts <= today_end):
                    continue  # fora da janela do dia

                home = _arb_strip_tags(fields.get('CX', '')).strip()
                away = _arb_strip_tags(fields.get('AF', '')).strip()
                if not home or not away or len(home) < 2 or len(away) < 2:
                    continue

                kickoff = datetime.fromtimestamp(ts, tz=timezone.utc).isoformat()

                # Baselines: médias das equipas ou default da liga
                hn = _norm_nome(home); an = _norm_nome(away)
                hd = _MEDIAS.get(hn, {})
                ad_m = _MEDIAS.get(an, {})
                # Fuzzy match por palavras se não encontrado exact
                if not hd:
                    hw = {w for w in hn.split() if len(w) > 3}
                    for k, v in _MEDIAS.items():
                        if v.get('liga') == code and hw & {w for w in k.split() if len(w) > 3}:
                            hd = v; break
                if not ad_m:
                    aw = {w for w in an.split() if len(w) > 3}
                    for k, v in _MEDIAS.items():
                        if v.get('liga') == code and aw & {w for w in k.split() if len(w) > 3}:
                            ad_m = v; break

                lh = hd.get('lanc_media') or 0;  la = ad_m.get('lanc_media') or 0
                fh = hd.get('faltas_media') or 0; fa = ad_m.get('faltas_media') or 0
                lb = round((lh + la) / 2, 1) if (lh or la) else _LANC_BASE.get(code, 37.0)
                fb = round((fh + fa) / 2, 1) if (fh or fa) else _FALTAS_BASE.get(code, 27.0)

                cur = con.execute(
                    '''INSERT OR IGNORE INTO live_games
                       (flash_mid,league,home,away,kickoff,
                        lanc_baseline,faltas_baseline,status,bet_ci,referee,statscore_id)
                       VALUES (?,?,?,?,?,?,?,'pending',NULL,NULL,NULL)''',
                    (mid, code, home, away, kickoff, lb, fb))
                if cur.rowcount:
                    logger.info(f'[AUTO] {code}: {home} vs {away} '
                                f'@ {kickoff[:16]}Z mid={mid} L={lb} F={fb}')
                    found += 1; total += 1

            con.commit()
            logger.info(f'[AUTO] {code}: {found} jogos novos adicionados')
        except Exception as e:
            logger.warning(f'[AUTO] {code}: erro — {e}')

    con.close()
    logger.info(f'[AUTO] total {total} jogos novos detectados')

# ── Live DB ───────────────────────────────────────────────────────────────────
def init_live_db():
    con = sqlite3.connect(LIVE_DB_PATH)
    con.execute('''CREATE TABLE IF NOT EXISTS live_games (
        flash_mid TEXT PRIMARY KEY, league TEXT, home TEXT, away TEXT,
        kickoff TEXT, lanc_baseline REAL, faltas_baseline REAL, status TEXT,
        bet_ci TEXT, referee TEXT, statscore_id TEXT)''')
    con.execute('''CREATE TABLE IF NOT EXISTS live_snapshots (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        flash_mid TEXT, league TEXT, home TEXT, away TEXT, kickoff TEXT,
        minuto_est INTEGER, lanc_casa INTEGER, lanc_fora INTEGER, lanc_total INTEGER,
        faltas_casa INTEGER, faltas_fora INTEGER, faltas_total INTEGER,
        lanc_extrap REAL, faltas_extrap REAL, lanc_pred REAL, faltas_pred REAL,
        live_line REAL, live_over_odds REAL, live_under_odds REAL, live_signal TEXT,
        fl_live_line REAL, fl_live_over_odds REAL, fl_live_under_odds REAL, fl_live_signal TEXT,
        captured_at TEXT)''')
    # Migração: adiciona colunas em DBs existentes (ignora erro se já existem)
    for sql in [
        'ALTER TABLE live_games ADD COLUMN bet_ci TEXT',
        'ALTER TABLE live_games ADD COLUMN referee TEXT',
        'ALTER TABLE live_snapshots ADD COLUMN live_line REAL',
        'ALTER TABLE live_snapshots ADD COLUMN live_over_odds REAL',
        'ALTER TABLE live_snapshots ADD COLUMN live_under_odds REAL',
        'ALTER TABLE live_snapshots ADD COLUMN live_signal TEXT',
        'ALTER TABLE live_snapshots ADD COLUMN fl_live_line REAL',
        'ALTER TABLE live_snapshots ADD COLUMN fl_live_over_odds REAL',
        'ALTER TABLE live_snapshots ADD COLUMN fl_live_under_odds REAL',
        'ALTER TABLE live_snapshots ADD COLUMN fl_live_signal TEXT',
        'ALTER TABLE live_games ADD COLUMN statscore_id TEXT',
        'ALTER TABLE live_snapshots ADD COLUMN sc_lanc_casa INTEGER',
        'ALTER TABLE live_snapshots ADD COLUMN sc_lanc_fora INTEGER',
        'ALTER TABLE live_snapshots ADD COLUMN sc_faltas_casa INTEGER',
        'ALTER TABLE live_snapshots ADD COLUMN sc_faltas_fora INTEGER',
    ]:
        try: con.execute(sql)
        except: pass
    con.commit()
    return con

def ler_live_jogos():
    if not LIVE_DB_PATH.exists(): return []
    con = sqlite3.connect(LIVE_DB_PATH); con.row_factory = sqlite3.Row
    rows = con.execute('SELECT * FROM live_games ORDER BY kickoff').fetchall()
    con.close(); return [dict(r) for r in rows]

def ler_live_snapshots(flash_mid=None, limit=200):
    if not LIVE_DB_PATH.exists(): return []
    con = sqlite3.connect(LIVE_DB_PATH); con.row_factory = sqlite3.Row
    if flash_mid:
        rows = con.execute('SELECT * FROM live_snapshots WHERE flash_mid=? ORDER BY minuto_est', (flash_mid,)).fetchall()
    else:
        rows = con.execute('SELECT * FROM live_snapshots ORDER BY captured_at DESC LIMIT ?', (limit,)).fetchall()
    con.close(); return [dict(r) for r in rows]

def ler_latest_snaps():
    """Devolve dict flash_mid → snapshot mais recente por jogo."""
    if not LIVE_DB_PATH.exists(): return {}
    con = sqlite3.connect(LIVE_DB_PATH); con.row_factory = sqlite3.Row
    rows = con.execute('''
        SELECT s.* FROM live_snapshots s
        INNER JOIN (
            SELECT flash_mid, MAX(captured_at) AS max_at FROM live_snapshots GROUP BY flash_mid
        ) t ON s.flash_mid=t.flash_mid AND s.captured_at=t.max_at
    ''').fetchall()
    con.close()
    return {r['flash_mid']: dict(r) for r in rows}

_LIVE_CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',sans-serif;background:#0e1b2e;color:#e0e0e0;padding:16px;font-size:13px}
h1{font-size:1rem;font-weight:600;color:#fff;margin-bottom:4px}
.sub{font-size:.72rem;color:#555;margin-bottom:6px}
.refresh-bar{font-size:.68rem;color:#444;margin-bottom:16px}
.tabs{display:flex;gap:6px;margin-bottom:14px}
.tab{padding:5px 14px;border-radius:20px;font-size:.74rem;font-weight:600;text-decoration:none;border:1px solid #2a2d3a;color:#555}
.tab.active{color:#fff;border-color:currentColor}
.tab-lanc.active{color:#64b5f6;border-color:#64b5f6;background:#0d1a2a}
.tab-falt.active{color:#9b6bd5;border-color:#9b6bd5;background:#160d2a}
.legend{display:flex;gap:14px;font-size:.68rem;color:#555;margin-bottom:14px;flex-wrap:wrap}
.legend span{display:flex;align-items:center;gap:5px}
.ldot{width:7px;height:7px;border-radius:50%;display:inline-block}
.alert-banner{background:#0d2016;border:1px solid #1a5c2a;border-radius:6px;padding:7px 11px;margin-bottom:10px;font-size:.74rem;color:#4caf50}
.section-h{font-size:.72rem;color:#3a3e52;text-transform:uppercase;letter-spacing:.07em;margin:18px 0 8px;border-bottom:1px solid #182338;padding-bottom:4px}
.game-card{background:#132033;border:1px solid #1e2d44;border-radius:8px;margin-bottom:10px;overflow:hidden}
.game-header{display:flex;align-items:center;justify-content:space-between;background:#172236;padding:7px 11px;border-bottom:1px solid #1e2d44}
.game-teams{font-weight:600;font-size:.82rem;color:#fff}
.liga-tag{color:#555;font-weight:400;font-size:.7rem;margin-left:6px}
.badge-live{background:#0f2f0f;color:#4caf50;font-size:.64rem;font-weight:700;padding:2px 6px;border-radius:4px;border:1px solid #1b5e1b}
.badge-min{color:#64b5f6;font-size:.72rem;font-weight:600}
.badge-ht{background:#3a2e00;color:#ffc107;font-size:.64rem;font-weight:700;padding:2px 7px;border-radius:4px;border:1px solid #7a6200}
.progress-wrap{height:3px;background:#1e2d44;width:100%}
.progress-bar{height:3px;background:linear-gradient(90deg,#1565c0,#64b5f6);transition:width .4s}
.progress-bar.ht{background:#ffc107;width:50%!important}
.game-body{display:grid;grid-template-columns:1fr 1fr 1fr}
.section{padding:9px 12px}
.section-label{font-size:.63rem;font-weight:700;letter-spacing:.09em;text-transform:uppercase;margin-bottom:7px}
.s-flash{border-right:1px solid #1e2d44}.s-flash .section-label{color:#7c9bc0}
.s-pred{border-right:1px solid #1e2d44}.s-pred .section-label{color:#9b6bd5}
.s-bet{background:#141720}.s-bet .section-label{color:#c8843a}
.s-bet-falt{background:#141720}.s-bet-falt .section-label{color:#88c868}
.stat-row{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:4px}
.stat-name{color:#666;font-size:.71rem}
.stat-val{font-size:.8rem;font-weight:600;color:#ccc;font-variant-numeric:tabular-nums}
.pred-val{color:#64b5f6;font-size:.9rem;font-weight:700}
.pred-val-falt{color:#9b6bd5;font-size:.9rem;font-weight:700}
.extrap-val{color:#7c9bc0;font-size:.8rem}
.extrap-val-falt{color:#9b6bd5;font-size:.8rem;opacity:.7}
.muted-val{color:#444;font-size:.75rem}
.divt{border-top:1px solid #1e2d44;margin:5px 0}
.bet-line-main{font-size:1.15rem;font-weight:700;color:#e8a043;margin-bottom:2px}
.bet-line-falt{font-size:1.15rem;font-weight:700;color:#88c868;margin-bottom:2px}
.bet-odds{display:flex;gap:10px;margin-bottom:7px;font-size:.72rem;color:#888}
.bet-odds b{color:#c8843a;font-weight:600}
.bet-odds-falt b{color:#88c868;font-weight:600}
.signal-over{background:#0d2e10;color:#4caf50;border:1px solid #1b5e1b;border-radius:5px;font-size:.78rem;font-weight:700;padding:4px 10px;display:inline-block}
.signal-under{background:#2e0d0d;color:#ef5350;border:1px solid #5e1b1b;border-radius:5px;font-size:.78rem;font-weight:700;padding:4px 10px;display:inline-block}
.signal-neu{background:#1e2d44;color:#666;border-radius:5px;font-size:.72rem;font-weight:700;padding:4px 10px;display:inline-block}
.signal-delta{font-size:.67rem;color:#555;margin-top:3px}
.bet-waiting{color:#333;font-size:.75rem;font-style:italic;margin-top:4px}
.pending-row{display:flex;justify-content:space-between;align-items:center;padding:6px 10px;border-bottom:1px solid #182338;font-size:.78rem;color:#444}
.sig{font-size:.68rem;font-weight:700;padding:2px 6px;border-radius:3px}
.sig-over{background:#0d2e10;color:#4caf50;border:1px solid #1b5e1b}
.sig-under{background:#2e0d0d;color:#ef5350;border:1px solid #5e1b1b}
.sig-neu{background:#1e2d44;color:#666}
.pending-row:last-child{border-bottom:none}
.pending-box{background:#0e1b2e;border:1px solid #1a1d2a;border-radius:6px;overflow:hidden;margin-bottom:10px}
.done-row{display:flex;justify-content:space-between;align-items:center;padding:5px 10px;border-bottom:1px solid #182338;font-size:.75rem;color:#3a3e52}
.done-box{background:#0e1b2e;border:1px solid #1a1d2a;border-radius:6px;overflow:hidden}
.empty{color:#333;font-size:.82rem;text-align:center;padding:30px}
"""

LIVE_TEMPLATE = """<!DOCTYPE html>
<html lang="pt"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Live — Lançamentos</title>
<style>{{ css }}</style>
</head><body>
<div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:10px">
  <div>
    <h1 style="margin:0">⚽ Live — Lançamentos</h1>
    <div class="sub">Flashscore Ninja API + 22bet TI=55 · atualizado a cada 2 min</div>
  </div>
</div>
<div class="tabs">
  <a class="tab tab-lanc active" href="/live">Lançamentos</a>
  <a class="tab tab-falt" href="/live/faltas">Faltas</a>
  <a class="tab" href="/live/historico" style="color:#ffc107;border-color:#ffc107">📊 Histórico</a>
</div>
<div class="refresh-bar">Auto-refresh em <span id="cnt">30</span>s</div>

<div class="legend">
  <span><div class="ldot" style="background:#7c9bc0"></div> Flashscore (real)</span>
  <span><div class="ldot" style="background:#64b5f6"></div> Previsão modelo</span>
  <span><div class="ldot" style="background:#c8843a"></div> 22bet linha viva (TI=55)</span>
</div>

{% if alertas %}{% for a in alertas %}
<div class="alert-banner">📲 {{ a }}</div>
{% endfor %}{% endif %}

{% if live_jogos %}
<div class="section-h">Ao vivo ({{ live_jogos|length }})</div>
{% for j in live_jogos %}{% set s = snaps_mapa.get(j.flash_mid) %}
{% set is_ht = s and s.minuto_est == -1 %}
{% set min_show = 45 if is_ht else (s.minuto_est if s else 0) %}
{% set pct = [[(min_show / 90 * 100)|int, 0]|max, 100]|min %}
<div class="game-card">
  <div class="game-header">
    <div class="game-teams">{{ j.home }} vs {{ j.away }}<span class="liga-tag">{{ j.league }}</span></div>
    <div style="display:flex;align-items:center;gap:8px">
      {% if is_ht %}<span class="badge-ht">HT</span>
      {% elif s %}<span class="badge-min">{{ s.minuto_est }}'</span>{% endif %}
      <span class="badge-live">LIVE</span>
    </div>
  </div>
  <div class="progress-wrap"><div class="progress-bar{% if is_ht %} ht{% endif %}" style="width:{{ pct }}%"></div></div>
  <div class="game-body">
    <div class="section s-flash">
      {% if s and s.sc_lanc_casa is not none %}
      <div class="section-label" style="color:#4db6ac">Statscore <span style="font-size:.6rem;opacity:.6">⚡ real</span></div>
      <div class="stat-row"><span class="stat-name">Lanç. laterais</span><span class="stat-val">{{ (s.sc_lanc_casa or 0) + (s.sc_lanc_fora or 0) }} <span style="font-size:.68rem;color:#666">({{ s.sc_lanc_casa }}-{{ s.sc_lanc_fora }})</span></span></div>
      <div class="stat-row"><span class="stat-name">Faltas</span><span class="stat-val">{{ (s.sc_faltas_casa or 0) + (s.sc_faltas_fora or 0) }} <span style="font-size:.68rem;color:#666">({{ s.sc_faltas_casa }}-{{ s.sc_faltas_fora }})</span></span></div>
      {% else %}
      <div class="section-label">Flashscore</div>
      <div class="stat-row"><span class="stat-name">Lanç. total</span><span class="stat-val">{{ s.lanc_total if s and s.lanc_total is not none else '—' }}</span></div>
      <div class="stat-row"><span class="stat-name">Casa / Fora</span><span class="stat-val" style="font-size:.72rem">{{ (s.lanc_casa|string+' / '+s.lanc_fora|string) if s and s.lanc_casa is not none else '—' }}</span></div>
      {% endif %}
    </div>
    <div class="section s-pred">
      <div class="section-label">Previsão</div>
      <div class="stat-row"><span class="stat-name">Pred</span><span class="pred-val">{{ '%.1f'|format(s.lanc_pred) if s and s.lanc_pred else '—' }}</span></div>
      <div class="stat-row"><span class="stat-name">Extrapol.</span><span class="extrap-val">{{ '%.1f'|format(s.lanc_extrap) if s and s.lanc_extrap else '—' }}</span></div>
      <div class="stat-row"><span class="stat-name">Base</span><span class="muted-val">{{ j.lanc_baseline }}</span></div>
    </div>
    <div class="section s-bet">
      <div class="section-label">22bet Linha</div>
      {% if s and s.live_line %}
      <div class="bet-line-main">{{ s.live_line }}</div>
      <div class="bet-odds">
        <span>O <b>{{ s.live_over_odds or '—' }}</b></span>
        <span>U <b>{{ s.live_under_odds or '—' }}</b></span>
      </div>
      {% if s.live_signal and s.live_signal.startswith('OVER') %}<div class="signal-over">▲ OVER</div>
      {% elif s.live_signal and s.live_signal.startswith('UNDER') %}<div class="signal-under">▼ UNDER</div>
      {% else %}<div class="signal-neu">— SEM SINAL</div>{% endif %}
      <div class="signal-delta">Pred {{ '%.1f'|format(s.lanc_pred) if s.lanc_pred else '?' }} vs {{ s.live_line }}{% if s.live_signal %} Δ {{ s.live_signal[4:] if s.live_signal.startswith('OVER') else s.live_signal[5:] if s.live_signal.startswith('UNDER') else s.live_signal[3:] }}{% endif %}</div>
      {% else %}
      <div class="bet-waiting">A aguardar...{% if not j.bet_ci %}<br><span style="font-size:.65rem;color:#2a2d3a">bet_ci pendente</span>{% endif %}</div>
      {% endif %}
    </div>
  </div>
</div>
{% endfor %}{% endif %}

{% if pending_jogos %}
<div class="section-h">A aguardar ({{ pending_jogos|length }})</div>
<div class="pending-box">
{% for j in pending_jogos %}{% set s = snaps_mapa.get(j.flash_mid) %}
<div class="pending-row">
  <span>{{ j.home }} vs {{ j.away }} <span style="color:#555">{{ j.league }}</span></span>
  <span style="display:flex;gap:12px;align-items:center">
    {% if s %}
      <span>Pred: <b style="color:#64b5f6">{{ '%.1f'|format(s.lanc_pred) if s.lanc_pred else '—' }}</b></span>
      {% if s.live_line %}<span>Linha 22bet: <b style="color:#c8843a">{{ s.live_line }}</b></span>
        {% if s.live_over_odds %}<span style="color:#888">O:{{ s.live_over_odds }} U:{{ s.live_under_odds }}</span>{% endif %}
        {% if s.live_signal %}<span class="sig sig-{{ 'over' if 'OVER' in s.live_signal else ('under' if 'UNDER' in s.live_signal else 'neu') }}">{{ s.live_signal }}</span>{% endif %}
      {% endif %}
    {% else %}
      <span style="color:#555">base {{ j.lanc_baseline }}</span>
    {% endif %}
    <span style="color:#555">{{ j.kickoff[11:16] if j.kickoff else '—' }} UTC</span>
  </span>
</div>{% endfor %}
</div>{% endif %}

{% if done_jogos %}
<div class="section-h">Terminados ({{ done_jogos|length }})</div>
<div class="done-box">
{% for j in done_jogos %}{% set s = snaps_mapa.get(j.flash_mid) %}
<div class="done-row">
  <span>{{ j.home }} vs {{ j.away }} <span style="color:#2a2d3a">{{ j.league }}</span></span>
  <span style="display:flex;gap:10px">{% if s %}
    <span>Lanç: <b style="color:#64b5f6">{{ '%.1f'|format(s.lanc_pred) if s.lanc_pred else '—' }}</b></span>
    {% if s.live_line %}<span>Linha: <b style="color:#c8843a">{{ s.live_line }}</b></span>{% endif %}
  {% endif %}</span>
</div>{% endfor %}
</div>{% endif %}

{% if not live_jogos and not pending_jogos and not done_jogos %}
<div class="empty">Sem jogos registados. Corre setup_weekend.py antes do fim de semana.</div>
{% endif %}
<script>
let c=30;const el=document.getElementById('cnt');
setInterval(()=>{c--;if(c<=0)location.reload();else el.textContent=c},1000);
</script>
</body></html>"""

LIVE_FALTAS_TEMPLATE = """<!DOCTYPE html>
<html lang="pt"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Live — Faltas</title>
<style>{{ css }}</style>
</head><body>
<div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:10px">
  <div>
    <h1 style="margin:0">🟣 Live — Faltas</h1>
    <div class="sub">Flashscore Ninja API · modelo real (reg + árbitro) · atualizado a cada 2 min</div>
  </div>
  <div style="display:flex;align-items:center;gap:8px">
    <input type="file" id="modelFileInput" accept=".json" style="display:none" onchange="uploadFaltasModel(this)">
    <button id="btnUpdateModel" onclick="document.getElementById('modelFileInput').click()"
      style="background:#2a6b4f;color:#fff;border:none;padding:8px 14px;border-radius:6px;cursor:pointer;font-size:.8rem;white-space:nowrap">
      🔄 Atualizar dados faltas
    </button>
  </div>
</div>
<div id="updateModelMsg" style="display:none;margin:6px 0;font-size:.8rem;padding:6px 10px;border-radius:4px"></div>
<div class="tabs">
  <a class="tab tab-lanc" href="/live">Lançamentos</a>
  <a class="tab tab-falt active" href="/live/faltas">Faltas</a>
  <a class="tab" href="/live/historico" style="color:#ffc107;border-color:#ffc107">📊 Histórico</a>
</div>
<div class="refresh-bar">Auto-refresh em <span id="cnt">30</span>s</div>

<div class="legend">
  <span><div class="ldot" style="background:#7c9bc0"></div> Flashscore (real)</span>
  <span><div class="ldot" style="background:#9b6bd5"></div> Previsão modelo + árbitro</span>
  <span><div class="ldot" style="background:#88c868"></div> 22bet linha faltas</span>
</div>

{% if alertas %}{% for a in alertas %}
<div class="alert-banner">📲 {{ a }}</div>
{% endfor %}{% endif %}

{% if live_jogos %}
<div class="section-h">Ao vivo ({{ live_jogos|length }})</div>
{% for j in live_jogos %}{% set s = snaps_mapa.get(j.flash_mid) %}
{% set is_ht = s and s.minuto_est == -1 %}
{% set min_show = 45 if is_ht else (s.minuto_est if s else 0) %}
{% set pct = [[(min_show / 90 * 100)|int, 0]|max, 100]|min %}
<div class="game-card">
  <div class="game-header">
    <div class="game-teams">
      {{ j.home }} vs {{ j.away }}<span class="liga-tag">{{ j.league }}</span>
      {% if j.referee %}<span style="font-size:.65rem;color:#aaa;margin-left:8px">👤 {{ j.referee }}</span>{% endif %}
    </div>
    <div style="display:flex;align-items:center;gap:8px">
      {% if is_ht %}<span class="badge-ht">HT</span>
      {% elif s %}<span class="badge-min">{{ s.minuto_est }}'</span>{% endif %}
      <span class="badge-live">LIVE</span>
    </div>
  </div>
  <div class="progress-wrap"><div class="progress-bar{% if is_ht %} ht{% endif %}" style="width:{{ pct }}%"></div></div>
  <div class="game-body">
    <div class="section s-flash">
      {% if s and s.sc_faltas_casa is not none %}
      <div class="section-label" style="color:#4db6ac">Statscore <span style="font-size:.6rem;opacity:.6">⚡ real</span></div>
      <div class="stat-row"><span class="stat-name">Faltas</span><span class="stat-val">{{ (s.sc_faltas_casa or 0) + (s.sc_faltas_fora or 0) }} <span style="font-size:.68rem;color:#666">({{ s.sc_faltas_casa }}-{{ s.sc_faltas_fora }})</span></span></div>
      <div class="stat-row"><span class="stat-name">Lanç. lat.</span><span class="stat-val">{{ (s.sc_lanc_casa or 0) + (s.sc_lanc_fora or 0) }} <span style="font-size:.68rem;color:#666">({{ s.sc_lanc_casa }}-{{ s.sc_lanc_fora }})</span></span></div>
      <div class="stat-row"><span class="stat-name">Ritmo/90</span><span class="extrap-val-falt">{{ '%.1f'|format(s.faltas_extrap) if s and s.faltas_extrap else '—' }}</span></div>
      {% else %}
      <div class="section-label">Flashscore</div>
      <div class="stat-row"><span class="stat-name">Faltas total</span><span class="stat-val">{{ s.faltas_total if s and s.faltas_total is not none else '—' }}</span></div>
      <div class="stat-row"><span class="stat-name">Casa / Fora</span><span class="stat-val" style="font-size:.72rem">{{ (s.faltas_casa|string+' / '+s.faltas_fora|string) if s and s.faltas_casa is not none else '—' }}</span></div>
      <div class="stat-row"><span class="stat-name">Ritmo/90</span><span class="extrap-val-falt">{{ '%.1f'|format(s.faltas_extrap) if s and s.faltas_extrap else '—' }}</span></div>
      {% endif %}
    </div>
    <div class="section s-pred">
      <div class="section-label">Previsão</div>
      <div class="stat-row"><span class="stat-name">Pred</span><span class="pred-val-falt">{{ '%.1f'|format(s.faltas_pred) if s and s.faltas_pred else '—' }}</span></div>
      <div class="stat-row"><span class="stat-name">Base</span><span class="muted-val">{{ j.faltas_baseline }}</span></div>
      {% if j.referee %}
      <div class="divt"></div>
      <div class="stat-row"><span class="stat-name" style="color:#6a4f8a">Árbitro</span><span style="font-size:.7rem;color:#7a5f9a">{{ j.referee }}</span></div>
      {% endif %}
    </div>
    <div class="section s-bet-falt">
      <div class="section-label">22bet Linha</div>
      {% if s and s.fl_live_line %}
      <div class="bet-line-falt">{{ s.fl_live_line }}</div>
      <div class="bet-odds bet-odds-falt">
        <span>O <b>{{ s.fl_live_over_odds or '—' }}</b></span>
        <span>U <b>{{ s.fl_live_under_odds or '—' }}</b></span>
      </div>
      {% if s.fl_live_signal and s.fl_live_signal.startswith('OVER') %}<div class="signal-over">▲ OVER</div>
      {% elif s.fl_live_signal and s.fl_live_signal.startswith('UNDER') %}<div class="signal-under">▼ UNDER</div>
      {% else %}<div class="signal-neu">— SEM SINAL</div>{% endif %}
      <div class="signal-delta">Pred {{ '%.1f'|format(s.faltas_pred) if s.faltas_pred else '?' }} vs {{ s.fl_live_line }}{% if s.fl_live_signal %} Δ {{ s.fl_live_signal[4:] if s.fl_live_signal.startswith('OVER') else s.fl_live_signal[5:] if s.fl_live_signal.startswith('UNDER') else s.fl_live_signal[3:] }}{% endif %}</div>
      {% else %}
      <div class="bet-waiting">{% if _FALTAS_TI_ACTIVO %}A aguardar...{% else %}TI faltas pendente{% endif %}</div>
      {% endif %}
    </div>
  </div>
</div>
{% endfor %}{% endif %}

{% if pending_jogos %}
<div class="section-h">A aguardar ({{ pending_jogos|length }})</div>
<div class="pending-box">
{% for j in pending_jogos %}{% set s = snaps_mapa.get(j.flash_mid) %}
<div class="pending-row">
  <span>{{ j.home }} vs {{ j.away }} <span style="color:#555">{{ j.league }}</span>{% if j.referee %} <span style="color:#4a3a5a">· {{ j.referee }}</span>{% endif %}</span>
  <span style="display:flex;gap:12px;align-items:center">
    {% if s %}
      <span>Pred: <b style="color:#9b6bd5">{{ '%.1f'|format(s.faltas_pred) if s.faltas_pred else '—' }}</b></span>
      {% if s.fl_live_line %}<span>Linha 22bet: <b style="color:#88c868">{{ s.fl_live_line }}</b></span>
        {% if s.fl_live_over_odds %}<span style="color:#888">O:{{ s.fl_live_over_odds }} U:{{ s.fl_live_under_odds }}</span>{% endif %}
        {% if s.fl_live_signal %}<span class="sig sig-{{ 'over' if 'OVER' in s.fl_live_signal else ('under' if 'UNDER' in s.fl_live_signal else 'neu') }}">{{ s.fl_live_signal }}</span>{% endif %}
      {% endif %}
    {% else %}
      {% if j.faltas_baseline %}<span style="color:#6a4f8a">base {{ j.faltas_baseline }}</span>{% endif %}
    {% endif %}
    <span style="color:#555">{{ j.kickoff[11:16] if j.kickoff else '—' }} UTC</span>
  </span>
</div>{% endfor %}
</div>{% endif %}

{% if done_jogos %}
<div class="section-h">Terminados ({{ done_jogos|length }})</div>
<div class="done-box">
{% for j in done_jogos %}{% set s = snaps_mapa.get(j.flash_mid) %}
<div class="done-row">
  <span>{{ j.home }} vs {{ j.away }} <span style="color:#2a2d3a">{{ j.league }}</span></span>
  <span style="display:flex;gap:10px">{% if s %}
    <span>Pred: <b style="color:#9b6bd5">{{ '%.1f'|format(s.faltas_pred) if s.faltas_pred else '—' }}</b></span>
    {% if s.fl_live_line %}<span>Linha: <b style="color:#88c868">{{ s.fl_live_line }}</b></span>{% endif %}
  {% endif %}</span>
</div>{% endfor %}
</div>{% endif %}

{% if not live_jogos and not pending_jogos and not done_jogos %}
<div class="empty">Sem jogos registados. Corre setup_weekend.py antes do fim de semana.</div>
{% endif %}
<script>
let c=30;const el=document.getElementById('cnt');
setInterval(()=>{c--;if(c<=0)location.reload();else el.textContent=c},1000);
function uploadFaltasModel(input){
  if(!input.files||!input.files[0]) return;
  const btn=document.getElementById('btnUpdateModel');
  const msg=document.getElementById('updateModelMsg');
  btn.disabled=true; btn.textContent='⏳ A atualizar...';
  const fd=new FormData(); fd.append('file',input.files[0]);
  fetch('/admin/update-faltas-model',{method:'POST',body:fd})
    .then(r=>r.json()).then(d=>{
      msg.style.display='block';
      if(d.ok){
        msg.style.background='#1a3d2b';msg.style.color='#6fcf97';msg.style.border='1px solid #2a6b4f';
        const t=d.updated_at?d.updated_at.substring(0,16).replace('T',' '):'—';
        msg.textContent=`✅ Modelo atualizado — ${d.equipas} equipas · ${d.arbitros} árbitros · ${t}`;
      } else {
        msg.style.background='#3d1a1a';msg.style.color='#eb5757';msg.style.border='1px solid #6b2a2a';
        msg.textContent='❌ Erro: '+(d.error||'desconhecido');
      }
      btn.disabled=false;btn.textContent='🔄 Atualizar dados faltas';input.value='';
    }).catch(e=>{
      msg.style.display='block';msg.style.background='#3d1a1a';msg.style.color='#eb5757';
      msg.textContent='❌ Erro de ligação: '+e;
      btn.disabled=false;btn.textContent='🔄 Atualizar dados faltas';input.value='';
    });
}
</script>
</body></html>"""

# ── Rotas live ────────────────────────────────────────────────────────────────
@app.route('/api/live/add_games', methods=['POST'])
def live_add_games():
    """Recebe lista de jogos do setup_weekend.py e regista na DB."""
    data = request.get_json(force=True)
    if data.get('token') != LIVE_TOKEN:
        return jsonify({'error': 'unauthorized'}), 401
    games = data.get('games', [])
    con = init_live_db()
    added = 0
    for g in games:
        try:
            con.execute('''INSERT OR IGNORE INTO live_games
                (flash_mid,league,home,away,kickoff,lanc_baseline,faltas_baseline,status,referee)
                VALUES (?,?,?,?,?,?,?,'pending',?)''',
                (g['flash_mid'],g['league'],g['home'],g['away'],g['kickoff'],
                 g.get('lanc_baseline'), g.get('faltas_baseline'), g.get('referee')))
            added += 1
        except: pass
    con.commit(); con.close()
    logger.info(f'[LIVE] {added} jogos registados via API')
    return jsonify({'ok': True, 'added': added})

@app.route('/api/live/push', methods=['POST'])
def live_push():
    try:
        data = request.get_json(force=True)
        if data.get('token') != LIVE_TOKEN:
            return jsonify({'error': 'unauthorized'}), 401
        con = init_live_db()
        g = data.get('game', {})
        s = data.get('snapshot', {})
        if g:
            con.execute('''INSERT OR REPLACE INTO live_games
                (flash_mid,league,home,away,kickoff,lanc_baseline,faltas_baseline,status,referee,bet_ci,statscore_id)
                VALUES (?,?,?,?,?,?,?,?,?,?,?)''',
                (g['flash_mid'],g['league'],g['home'],g['away'],g['kickoff'],
                 g.get('lanc_baseline'),g.get('faltas_baseline'),g.get('status','live'),
                 g.get('referee'),g.get('bet_ci'),g.get('statscore_id')))
        if s and g:
            con.execute('''INSERT INTO live_snapshots
                (flash_mid,league,home,away,kickoff,minuto_est,
                 lanc_casa,lanc_fora,lanc_total,faltas_casa,faltas_fora,faltas_total,
                 lanc_extrap,faltas_extrap,lanc_pred,faltas_pred,captured_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                (g['flash_mid'],g['league'],g['home'],g['away'],g['kickoff'],
                 s.get('minuto_est'),s.get('lanc_casa'),s.get('lanc_fora'),s.get('lanc_total'),
                 s.get('faltas_casa'),s.get('faltas_fora'),s.get('faltas_total'),
                 s.get('lanc_extrap'),s.get('faltas_extrap'),s.get('lanc_pred'),s.get('faltas_pred'),
                 s.get('captured_at')))
        con.commit(); con.close()
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/admin/collect-now', methods=['POST'])
def admin_collect_now():
    """Força um ciclo de collect imediato."""
    data = request.get_json(force=True, silent=True) or {}
    if data.get('token') != LIVE_TOKEN:
        return jsonify({'error': 'unauthorized'}), 401
    try:
        collect_live_once()
        snaps = ler_latest_snaps()
        return jsonify({'ok': True, 'snaps': {k: {kk: v for kk,v in vv.items() if kk in ('minuto_est','lanc_pred','faltas_pred','live_line','fl_live_line','live_signal','fl_live_signal','captured_at')} for k,vv in snaps.items()}})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/admin/reload-games', methods=['POST'])
def admin_reload_games():
    """Força execução imediata do auto_load_daily_games. Body: {token}"""
    data = request.get_json(force=True, silent=True) or {}
    if data.get('token') != LIVE_TOKEN:
        return jsonify({'error': 'unauthorized'}), 401
    try:
        auto_load_daily_games()
        return jsonify({'ok': True, 'msg': 'auto_load_daily_games executado'})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/admin/set-status', methods=['POST'])
def admin_set_status():
    """Força status de um jogo: done/pending/live. Body: {token, flash_mid, status}"""
    data = request.get_json(force=True, silent=True) or {}
    if data.get('token') != LIVE_TOKEN:
        return jsonify({'error': 'unauthorized'}), 401
    flash_mid = data.get('flash_mid')
    status    = data.get('status')
    if not flash_mid or status not in ('done', 'pending', 'live'):
        return jsonify({'error': 'flash_mid e status (done/pending/live) obrigatórios'}), 400
    if not LIVE_DB_PATH.exists():
        return jsonify({'error': 'DB não existe'}), 404
    con = sqlite3.connect(LIVE_DB_PATH)
    con.execute('UPDATE live_games SET status=? WHERE flash_mid=?', (status, flash_mid))
    con.commit(); con.close()
    return jsonify({'ok': True, 'flash_mid': flash_mid, 'status': status})

@app.route('/admin/debug-live', methods=['GET'])
def admin_debug_live():
    """Debug: mostra hora do servidor, jogos na DB e janela do collector."""
    now = datetime.now(timezone.utc)
    window_start = (now - timedelta(minutes=115)).isoformat()
    window_end   = (now + timedelta(hours=3)).isoformat()
    if not LIVE_DB_PATH.exists():
        return jsonify({'error': 'DB não existe', 'db_path': str(LIVE_DB_PATH)})
    con = sqlite3.connect(LIVE_DB_PATH); con.row_factory = sqlite3.Row
    games = [dict(r) for r in con.execute('SELECT * FROM live_games').fetchall()]
    in_window = [dict(r) for r in con.execute(
        '''SELECT * FROM live_games WHERE status IN ("pending","live")
           AND kickoff <= ? AND kickoff >= ?''', (window_end, window_start)).fetchall()]
    snap_count = con.execute('SELECT COUNT(*) FROM live_snapshots').fetchone()[0]
    con.close()
    return jsonify({'server_utc': now.isoformat(), 'window_start': window_start,
                    'window_end': window_end, 'all_games': games,
                    'in_window': in_window, 'snap_count': snap_count})

@app.route('/admin/snaps-detail', methods=['GET'])
def admin_snaps_detail():
    """Retorna todos os snapshots em detalhe (para exportação Excel)."""
    if not LIVE_DB_PATH.exists():
        return jsonify({'error': 'DB não existe'})
    con = sqlite3.connect(LIVE_DB_PATH); con.row_factory = sqlite3.Row
    rows = con.execute('''SELECT flash_mid, home, away, league,
        minuto_est, lanc_total, faltas_total,
        lanc_extrap, faltas_extrap, lanc_pred, faltas_pred,
        live_line, live_over_odds, live_under_odds, live_signal,
        fl_live_line, fl_live_over_odds, fl_live_under_odds, fl_live_signal,
        sc_lanc_casa, sc_lanc_fora, sc_faltas_casa, sc_faltas_fora,
        captured_at
        FROM live_snapshots ORDER BY captured_at''').fetchall()
    con.close()
    return jsonify({'snaps': [dict(r) for r in rows], 'total': len(rows)})

@app.route('/admin/snap-stats', methods=['GET'])
def admin_snap_stats():
    """Resumo de sinais OVER/UNDER/NEU nos snapshots da sessão."""
    if not LIVE_DB_PATH.exists():
        return jsonify({'error': 'DB não existe'})
    con = sqlite3.connect(LIVE_DB_PATH); con.row_factory = sqlite3.Row
    rows = con.execute('''SELECT flash_mid, home, away,
        live_signal, fl_live_signal, minuto_est, captured_at
        FROM live_snapshots ORDER BY captured_at''').fetchall()
    con.close()

    def classify(sig):
        if not sig: return 'neu'
        s = str(sig)
        if s.startswith('OVER'): return 'over'
        if s.startswith('UNDER'): return 'under'
        return 'neu'

    summary = {}
    for r in rows:
        key = r['flash_mid']
        if key not in summary:
            summary[key] = {'home': r['home'], 'away': r['away'],
                            'lanc': {'over':0,'under':0,'neu':0},
                            'falt': {'over':0,'under':0,'neu':0},
                            'total_snaps': 0}
        summary[key]['total_snaps'] += 1
        summary[key]['lanc'][classify(r['live_signal'])] += 1
        summary[key]['falt'][classify(r['fl_live_signal'])] += 1

    return jsonify({'jogos': list(summary.values()), 'total_snaps': len(rows)})

@app.route('/admin/jogos-status', methods=['GET'])
def admin_jogos_status():
    """Diagnóstico: quantos jogos há em lancamentos_lines.db e quando foi o último monitor."""
    if not DB_PATH.exists():
        return jsonify({'error': 'DB não existe', 'path': str(DB_PATH)})
    try:
        con = sqlite3.connect(DB_PATH); con.row_factory = sqlite3.Row
        tbls = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
        if 'lancamentos_lines' not in tbls:
            con.close()
            return jsonify({'error': 'tabela lancamentos_lines não existe', 'tabelas': list(tbls)})
        rows = con.execute('SELECT match_id,league,home,away,kickoff,modelo_pred,edge_signal,detected_at FROM lancamentos_lines ORDER BY kickoff').fetchall()
        con.close()
        return jsonify({'total': len(rows), 'ultima_exec': ultima_exec, 'proxima_exec': proxima_exec,
                        'jogos': [dict(r) for r in rows]})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/admin/jogos-push', methods=['POST'])
def admin_jogos_push():
    """Push de jogos pré-jogo do local para o Railway.
    Body: {token, jogos: [{match_id,league,home,away,kickoff,line,over_odds,under_odds,
                           all_lines,modelo_pred,edge_signal,flash_mid}]}
    """
    data = request.get_json(force=True, silent=True) or {}
    if data.get('token') != LIVE_TOKEN:
        return jsonify({'error': 'unauthorized'}), 401
    jogos_in = data.get('jogos', [])
    if not jogos_in:
        return jsonify({'error': 'lista jogos vazia'}), 400
    con = sqlite3.connect(DB_PATH)
    con.execute('''CREATE TABLE IF NOT EXISTS lancamentos_lines (
        id TEXT PRIMARY KEY,
        match_id TEXT, league TEXT, home TEXT, away TEXT,
        kickoff TEXT, line REAL, over_odds REAL, under_odds REAL,
        all_lines TEXT, modelo_pred REAL, edge_signal TEXT,
        flash_mid TEXT, detected_at TEXT)''')
    added = updated = 0
    now_str = datetime.now(timezone.utc).isoformat()
    for j in jogos_in:
        mid = j.get('match_id','')
        existing = con.execute('SELECT id FROM lancamentos_lines WHERE match_id=?', (mid,)).fetchone()
        if existing:
            con.execute('''UPDATE lancamentos_lines SET league=?,home=?,away=?,kickoff=?,line=?,
                over_odds=?,under_odds=?,all_lines=?,modelo_pred=?,edge_signal=?,flash_mid=?
                WHERE match_id=?''',
                (j.get('league'),j.get('home'),j.get('away'),j.get('kickoff'),
                 j.get('line'),j.get('over_odds'),j.get('under_odds'),
                 json.dumps(j.get('all_lines',[])) if isinstance(j.get('all_lines'), list) else j.get('all_lines','[]'),
                 j.get('modelo_pred'),j.get('edge_signal'),j.get('flash_mid'), mid))
            updated += 1
        else:
            row_id = j.get('id') or mid
            con.execute('''INSERT OR IGNORE INTO lancamentos_lines
                (id,match_id,league,home,away,kickoff,line,over_odds,under_odds,
                 all_lines,modelo_pred,edge_signal,flash_mid,detected_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                (row_id, mid, j.get('league'),j.get('home'),j.get('away'),j.get('kickoff'),
                 j.get('line'),j.get('over_odds'),j.get('under_odds'),
                 json.dumps(j.get('all_lines',[])) if isinstance(j.get('all_lines'), list) else j.get('all_lines','[]'),
                 j.get('modelo_pred'),j.get('edge_signal'),j.get('flash_mid'), now_str))
            added += 1
    con.commit(); con.close()
    return jsonify({'ok': True, 'added': added, 'updated': updated, 'total': added + updated})

@app.route('/admin/update-faltas-model', methods=['POST'])
def update_faltas_model():
    """Recebe model_data.json via upload ou lê do path local, e actualiza dados em memória + cache."""
    try:
        # Modo 1: ficheiro enviado via upload (multipart)
        if 'file' in request.files:
            f = request.files['file']
            data = json.load(f.stream)
        # Modo 2: path local (fallback para uso em desenvolvimento)
        else:
            path = Path(request.form.get('path') or FALTAS_MODEL_PATH)
            if not path.exists():
                return jsonify({'error': f'Ficheiro não encontrado: {path}'}), 404
            with open(path, encoding='utf-8') as f:
                data = json.load(f)

        new_teams, new_refs, new_base, new_coefs = {}, {}, {}, {}
        for json_liga, mod_liga in _LIGA_MAP.items():
            # Equipas: lista de dicts com campo 'name'
            raw_teams = data.get('teams', {}).get(json_liga, [])
            new_teams[mod_liga] = {
                t['name']: dict(ffh=float(t.get('ffh',0)), fsh=float(t.get('fsh',0)),
                                ffa=float(t.get('ffa',0)), fsa=float(t.get('fsa',0)))
                for t in raw_teams if t.get('name')
            }
            new_refs[mod_liga] = {
                r['name']: r['avg_total']
                for r in data.get('refs', {}).get(json_liga, [])
                if r.get('n_games', 0) > 0
            }
            new_base[mod_liga]  = data.get('league_avg', {}).get(json_liga, 26.0)
            new_coefs[mod_liga] = data.get('coefs', {}).get(json_liga, {})

        _apply_faltas_model(new_teams, new_refs, new_base, new_coefs)
        now_str = datetime.now().isoformat()
        cache = {'teams': new_teams, 'refs': new_refs, 'base': new_base, 'coefs': new_coefs,
                 'updated_at': now_str}
        with open(FALTAS_CACHE_PATH, 'w', encoding='utf-8') as cf:
            json.dump(cache, cf, ensure_ascii=False, indent=2)
        n_teams = sum(len(v) for v in new_teams.values())
        n_refs  = sum(len(v) for v in new_refs.values())
        logger.info(f'[FALTAS] modelo atualizado: {n_teams} equipas, {n_refs} árbitros')
        return jsonify({'ok': True, 'equipas': n_teams, 'arbitros': n_refs, 'updated_at': now_str})
    except Exception as e:
        logger.error(f'[FALTAS] update-faltas-model: {e}')
        return jsonify({'error': str(e)}), 500

def _live_context():
    """Dados comuns às duas páginas live."""
    jogos      = ler_live_jogos()
    snaps_mapa = ler_latest_snaps()
    live_jogos    = [j for j in jogos if j['status'] == 'live']
    pending_jogos = [j for j in jogos if j['status'] == 'pending']
    done_jogos    = [j for j in jogos if j['status'] == 'done']
    return live_jogos, pending_jogos, done_jogos, snaps_mapa

def _alertas_lanc(live_jogos, snaps_mapa):
    alertas = []
    for j in live_jogos:
        s = snaps_mapa.get(j['flash_mid'])
        if not s: continue
        sig = s.get('live_signal', '')
        if sig and not sig.startswith('NEU'):
            d = sig[4:] if sig.startswith('OVER') else sig[5:]
            alertas.append(f"{j['home']} vs {j['away']} · Lanç {'OVER' if sig.startswith('OVER') else 'UNDER'} · Pred {s.get('lanc_pred')} vs {s.get('live_line')} · Δ{d}")
    return alertas

def _alertas_faltas(live_jogos, snaps_mapa):
    alertas = []
    for j in live_jogos:
        s = snaps_mapa.get(j['flash_mid'])
        if not s: continue
        sig = s.get('fl_live_signal', '')
        if sig and not sig.startswith('NEU'):
            d = sig[4:] if sig.startswith('OVER') else sig[5:]
            alertas.append(f"{j['home']} vs {j['away']} · Faltas {'OVER' if sig.startswith('OVER') else 'UNDER'} · Pred {s.get('faltas_pred')} vs {s.get('fl_live_line')} · Δ{d}")
    return alertas

# ── Jogos FDS — TheSportsDB cache ────────────────────────────────────────────
_TSDB_CACHE      = {}          # liga -> list of game dicts
_TSDB_CACHE_TIME = {}          # liga -> datetime fetched
_REFEREE_OVERRIDES = {}        # norm(home)+'__'+norm(away) -> nome_arbitro

LIGAS_TSDB = [
    {'code':'PPL','id':4344,'nome':'Liga Portugal Betclic','flag':'🇵🇹','season':'2026-2027'},
    {'code':'BRA','id':4351,'nome':'Brasil: Série A Betano','flag':'🇧🇷','season':'2026'},
    {'code':'ESP','id':4335,'nome':'Espanha: LaLiga','flag':'🇪🇸','season':'2026-2027'},
]

def _tsdb_fetch(url):
    import gzip as _gz
    try:
        req = urllib.request.Request(url, headers={
            'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
            'Accept':'application/json, */*',
        })
        with urllib.request.urlopen(req, timeout=12) as r:
            raw = r.read()
            try: raw = _gz.decompress(raw)
            except: pass
            return json.loads(raw.decode('utf-8','replace'))
    except Exception as e:
        logger.warning(f'[TSDB] {url[-60:]}: {e}')
        return None

def _fetch_liga(liga):
    """Busca jogos de uma liga no TheSportsDB. Retorna lista de dicts."""
    base = 'https://www.thesportsdb.com/api/v1/json/3'
    lid, season = liga['id'], liga['season']
    # 1. Descobrir jornada actual
    nxt = _tsdb_fetch(f'{base}/eventsnextleague.php?id={lid}')
    if not nxt or not nxt.get('events'):
        return []
    rnd = int(nxt['events'][0].get('intRound') or 1)
    # 2. Buscar jogos das próximas 2 jornadas
    jogos = []
    for r in range(rnd, rnd + 2):
        data = _tsdb_fetch(f'{base}/eventsround.php?id={lid}&r={r}&s={season}')
        if data and data.get('events'):
            jogos.extend(data['events'])
    return jogos

def _get_tsdb_jogos():
    """Retorna todos os jogos com cache de 60 min."""
    now = datetime.now(timezone.utc)
    result = []
    for liga in LIGAS_TSDB:
        code = liga['code']
        age  = (now - _TSDB_CACHE_TIME.get(code, datetime.min.replace(tzinfo=timezone.utc))).total_seconds()
        if age > 3600 or code not in _TSDB_CACHE:
            raw = _fetch_liga(liga)
            _TSDB_CACHE[code]      = raw
            _TSDB_CACHE_TIME[code] = now
            logger.info(f'[TSDB] {code}: {len(raw)} eventos carregados')
        for ev in _TSDB_CACHE.get(code, []):
            ts = ev.get('strTimestamp','')
            try:
                ko = datetime.fromisoformat(ts).replace(tzinfo=timezone.utc)
                from zoneinfo import ZoneInfo
                ko_lisbon = ko.astimezone(ZoneInfo('Europe/Lisbon'))
            except:
                continue
            home = ev.get('strHomeTeam','')
            away = ev.get('strAwayTeam','')
            # árbitro: primeiro do override (push do faltas.db local), depois do TheSportsDB
            _ref_key = _norm_nome(home) + '__' + _norm_nome(away)
            arbitro = _REFEREE_OVERRIDES.get(_ref_key)
            if not arbitro:
                # fuzzy: match por palavras significativas (len>2) em home e away
                _h_w = {w for w in _norm_nome(home).split() if len(w) > 2}
                _a_w = {w for w in _norm_nome(away).split() if len(w) > 2}
                for _k, _v in _REFEREE_OVERRIDES.items():
                    _kh, _ka = (_k.split('__', 1) + [''])[:2]
                    _kh_w = {w for w in _kh.split() if len(w) > 2}
                    _ka_w = {w for w in _ka.split() if len(w) > 2}
                    if _h_w & _kh_w and _a_w & _ka_w:
                        arbitro = _v; break
            arbitro = arbitro or (ev.get('strReferee') or '').strip() or None
            hs  = _get_stats(home)
            as_ = _get_stats(away)
            # adicionar faltas_ctx: FFH para casa, FFA para fora
            if hs:
                hs = dict(hs)
                hs['faltas_ctx'] = hs.get('ffh') or hs['faltas_media']
            if as_:
                as_ = dict(as_)
                as_['faltas_ctx'] = as_.get('ffa') or as_['faltas_media']
            arb_stats = _get_arbitro_stats(code, arbitro) if arbitro else None
            pred_auto = round(hs['lanc_media'] + as_['lanc_media'], 1) if (hs and as_) else None
            # falt_auto usa FFH (casa) + FFA (fora) — valores contextuais
            falt_auto = round(hs['faltas_ctx'] + as_['faltas_ctx'], 1) if (hs and as_) else None
            # previsão com factor árbitro se disponível
            falt_pred = None
            if falt_auto and arb_stats:
                # média simples entre previsão baseada em equipas e média do árbitro
                falt_pred = round((falt_auto + arb_stats['media_faltas']) / 2, 1)
            result.append({
                'code':        code,
                'liga_nome':   liga['nome'],
                'liga_flag':   liga['flag'],
                'home':        home,
                'away':        away,
                'arbitro':     arbitro,
                'arbitro_stats': arb_stats,
                'kickoff':     ts,
                'ko_dt':       ko,                      # usado só em Python (sort/group)
                'ko_str':      ko.isoformat(),           # string — seguro para JSON
                'hora':        ko_lisbon.strftime('%H:%M'),
                'data':        ko.date(),                # usado só em Python (dia_label)
                'data_str':    ko.date().isoformat(),    # string — seguro para JSON
                'status':      ev.get('strStatus','NS'),
                'jornada':     ev.get('intRound',''),
                'home_stats':  hs,
                'away_stats':  as_,
                'pred_auto':   pred_auto,
                'falt_auto':   falt_auto,
                'falt_pred':   falt_pred,
                'past':        ko < now,
            })
    return result

@app.route('/jogos')
def jogos_page():
    from collections import OrderedDict
    now_utc = datetime.now(timezone.utc)
    DIAS_PT = {0:'Segunda',1:'Terça',2:'Quarta',3:'Quinta',4:'Sexta',5:'Sábado',6:'Domingo'}

    def dia_label(d):
        hoje   = now_utc.date()
        amanha = (now_utc + timedelta(days=1)).date()
        if d == hoje:    return f'Hoje - {d.strftime("%d.%m.")}'
        if d == amanha:  return f'Amanhã - {d.strftime("%d.%m.")}'
        return f'{DIAS_PT[d.weekday()]} - {d.strftime("%d.%m.")}'

    jogos = _get_tsdb_jogos()
    # ordenar e agrupar: dia -> liga -> jogos
    dias = OrderedDict()
    for j in sorted(jogos, key=lambda x: (x['ko_dt'], x['liga_nome'])):
        dl = dia_label(j['data'])
        ln = j['liga_nome']
        if dl not in dias: dias[dl] = OrderedDict()
        if ln not in dias[dl]: dias[dl][ln] = []
        dias[dl][ln].append(j)

    cache_info = {c: _TSDB_CACHE_TIME.get(c,'') for c in ['PPL','BRA','ESP']}
    return render_template_string(JOGOS_TEMPLATE, dias=dias, now_utc=now_utc)

@app.route('/api/jogos/refresh', methods=['POST'])
def jogos_refresh():
    """Força refresh do cache de jogos."""
    global _TSDB_CACHE, _TSDB_CACHE_TIME
    _TSDB_CACHE = {}; _TSDB_CACHE_TIME = {}
    return jsonify({'ok': True, 'msg': 'Cache limpo. Próximo acesso vai buscar dados novos.'})

@app.route('/admin/arbitros-push', methods=['POST'])
def admin_arbitros_push():
    """Recebe nomeações de árbitros do faltas.db local e guarda em memória.
    Body: {token, arbitros: [{home, away, referee, league}]}
    """
    global _REFEREE_OVERRIDES
    data = request.get_json(force=True, silent=True) or {}
    if data.get('token') != LIVE_TOKEN:
        return jsonify({'error': 'token inválido'}), 403
    arbitros = data.get('arbitros', [])
    added = 0
    for a in arbitros:
        home = a.get('home','').strip()
        away = a.get('away','').strip()
        ref  = a.get('referee','').strip()
        if home and away and ref:
            key = _norm_nome(home) + '__' + _norm_nome(away)
            _REFEREE_OVERRIDES[key] = ref
            added += 1
    logger.info(f'[ARBITROS-PUSH] {added} nomeações recebidas, total={len(_REFEREE_OVERRIDES)}')
    return jsonify({'ok': True, 'added': added, 'total': len(_REFEREE_OVERRIDES)})

@app.route('/admin/arbitros-status', methods=['GET'])
def admin_arbitros_status():
    return jsonify({'total': len(_REFEREE_OVERRIDES),
                    'arbitros': [{'key': k, 'ref': v} for k,v in _REFEREE_OVERRIDES.items()]})

@app.route('/debug/ninja')
def debug_ninja():
    """Testa o _fetch_ninja para um flash_mid. Ex: /debug/ninja?mid=hpkd2WE4"""
    mid = request.args.get('mid', '')
    if not mid:
        return jsonify({'error': 'mid em falta'})
    try:
        req = urllib.request.Request(_NINJA_URL.format(mid=mid), headers=_FS_HDR)
        with urllib.request.urlopen(req, timeout=12) as r:
            raw = r.read().decode('utf-8', errors='replace')
        stats = _fetch_ninja(mid)
        # Extract all stat names found
        all_stats = []
        for m in _STAT_RE.finditer(raw):
            all_stats.append({'name': m.group(1).strip(), 'home': m.group(2).strip(), 'away': m.group(3).strip()})
        return jsonify({'mid': mid, 'parsed': stats, 'all_stats': all_stats, 'raw_len': len(raw)})
    except Exception as e:
        return jsonify({'mid': mid, 'error': str(e)})

@app.route('/debug/arbitro')
def debug_arbitro():
    """Testa o lookup de árbitro. Ex: /debug/arbitro?nome=Bessa+J.&liga=PPL"""
    nome  = request.args.get('nome', '')
    liga  = request.args.get('liga', 'PPL')
    key   = _norm_nome(nome)
    stats = _get_arbitro_stats(liga, nome)
    # mostrar todos os árbitros disponíveis para essa liga
    available = list(_ARBITROS.get(liga, {}).keys())
    return jsonify({
        'input_nome':   nome,
        'input_liga':   liga,
        'norm_key':     key,
        'key_words':    [w for w in key.split() if len(w) > 2],
        'found':        stats is not None,
        'stats':        stats,
        'available_refs_count': len(available),
        'available_refs': available,
    })

@app.route('/debug/medias')
def debug_medias():
    """Mostra o que está carregado em _MEDIAS e _ARBITROS para diagnóstico."""
    arbs_out = {}
    for liga, arbs in _ARBITROS.items():
        arbs_out[liga] = [
            {'nome': v['nome'], 'jogos': v['jogos'],
             'media_faltas': v['media_faltas'],
             'last_totals': v.get('last_totals', [])}
            for v in arbs.values()
        ]
    sample_ppl = []
    for k, v in _MEDIAS.items():
        if v.get('liga') == 'PPL':
            sample_ppl.append({'key': k, 'equipa': v['equipa'],
                                'ffh': v.get('ffh'), 'ffa': v.get('ffa'),
                                'faltas_media': v.get('faltas_media')})
        if len(sample_ppl) >= 5:
            break
    out = {
        'medias_total': len(_MEDIAS),
        'arbitros_total': {liga: len(arbs) for liga, arbs in _ARBITROS.items()},
        'arbitros': arbs_out,
        'sample_ppl_equipas': sample_ppl,
    }
    return jsonify(out)

# ── Auto-fetch de árbitros (corre no Railway, sem PC) ─────────────────────────
import html as _html_mod, time as _time_mod

def _arb_fetch(url, timeout=15, ua=None):
    hdrs = {'User-Agent': ua or 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'text/html,*/*', 'Accept-Language': 'pt-PT,pt;q=0.9,en;q=0.8'}
    req = urllib.request.Request(url, headers=hdrs)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            raw = r.read()
            try:
                import gzip as _gz; raw = _gz.decompress(raw)
            except Exception: pass
            return raw.decode(r.headers.get_content_charset() or 'utf-8', errors='replace')
    except Exception as e:
        logger.warning(f'[ARB-FETCH] {url[-70:]}: {e}')
        return None

def _arb_strip_tags(s):
    return _re.sub(r'<[^>]+>', ' ', str(s))

def _arb_set(home, away, referee, source='?'):
    key = _norm_nome(home) + '__' + _norm_nome(away)
    if key not in _REFEREE_OVERRIDES:
        _REFEREE_OVERRIDES[key] = referee
        logger.info(f'[ARB-AUTO/{source}] {home} vs {away} → {referee}')

# ── Source A: Flashscore PPL (Googlebot UA + ninja API) ──────────────────────
_FS_MATCH_RE = _re.compile(r'~AA÷([a-zA-Z0-9]{8})¬(.*?)(?=~AA÷|~~|\Z)', _re.DOTALL)
_FS_FIELD_RE = _re.compile(r'([A-Z]{1,4})÷([^¬~\n]*)')
_FS_REF_RE   = _re.compile(r'MIT÷REF¬MIV÷([^¬~\n]+)', _re.IGNORECASE)
_FS_SIGN_RE  = _re.compile(r'SW9D1eZo|"sign"\s*:\s*"([A-Za-z0-9]{6,12})"')

def _arb_flashscore_ppl():
    GBOT_UA = 'Mozilla/5.0 (compatible; Googlebot/2.1; +http://www.google.com/bot.html)'
    html = _arb_fetch('https://www.flashscore.pt/futebol/portugal/liga-portugal-betclic/', ua=GBOT_UA)
    if not html or len(html) < 10000:
        return
    fsign = 'SW9D1eZo'
    ms = _FS_SIGN_RE.search(html)
    if ms:
        fsign = ms.group(0) if ms.group(0).isalnum() else (ms.group(1) or fsign)
    now = _time_mod.time()
    window_end = now + 86400 * 12
    for rec in _FS_MATCH_RE.finditer(html):
        mid = rec.group(1)
        fields = {kv.group(1): kv.group(2).strip() for kv in _FS_FIELD_RE.finditer(rec.group(2))}
        ts = int(fields.get('AD','0') or '0')
        if ts < now - 3600 or ts > window_end: continue
        if fields.get('AB','') not in ('1','2'): continue
        home = _arb_strip_tags(fields.get('CX','')).strip()
        away = _arb_strip_tags(fields.get('AF','')).strip()
        if not home or not away: continue
        key = _norm_nome(home) + '__' + _norm_nome(away)
        if key in _REFEREE_OVERRIDES: continue
        # buscar árbitro na ninja API
        _time_mod.sleep(1)
        ninja_hdrs = {'User-Agent': 'Mozilla/5.0', 'X-Fsign': fsign,
                      'Referer': 'https://www.flashscore.pt/'}
        try:
            req = urllib.request.Request(
                f'https://global.flashscore.ninja/20/x/feed/df_sur_1_{mid}',
                headers=ninja_hdrs)
            with urllib.request.urlopen(req, timeout=12) as r:
                raw = r.read()
                try:
                    import gzip as _gz; raw = _gz.decompress(raw)
                except Exception: pass
                ref_raw = raw.decode('utf-8', errors='replace')
            rm = _FS_REF_RE.search(ref_raw)
            if rm:
                ref = rm.group(1).strip()
                if ref and len(ref) > 3 and ref not in ('-','TBA','TBC'):
                    _arb_set(home, away, ref, 'fs-ppl')
        except Exception as e:
            logger.debug(f'[ARB-AUTO/fs] ninja {mid}: {e}')

# ── Source B: maisfutebol PPL ─────────────────────────────────────────────────
_MF_ART_RE  = _re.compile(r'href=["\']'
    r'(https://maisfutebol\.iol\.pt/liga/(?:[^"\']*arbitros[^"\']*jornada[^"\']*|[^"\']*jornada[^"\']*arbitros[^"\']*))'
    r'["\']', _re.IGNORECASE)
_MF_BLOCK_RE = _re.compile(
    r'(?:<(?:strong|b)[^>]*>([^<\n]+)</(?:strong|b)>|\*\*([^*\n]+)\*\*)'
    r'(?:\s*(?:<[^>]+>|\s))*?[Áá]rbitro\s*:\s*([^\n<]+)',
    _re.IGNORECASE | _re.DOTALL)

def _arb_maisfutebol():
    for url in ('https://maisfutebol.iol.pt/liga', 'https://maisfutebol.iol.pt/liga/arbitros/'):
        html = _arb_fetch(url)
        if not html: continue
        arts = list(dict.fromkeys(_MF_ART_RE.findall(html)))
        if arts:
            _time_mod.sleep(1)
            art_html = _arb_fetch(arts[0])
            if not art_html: continue
            for m in _MF_BLOCK_RE.finditer(art_html):
                matchup = (m.group(1) or m.group(2) or '').strip()
                referee = _arb_strip_tags(m.group(3) or '').strip().rstrip(',. ')
                if not referee or len(referee) < 4: continue
                parts = matchup.split('-', 1)
                if len(parts) != 2: continue
                home, away = _arb_strip_tags(parts[0]).strip(), _arb_strip_tags(parts[1]).strip()
                if home and away: _arb_set(home, away, referee, 'mf')
            return
        _time_mod.sleep(1)

# ── Source C: worldreferee.com (La Liga ESP + outras europeias) ───────────────
_WR_LEAGUES = {'La Liga': 'ESP'}
_WR_ROW_RE  = _re.compile(r'\|\s*(\d{1,2}:\d{2})\s*\|[^|]+vs[^|]+\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|', _re.IGNORECASE)
_WR_HTML_TR = _re.compile(r'<tr[^>]*>(.*?)</tr>', _re.IGNORECASE | _re.DOTALL)
_WR_HTML_TD = _re.compile(r'<td[^>]*>(.*?)</td>', _re.IGNORECASE | _re.DOTALL)
_WR_LIG_RE  = _re.compile(r'\b(La Liga)\b', _re.IGNORECASE)

def _arb_worldreferee():
    html = _arb_fetch('https://worldreferee.com/upcoming')
    if not html: return
    current_league = None
    # HTML table format
    for seg in _re.compile(r'(<tr[^>]*>.*?</tr>)', _re.IGNORECASE | _re.DOTALL).split(html):
        if _re.match(r'<tr', seg, _re.IGNORECASE):
            if not current_league: continue
            cells = [_html_mod.unescape(_arb_strip_tags(c)).strip()
                     for c in _WR_HTML_TD.findall(seg)]
            if len(cells) < 5: continue
            hora, home, vs, away, referee = cells[0], cells[1], cells[2], cells[3], cells[4]
            if not _re.match(r'^\d{1,2}:\d{2}$', hora): continue
            if vs.lower() != 'vs': continue
            if not referee or referee.upper() == 'TBA' or len(referee) < 4: continue
            _arb_set(home, away, referee, f'wr-{current_league}')
        else:
            text = _html_mod.unescape(_re.sub(r'<[^>]+>', ' ', seg))
            lm = _WR_LIG_RE.search(text)
            if lm: current_league = 'ESP'

def _fetch_arbitros_auto():
    """Job APScheduler: vai buscar árbitros a todas as fontes. Corre no Railway."""
    logger.info('[ARB-AUTO] a actualizar árbitros...')
    try: _arb_flashscore_ppl()
    except Exception as e: logger.warning(f'[ARB-AUTO] flashscore: {e}')
    _time_mod.sleep(2)
    try: _arb_maisfutebol()
    except Exception as e: logger.warning(f'[ARB-AUTO] maisfutebol: {e}')
    _time_mod.sleep(2)
    try: _arb_worldreferee()
    except Exception as e: logger.warning(f'[ARB-AUTO] worldreferee: {e}')
    logger.info(f'[ARB-AUTO] concluído — {len(_REFEREE_OVERRIDES)} árbitros em memória')

JOGOS_TEMPLATE = """<!DOCTYPE html>
<html lang="pt"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Jogos do Fim de Semana</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#0b1525;color:#c8d4e0;font-size:14px;min-height:100vh}
.page-header{background:#0e1d30;border-bottom:2px solid #1a2d42;padding:14px 20px;display:flex;align-items:center;gap:10px;position:sticky;top:0;z-index:10}
.page-header h1{font-size:.95rem;font-weight:700;color:#fff}
.page-header .meta{font-size:.7rem;color:#3a5070;margin-left:auto;display:flex;align-items:center;gap:10px}
.refresh-btn{font-size:.68rem;padding:4px 10px;border-radius:4px;border:1px solid #1e3050;background:#132033;color:#64b5f6;cursor:pointer}
.refresh-btn:hover{background:#1a2d44}
.container{max-width:760px;margin:0 auto;padding:14px 12px 50px}
.day-header{font-size:.95rem;font-weight:700;color:#fff;margin:22px 0 10px 2px;display:flex;align-items:center;gap:8px}
.day-header::after{content:'';flex:1;height:1px;background:#1a2d42;margin-left:8px}
.league-block{background:#0e1d30;border-radius:8px;margin-bottom:8px;overflow:hidden;border:1px solid #162640}
.league-header{display:flex;align-items:center;gap:9px;padding:7px 14px;background:#132033;border-bottom:1px solid #162640}
.l-flag{font-size:.85rem}
.l-name{font-size:.7rem;font-weight:700;color:#5080a0;text-transform:uppercase;letter-spacing:.07em}
.l-round{font-size:.62rem;color:#2a4060;margin-left:auto}
.game-row{display:flex;align-items:center;padding:9px 14px;border-bottom:1px solid #0c1828;cursor:pointer;transition:background .1s;gap:10px}
.game-row:last-child{border-bottom:none}
.game-row:hover{background:#142030}
.game-row.past{opacity:.4}
.hora{font-size:.78rem;font-weight:700;color:#3a6ea5;width:36px;flex-shrink:0;text-align:center}
.teams{flex:1;min-width:0}
.tn{font-size:.84rem;font-weight:500;color:#e0eaf8;line-height:1.5;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tags{display:flex;gap:5px;flex-shrink:0;align-items:center}
.tag{font-size:.64rem;font-weight:700;padding:2px 6px;border-radius:3px;border:1px solid transparent}
.tag-pred{background:#140f25;color:#9b8fd5;border-color:#1e1540}
.tag-falt{background:#0e1825;color:#6070a0;border-color:#162030;font-weight:400}
.tag-stats{background:#132033;color:#2a4060;border-color:#1a2d42;font-size:.6rem}

/* Overlay */
.ov{display:none;position:fixed;inset:0;background:rgba(5,10,20,.82);z-index:200;align-items:center;justify-content:center;padding:10px}
.ov.open{display:flex}
.popup{background:#0e1d30;border:1px solid #1e3050;border-radius:12px;width:min(520px,97vw);max-height:88vh;overflow-y:auto}
.pop-head{padding:18px 20px 12px;border-bottom:1px solid #142030;position:relative}
.pop-x{position:absolute;top:14px;right:16px;background:none;border:none;color:#2a4060;font-size:1rem;cursor:pointer;line-height:1}
.pop-x:hover{color:#e0eaf8}
.pop-title{font-size:.98rem;font-weight:700;color:#fff;margin-bottom:3px;padding-right:28px;line-height:1.3}
.pop-sub{font-size:.7rem;color:#3a5070}
.ptabs{display:flex;gap:0;border-bottom:1px solid #142030}
.ptab{padding:9px 16px;font-size:.73rem;font-weight:600;cursor:pointer;border:none;background:none;color:#3a5070;border-bottom:2px solid transparent;margin-bottom:-1px}
.ptab.on{color:#64b5f6;border-bottom-color:#64b5f6}
.pbody{padding:16px 20px 20px}
.pp{display:none}
.pp.on{display:block}
.sr{display:flex;justify-content:space-between;align-items:center;padding:7px 0;border-bottom:1px solid #0c1828;font-size:.78rem}
.sr:last-child{border-bottom:none}
.sk{color:#3a5070}
.sv{font-weight:700;color:#e0eaf8}
.sv.b{color:#64b5f6}.sv.g{color:#4caf50}.sv.r{color:#ef5350}.sv.p{color:#9b8fd5}.sv.y{color:#ffc107}
.sec{font-size:.62rem;font-weight:700;color:#3a6ea5;text-transform:uppercase;letter-spacing:.07em;padding:10px 0 5px;display:flex;align-items:center;gap:6px}
.sec::after{content:'';flex:1;height:1px;background:#142030}
.tc{background:#132033;border-radius:7px;padding:10px 12px;margin-bottom:8px}
.tc-h{font-size:.8rem;font-weight:700;color:#e0eaf8;margin-bottom:7px;display:flex;justify-content:space-between}
.tc-m{font-size:.63rem;color:#3a5070;font-weight:400}
.sg{display:grid;grid-template-columns:repeat(3,1fr);gap:5px}
.sc{background:#0c1828;border-radius:5px;padding:6px 8px;text-align:center}
.sc-v{font-size:.88rem;font-weight:700;color:#64b5f6}
.sc-l{font-size:.56rem;color:#3a5070;margin-top:2px;text-transform:uppercase;letter-spacing:.04em}
.pred-box{background:#12102a;border:1px solid #1e1a40;border-radius:8px;padding:14px 16px;margin-top:12px;text-align:center}
.pred-big{font-size:2.2rem;font-weight:700;color:#9b8fd5;line-height:1}
.pred-lbl{font-size:.66rem;color:#3a5070;margin-top:4px}
.pred-row{display:flex;justify-content:center;gap:24px;margin-top:10px}
.pr-item .val{font-size:.95rem;font-weight:700;color:#64b5f6}
.pr-item .lbl{font-size:.6rem;color:#3a5070;margin-top:2px}
.no-d{color:#2a3a50;font-size:.75rem;text-align:center;padding:22px;font-style:italic}
.empty{text-align:center;padding:60px 20px;color:#2a3a50}
.empty .ic{font-size:2.5rem;margin-bottom:12px}
.empty .msg{font-size:.85rem}
.empty .sub{font-size:.72rem;margin-top:6px}
</style>
</head><body>

<div class="page-header">
  <span style="font-size:1.1rem">📅</span>
  <h1>Jogos do Fim de Semana</h1>
  <div class="meta">
    <span>{{ now_utc.strftime('%d/%m %H:%M') }} UTC</span>
    <span style="color:#2a4060">· Horas em GMT Lisboa</span>
    <button class="refresh-btn" onclick="refreshCache()">↻ Atualizar</button>
  </div>
</div>

<div class="container">
{% if not dias %}
  <div class="empty">
    <div class="ic">📭</div>
    <div class="msg">Nenhum jogo carregado</div>
    <div class="sub">A carregar dados das ligas...</div>
  </div>
{% endif %}

{% for dia, ligas in dias.items() %}
<div class="day-header">{{ dia }}</div>
{% for liga_nome, jgs in ligas.items() %}
<div class="league-block">
  <div class="league-header">
    <span class="l-flag">{{ jgs[0].liga_flag }}</span>
    <span class="l-name">{{ liga_nome }}</span>
    {% if jgs[0].jornada %}<span class="l-round">Jornada {{ jgs[0].jornada }}</span>{% endif %}
  </div>
  {% for j in jgs %}
  {% set jd = {'home':j.home,'away':j.away,'liga_nome':j.liga_nome,'liga_flag':j.liga_flag,'hora':j.hora,'kickoff':j.ko_str,'pred_auto':j.pred_auto,'falt_auto':j.falt_auto,'falt_pred':j.falt_pred,'home_stats':j.home_stats,'away_stats':j.away_stats,'arbitro':j.arbitro,'arbitro_stats':j.arbitro_stats,'past':j.past} %}
  <div class="game-row{% if j.past %} past{% endif %}" data-jogo='{{ jd|tojson }}' onclick="openPop(this)">
    <div class="hora">{{ j.hora }}</div>
    <div class="teams">
      <div class="tn">{{ j.home }}</div>
      <div class="tn">{{ j.away }}</div>
    </div>
  </div>
  {% endfor %}
</div>
{% endfor %}
{% endfor %}
</div>

<div class="ov" id="ov" onclick="if(event.target===this)closePop()">
<div class="popup">
  <div class="pop-head">
    <button class="pop-x" onclick="closePop()">✕</button>
    <div class="pop-title" id="p-title"></div>
    <div class="pop-sub" id="p-sub"></div>
  </div>
  <div class="ptabs">
    <button class="ptab on" onclick="ptab('est')">⚽ Estatísticas</button>
    <button class="ptab" onclick="ptab('ref')">🟨 Árbitro</button>
  </div>
  <div class="pbody">
    <div class="pp on" id="pp-est"></div>
    <div class="pp" id="pp-ref"></div>
  </div>
</div>
</div>

<script>
function openPop(el) {
  const j = JSON.parse(el.dataset.jogo);
  document.getElementById('p-title').textContent = j.home + ' vs ' + j.away;
  document.getElementById('p-sub').textContent = j.liga_flag + ' ' + j.liga_nome + '  ·  ' + (j.kickoff||'').slice(0,10) + '  ·  ' + j.hora + ' (Lisboa)';
  renderEst(j);
  renderArbitro(j);
  ptab('est');
  document.getElementById('ov').classList.add('open');
}

function renderEst(j) {
  const el = document.getElementById('pp-est');
  const h = j.home_stats, a = j.away_stats;
  if (!h && !a) { el.innerHTML = '<div class="no-d">Sem dados de médias para estas equipas.</div>'; return; }
  function card(s, nm, isHome) {
    if (!s) return '<div class="no-d">Sem dados para ' + nm + '</div>';
    const fCtx  = isHome ? (s.ffh || s.faltas_ctx || s.faltas_media) : (s.ffa || s.faltas_ctx || s.faltas_media);
    const fCtxL = isHome ? 'FF Casa' : 'FF Fora';
    const fSuf  = isHome ? (s.fsh ? ' · FS ' + s.fsh.toFixed(1) : '') : (s.fsa ? ' · FS ' + s.fsa.toFixed(1) : '');
    const hasLanc = s.lanc_media && s.lanc_media > 0;
    return '<div class="tc">'
      + '<div class="tc-h">' + nm + '<span class="tc-m">' + s.liga + ' · ' + s.jogos + ' jogos</span></div>'
      + '<div class="sg">'
      + (hasLanc ? '<div class="sc"><div class="sc-v" style="color:#e8a040">' + s.lanc_media.toFixed(1) + '</div><div class="sc-l">Lanç/jogo</div></div>' : '')
      + '<div class="sc"><div class="sc-v" style="color:#9b6bd5">' + fCtx.toFixed(1) + fSuf + '</div><div class="sc-l">' + fCtxL + '</div></div>'
      + (s.posse_media && s.posse_media > 0 ? '<div class="sc"><div class="sc-v" style="color:#5080a0">' + s.posse_media.toFixed(0) + '%</div><div class="sc-l">Posse</div></div>' : '')
      + '</div></div>';
  }
  let html = '<div class="sec">Equipas</div>' + card(h, j.home, true) + card(a, j.away, false);
  if (h && a) {
    const fCtxH   = h.ffh || h.faltas_ctx || h.faltas_media;
    const fCtxA   = a.ffa || a.faltas_ctx || a.faltas_media;
    const totL    = (h.lanc_media + a.lanc_media).toFixed(1);
    const totF    = (fCtxH + fCtxA).toFixed(1);
    const faltFinal = j.falt_pred ? j.falt_pred.toFixed(1) : totF;
    const faltLabel = j.falt_pred ? 'Faltas (c/ árbitro)' : 'Faltas (FFH+FFA)';
    const totLanc   = (h.lanc_media > 0 && a.lanc_media > 0) ? (h.lanc_media + a.lanc_media).toFixed(1) : null;
    html += '<div class="pred-box">'
      + '<div class="pred-row">'
      + (totLanc ? '<div class="pr-item"><div class="val" style="color:#e8a040">' + totLanc + '</div><div class="lbl">Lançamentos previstos</div></div>' : '')
      + '<div class="pr-item"><div class="val" style="color:#9b6bd5">' + faltFinal + '</div><div class="lbl">' + faltLabel + '</div></div>'
      + '</div></div>';
  }
  el.innerHTML = html;
}

function renderArbitro(j) {
  const el = document.getElementById('pp-ref');
  const r = j.arbitro_stats;
  const nome = j.arbitro;
  if (!nome) {
    el.innerHTML = '<div class="no-d">🟨 Árbitro ainda não nomeado.</div>';
    return;
  }
  let html = '<div class="sec">Árbitro</div>';
  html += '<div class="tc"><div class="tc-h">🟨 ' + nome;
  if (r) {
    const avg = r.media_faltas;
    html += '<span style="font-size:1rem;font-weight:600;color:#fff;margin-left:10px">' + r.jogos + ' jogos · avg ' + avg.toFixed(2) + '</span></div>';
    // últimos jogos — tiles OVER/UNDER
    const tots = (r.last_totals || []);
    if (tots.length) {
      const last10 = tots.slice(-11);  // mostrar até 11 para consistência
      html += '<div style="display:flex;flex-wrap:wrap;gap:5px;margin:10px 0">';
      last10.forEach(function(v, i) {
        const isLast10 = i >= (last10.length - 10);
        const isOver   = v >= avg;
        const bg       = isOver ? '#1a4a2e' : '#1a2a4a';
        const col      = isOver ? '#4caf50' : '#5b9bd5';
        const opacity  = isLast10 ? '1' : '0.45';
        html += '<div style="width:36px;height:36px;border-radius:6px;background:' + bg + ';display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700;color:' + col + ';opacity:' + opacity + '">' + v + '</div>';
      });
      html += '</div>';
      html += '<div style="font-size:11px;color:#888;margin-bottom:8px">🟢 OVER · 🔵 UNDER  (ref. avg ' + avg.toFixed(1) + ')</div>';
    } else {
      html += '</div>';
    }
    // previsão final de faltas com árbitro
    if (j.falt_auto) {
      const prev = j.falt_pred ? j.falt_pred.toFixed(1) : j.falt_auto.toFixed(1);
      html += '<div class="pred-box" style="margin-top:12px">'
        + '<div class="pred-big" style="color:#9b6bd5">' + prev + '</div>'
        + '<div class="pred-lbl">Previsão faltas (média equipas + árbitro)</div>'
        + '</div>';
    }
  } else {
    html += '<span class="tc-m">sem dados históricos</span></div>';
  }
  html += '</div>';
  el.innerHTML = html;
}

function closePop() { document.getElementById('ov').classList.remove('open'); }
function ptab(name) {
  document.querySelectorAll('.ptab').forEach((t,i)=>t.classList.toggle('on',['est','ref'][i]===name));
  document.querySelectorAll('.pp').forEach(p=>p.classList.remove('on'));
  document.getElementById('pp-'+name).classList.add('on');
}
function refreshCache() {
  fetch('/api/jogos/refresh', {method:'POST'}).then(()=>location.reload());
}
document.addEventListener('keydown', e=>{ if(e.key==='Escape') closePop(); });
</script>
</body></html>"""


LIVE_HISTORICO_TEMPLATE = """<!DOCTYPE html>
<html lang="pt"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Live — Histórico Snapshots</title>
<style>
{{ css }}
.snap-table{width:100%;border-collapse:collapse;font-size:.72rem;margin-top:8px}
.snap-table th{background:#0b1525;color:#ffc107;font-weight:700;padding:5px 8px;text-align:center;border:1px solid #1e2d44;position:sticky;top:0;z-index:1}
.snap-table td{padding:4px 8px;text-align:center;border:1px solid #1e2d44;color:#c0c8d8}
.snap-table tr:nth-child(even) td{background:#132033}
.snap-table tr:nth-child(odd) td{background:#0e1b2e}
.snap-table td.sig-over{background:#1b5e20;color:#fff;font-weight:700}
.snap-table td.sig-under{background:#7b0000;color:#fff;font-weight:700}
.snap-table td.sig-neu{color:#555}
.snap-table td.num-hi{color:#64b5f6;font-weight:600}
.snap-table td.num-falt{color:#9b6bd5;font-weight:600}
.snap-table td.num-line{color:#c8843a;font-weight:600}
.game-block{margin-bottom:28px}
.game-title{font-size:.88rem;font-weight:700;color:#fff;padding:8px 12px;background:#172236;border-radius:6px 6px 0 0;border:1px solid #1e2d44;display:flex;justify-content:space-between;align-items:center}
.game-meta{font-size:.7rem;color:#607d8b;font-weight:400}
.snap-count{background:#0e3a5c;color:#64b5f6;font-size:.65rem;padding:2px 7px;border-radius:10px}
.over-count{background:#1b5e20;color:#4caf50;font-size:.65rem;padding:2px 6px;border-radius:10px;margin-left:4px}
.under-count{background:#7b0000;color:#ef5350;font-size:.65rem;padding:2px 6px;border-radius:10px;margin-left:4px}
.col-group-lanc{border-left:2px solid #1a3a5c!important}
.col-group-falt{border-left:2px solid #2a1a4c!important}
.col-group-bet{border-left:2px solid #3a2a0a!important}
.empty-hist{color:#3a3e52;text-align:center;padding:40px;font-size:.85rem}
.summary-bar{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:16px;font-size:.75rem}
.summary-chip{background:#132033;border:1px solid #1e2d44;border-radius:6px;padding:5px 10px;color:#b0bec5}
.summary-chip b{color:#fff}
</style>
</head><body>
<div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px;margin-bottom:10px">
  <div>
    <h1 style="margin:0">📊 Histórico de Snapshots</h1>
    <div class="sub">Todos os jogos monitorizados · {{ total_snaps }} snapshots registados</div>
  </div>
</div>
<div class="tabs">
  <a class="tab tab-lanc" href="/live">Lançamentos</a>
  <a class="tab tab-falt" href="/live/faltas">Faltas</a>
  <a class="tab" href="/live/historico" style="color:#ffc107;border-color:#ffc107;background:#1a1600">📊 Histórico</a>
</div>

{% if not jogos_snaps %}
<div class="empty-hist">Sem snapshots registados ainda.<br>Os dados aparecem aqui durante e após os jogos monitorizados.</div>
{% else %}

<div class="summary-bar">
  <div class="summary-chip">🎮 <b>{{ jogos_snaps|length }}</b> jogos</div>
  <div class="summary-chip">📸 <b>{{ total_snaps }}</b> snapshots</div>
  <div class="summary-chip" style="border-color:#1b5e20">▲ OVER Lanç: <b style="color:#4caf50">{{ total_over_lanc }}</b></div>
  <div class="summary-chip" style="border-color:#7b0000">▼ UNDER Lanç: <b style="color:#ef5350">{{ total_under_lanc }}</b></div>
  <div class="summary-chip" style="border-color:#1b5e20">▲ OVER Falt: <b style="color:#4caf50">{{ total_over_falt }}</b></div>
  <div class="summary-chip" style="border-color:#7b0000">▼ UNDER Falt: <b style="color:#ef5350">{{ total_under_falt }}</b></div>
</div>

{% for jogo in jogos_snaps %}
<div class="game-block">
  <div class="game-title">
    <span>
      {{ jogo.home }} vs {{ jogo.away }}
      <span class="game-meta">&nbsp;·&nbsp;{{ jogo.league }}&nbsp;·&nbsp;{{ jogo.kickoff[:10] }}</span>
    </span>
    <span>
      <span class="snap-count">{{ jogo.snaps|length }} snaps</span>
      {% if jogo.ov_lanc > 0 %}<span class="over-count">▲{{ jogo.ov_lanc }} L</span>{% endif %}
      {% if jogo.un_lanc > 0 %}<span class="under-count">▼{{ jogo.un_lanc }} L</span>{% endif %}
      {% if jogo.ov_falt > 0 %}<span class="over-count" style="background:#0d2e10;color:#4caf50">▲{{ jogo.ov_falt }} F</span>{% endif %}
      {% if jogo.un_falt > 0 %}<span class="under-count" style="background:#2e0d0d;color:#ef5350">▼{{ jogo.un_falt }} F</span>{% endif %}
    </span>
  </div>
  <div style="overflow-x:auto;border:1px solid #1e2d44;border-top:none;border-radius:0 0 6px 6px">
  <table class="snap-table">
    <thead><tr>
      <th>Hora UTC</th>
      <th>Min'</th>
      <th class="col-group-lanc">SC Lanç</th>
      <th>Pred L</th>
      <th>Linha L</th>
      <th>Δ L</th>
      <th>Sinal L</th>
      <th class="col-group-falt">SC Falt</th>
      <th>Pred F</th>
      <th>Linha F</th>
      <th>Δ F</th>
      <th>Sinal F</th>
      <th class="col-group-bet">O Lanç</th>
      <th>U Lanç</th>
      <th>O Falt</th>
      <th>U Falt</th>
    </tr></thead>
    <tbody>
    {% for s in jogo.snaps %}
    {% set sc_lt = ((s.sc_lanc_casa or 0) + (s.sc_lanc_fora or 0)) if s.sc_lanc_casa is not none else (s.lanc_total or '—') %}
    {% set sc_ft = ((s.sc_faltas_casa or 0) + (s.sc_faltas_fora or 0)) if s.sc_faltas_casa is not none else (s.faltas_total or '—') %}
    {% set dl = ((s.lanc_pred or 0) - (s.live_line or 0))|round(1) if s.lanc_pred and s.live_line else '—' %}
    {% set df = ((s.faltas_pred or 0) - (s.fl_live_line or 0))|round(1) if s.faltas_pred and s.fl_live_line else '—' %}
    <tr>
      <td>{{ s.captured_at[11:19] }}</td>
      <td>{% if s.minuto_est == -1 %}HT{% else %}{{ s.minuto_est }}'{% endif %}</td>
      <td class="num-hi col-group-lanc">{% if s.sc_lanc_casa is not none %}{{ (s.sc_lanc_casa or 0) + (s.sc_lanc_fora or 0) }} <span style="font-size:.62rem;color:#555">({{ s.sc_lanc_casa }}-{{ s.sc_lanc_fora }})</span>{% else %}{{ s.lanc_total or '—' }}{% endif %}</td>
      <td class="num-hi">{{ '%.1f'|format(s.lanc_pred) if s.lanc_pred else '—' }}</td>
      <td class="num-line">{{ s.live_line or '—' }}</td>
      <td style="color:{% if dl != '—' and dl > 0 %}#4caf50{% elif dl != '—' and dl < 0 %}#ef5350{% else %}#555{% endif %}">{{ dl }}</td>
      <td class="{% if s.live_signal and s.live_signal.startswith('OVER') %}sig-over{% elif s.live_signal and s.live_signal.startswith('UNDER') %}sig-under{% else %}sig-neu{% endif %}">{{ s.live_signal or '—' }}</td>
      <td class="num-falt col-group-falt">{% if s.sc_faltas_casa is not none %}{{ (s.sc_faltas_casa or 0) + (s.sc_faltas_fora or 0) }} <span style="font-size:.62rem;color:#555">({{ s.sc_faltas_casa }}-{{ s.sc_faltas_fora }})</span>{% else %}{{ s.faltas_total or '—' }}{% endif %}</td>
      <td class="num-falt">{{ '%.1f'|format(s.faltas_pred) if s.faltas_pred else '—' }}</td>
      <td class="num-line">{{ s.fl_live_line or '—' }}</td>
      <td style="color:{% if df != '—' and df > 0 %}#4caf50{% elif df != '—' and df < 0 %}#ef5350{% else %}#555{% endif %}">{{ df }}</td>
      <td class="{% if s.fl_live_signal and s.fl_live_signal.startswith('OVER') %}sig-over{% elif s.fl_live_signal and s.fl_live_signal.startswith('UNDER') %}sig-under{% else %}sig-neu{% endif %}">{{ s.fl_live_signal or '—' }}</td>
      <td class="col-group-bet">{{ s.live_over_odds or '—' }}</td>
      <td>{{ s.live_under_odds or '—' }}</td>
      <td>{{ s.fl_live_over_odds or '—' }}</td>
      <td>{{ s.fl_live_under_odds or '—' }}</td>
    </tr>
    {% endfor %}
    </tbody>
  </table>
  </div>
</div>
{% endfor %}
{% endif %}
</body></html>"""

@app.route('/live')
def live_page():
    live_jogos, pending_jogos, done_jogos, snaps_mapa = _live_context()
    return render_template_string(LIVE_TEMPLATE,
        live_jogos=live_jogos, pending_jogos=pending_jogos,
        done_jogos=done_jogos, snaps_mapa=snaps_mapa,
        alertas=_alertas_lanc(live_jogos, snaps_mapa),
        css=_LIVE_CSS, _FALTAS_TI_ACTIVO=bool(_FALTAS_TI))

@app.route('/live/faltas')
def live_faltas_page():
    live_jogos, pending_jogos, done_jogos, snaps_mapa = _live_context()
    return render_template_string(LIVE_FALTAS_TEMPLATE,
        live_jogos=live_jogos, pending_jogos=pending_jogos,
        done_jogos=done_jogos, snaps_mapa=snaps_mapa,
        alertas=_alertas_faltas(live_jogos, snaps_mapa),
        css=_LIVE_CSS, _FALTAS_TI_ACTIVO=bool(_FALTAS_TI))

@app.route('/live/historico')
def live_historico_page():
    total_over_lanc = total_under_lanc = total_over_falt = total_under_falt = 0
    if not LIVE_DB_PATH.exists():
        jogos_snaps = []
    else:
        con = sqlite3.connect(LIVE_DB_PATH); con.row_factory = sqlite3.Row
        # Buscar todos os jogos com pelo menos 1 snap
        games = con.execute('''SELECT DISTINCT g.flash_mid, g.home, g.away, g.league, g.kickoff
            FROM live_games g
            INNER JOIN live_snapshots s ON s.flash_mid = g.flash_mid
            ORDER BY g.kickoff DESC''').fetchall()
        jogos_snaps = []
        total_over_lanc = total_under_lanc = total_over_falt = total_under_falt = 0
        for g in games:
            snaps = con.execute('''SELECT * FROM live_snapshots
                WHERE flash_mid=? ORDER BY captured_at''', (g['flash_mid'],)).fetchall()
            snaps = [dict(s) for s in snaps]
            for s in snaps:
                sig_l = s.get('live_signal') or ''
                sig_f = s.get('fl_live_signal') or ''
                if sig_l.startswith('OVER'):   total_over_lanc += 1
                elif sig_l.startswith('UNDER'): total_under_lanc += 1
                if sig_f.startswith('OVER'):   total_over_falt += 1
                elif sig_f.startswith('UNDER'): total_under_falt += 1
            ov_lanc = sum(1 for s in snaps if (s.get('live_signal') or '').startswith('OVER'))
            un_lanc = sum(1 for s in snaps if (s.get('live_signal') or '').startswith('UNDER'))
            ov_falt = sum(1 for s in snaps if (s.get('fl_live_signal') or '').startswith('OVER'))
            un_falt = sum(1 for s in snaps if (s.get('fl_live_signal') or '').startswith('UNDER'))
            jogos_snaps.append({**dict(g), 'snaps': snaps,
                                'ov_lanc': ov_lanc, 'un_lanc': un_lanc,
                                'ov_falt': ov_falt, 'un_falt': un_falt})
        con.close()

    total_snaps = sum(len(j['snaps']) for j in jogos_snaps)
    return render_template_string(LIVE_HISTORICO_TEMPLATE,
        jogos_snaps=jogos_snaps,
        total_snaps=total_snaps,
        total_over_lanc=total_over_lanc,
        total_under_lanc=total_under_lanc,
        total_over_falt=total_over_falt,
        total_under_falt=total_under_falt,
        css=_LIVE_CSS)

# ── Registar job de árbitros (função já definida acima) ───────────────────────
scheduler.add_job(_fetch_arbitros_auto, 'interval', hours=1, id='arbitros',
                  next_run_time=datetime.now())   # primeiro fetch imediato no arranque

# ── Registar auto-loader diário de jogos ──────────────────────────────────────
scheduler.add_job(auto_load_daily_games, 'cron', hour='9,13,17,19', minute=0, id='auto_load',
                  next_run_time=datetime.now())   # corre às 09h, 13h, 17h e 19h UTC + arranque

_seed_ht_totals()   # semeia HT snapshots do DB no arranque (recovery mid-2nd-half)

# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print('=' * 55)
    print(f'  Lançamentos Laterais — a correr na porta {port}')
    print(f'  DB: {DB_PATH}')
    print('  Monitor de 2 em 2 horas activo')
    print('=' * 55)
    app.run(host='0.0.0.0', port=port, debug=False, use_reloader=False)

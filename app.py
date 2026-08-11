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
import json, sqlite3, threading, logging, sys
from datetime import datetime, timezone, timedelta
from pathlib import Path
from flask import Flask, render_template_string, jsonify, Response

# ── APScheduler ───────────────────────────────────────────────────────────────
from apscheduler.schedulers.background import BackgroundScheduler

# ── Paths ─────────────────────────────────────────────────────────────────────
import os
BASE     = Path(__file__).parent
DATA_DIR = Path(os.environ.get('DATA_DIR', str(BASE)))
DB_PATH  = DATA_DIR / 'lancamentos_lines.db'
DATA_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
log_buffer = []          # guarda últimas 200 linhas de log
_running   = False       # lock para evitar runs simultâneos

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
        import monitor_lancamentos as m
        alerts = m.run_once(verbose=True)
        logger.info(f'[MONITOR] Concluído. {alerts} alertas.')
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

# ── HTML template ─────────────────────────────────────────────────────────────
TEMPLATE = """<!DOCTYPE html>
<html lang="pt">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Lançamentos · Dashboard</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',sans-serif;background:#0f1117;color:#e0e0e0;padding:20px}
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
.log-box{margin-top:24px;background:#0a0c12;border:1px solid #1e2130;border-radius:8px;
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
    <h1>⚽ Lançamentos Laterais</h1>
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
      <th>Linha</th><th>Modelo</th><th>Signal</th>
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
// Auto-refresh
let cnt = 60;
const el = document.getElementById('cnt');
setInterval(()=>{ cnt--; if(cnt<=0){location.reload()}else{el.textContent=cnt} },1000);

// Scroll log para baixo
const lb = document.getElementById('logBox');
if(lb) lb.scrollTop = lb.scrollHeight;

function correrAgora(){
  document.getElementById('btnRun').disabled = true;
  document.getElementById('btnRun').textContent = '⏳ A correr...';
  fetch('/run').then(r=>r.json()).then(d=>{
    setTimeout(()=>location.reload(), 2000);
  });
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
    jogos = ler_jogos()
    return render_template_string(TEMPLATE,
        jogos       = jogos,
        st          = stats(jogos),
        logs        = log_buffer[-80:],
        running     = _running,
        ultima_exec = ultima_exec,
        proxima_exec= proxima_exec,
    )

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

# ── Main ──────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print('=' * 55)
    print(f'  Lançamentos Laterais — a correr na porta {port}')
    print(f'  DB: {DB_PATH}')
    print('  Monitor de 2 em 2 horas activo')
    print('=' * 55)
    app.run(host='0.0.0.0', port=port, debug=False, use_reloader=False)

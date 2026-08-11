"""
Gera dashboard.html com todos os dados da lancamentos_lines.db.
Uso: python gerar_dashboard.py
Depois abre dashboard.html no browser.
"""
import sqlite3, json
from pathlib import Path
from datetime import datetime

DB_PATH  = Path(__file__).parent / 'lancamentos_lines.db'
OUT_PATH = Path(__file__).parent / 'dashboard.html'

def ler_dados():
    if not DB_PATH.exists():
        return []
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    # Descobrir colunas que realmente existem
    cols_existentes = {r[1] for r in con.execute('PRAGMA table_info(lancamentos_lines)')}
    opcionals = ['resultado_real', 'lanc_real_casa', 'lanc_real_fora', 'resultado', 'flash_mid']
    select_extra = ', '.join(f'{c}' for c in opcionals if c in cols_existentes)
    select_base  = 'id, league, home, away, kickoff, line, over_odds, under_odds, modelo_pred, edge_signal, detected_at, updated_at'
    select = select_base + (f', {select_extra}' if select_extra else '')
    rows = con.execute(f'SELECT {select} FROM lancamentos_lines ORDER BY kickoff DESC').fetchall()
    con.close()
    # Garantir que todas as chaves opcionais existem (None se ausentes)
    result = []
    for r in rows:
        d = dict(r)
        for c in opcionals:
            d.setdefault(c, None)
        result.append(d)
    return result

def formatar(v, casas=1):
    if v is None: return '—'
    try: return f'{float(v):.{casas}f}'
    except: return str(v)

def fmt_dt(s):
    if not s: return '—'
    try:
        dt = datetime.fromisoformat(s)
        return dt.strftime('%d/%m %H:%M')
    except: return s

def resultado_class(r):
    if r == 'GREEN': return 'green'
    if r == 'RED':   return 'red'
    if r == 'Void':  return 'void'
    return ''

def gerar():
    rows = ler_dados()
    total   = len(rows)
    fechados = [r for r in rows if r.get('resultado')]
    greens  = sum(1 for r in fechados if r['resultado'] == 'GREEN')
    reds    = sum(1 for r in fechados if r['resultado'] == 'RED')
    win_rate = f'{greens/len(fechados)*100:.0f}%' if fechados else '—'

    # ROI simples
    roi_val = None
    if fechados:
        lucro = 0
        n = 0
        for r in fechados:
            if r['resultado'] == 'GREEN':
                edge = (r.get('edge_signal') or '').upper()
                odd = r.get('over_odds') if 'OVER' in edge else r.get('under_odds')
                if odd:
                    lucro += float(odd) - 1
                    n += 1
                else:
                    lucro += 0.85  # assume odd ~1.85 se não disponível
                    n += 1
            elif r['resultado'] == 'RED':
                lucro -= 1
                n += 1
        roi_val = f'{lucro/n*100:.1f}%' if n else '—'

    linhas_html = ''
    for r in rows:
        res   = r.get('resultado') or ''
        cls   = resultado_class(res)
        real  = formatar(r.get('resultado_real'), 0) if r.get('resultado_real') is not None else '—'
        casa  = formatar(r.get('lanc_real_casa'), 0)
        fora  = formatar(r.get('lanc_real_fora'), 0)
        detail = f'{casa}+{fora}' if casa != '—' else ''

        linhas_html += f'''
        <tr class="res-{cls}">
          <td>{r["league"]}</td>
          <td class="equipa">{r["home"]}</td>
          <td class="equipa">{r["away"]}</td>
          <td>{fmt_dt(r["kickoff"])}</td>
          <td class="num">{formatar(r["line"])}</td>
          <td class="num">{formatar(r["modelo_pred"])}</td>
          <td class="signal">{r["edge_signal"] or "—"}</td>
          <td class="num">{formatar(r["over_odds"], 2)}</td>
          <td class="num">{formatar(r["under_odds"], 2)}</td>
          <td class="num real">{real}<span class="detail">{detail}</span></td>
          <td class="resultado {cls}">{res or "Pendente"}</td>
        </tr>'''

    agora = datetime.now().strftime('%d/%m/%Y %H:%M')
    html = f'''<!DOCTYPE html>
<html lang="pt">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Lançamentos — Dashboard</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', sans-serif; background: #0f1117; color: #e0e0e0; padding: 20px; }}
  h1 {{ font-size: 1.4rem; margin-bottom: 4px; color: #fff; }}
  .sub {{ font-size: 0.8rem; color: #666; margin-bottom: 20px; }}
  .stats {{ display: flex; gap: 12px; margin-bottom: 24px; flex-wrap: wrap; }}
  .stat {{ background: #1a1d27; border: 1px solid #2a2d3a; border-radius: 8px;
           padding: 14px 20px; min-width: 110px; }}
  .stat .label {{ font-size: 0.7rem; color: #888; text-transform: uppercase; letter-spacing: .05em; }}
  .stat .val {{ font-size: 1.6rem; font-weight: 700; margin-top: 2px; color: #fff; }}
  .stat.green .val {{ color: #4caf50; }}
  .stat.red .val {{ color: #f44336; }}
  .stat.roi .val {{ color: #64b5f6; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 0.82rem; }}
  thead th {{ background: #1a1d27; color: #aaa; font-weight: 600; padding: 10px 8px;
              text-align: left; border-bottom: 1px solid #2a2d3a; white-space: nowrap; }}
  tbody tr {{ border-bottom: 1px solid #1e2130; transition: background .15s; }}
  tbody tr:hover {{ background: #1a1d27; }}
  td {{ padding: 9px 8px; vertical-align: middle; }}
  td.num {{ text-align: right; font-variant-numeric: tabular-nums; }}
  td.equipa {{ font-weight: 500; }}
  td.signal {{ color: #90caf9; font-size: 0.75rem; font-weight: 600; }}
  td.real {{ color: #fff; font-weight: 600; }}
  .detail {{ color: #555; font-size: 0.72rem; margin-left: 4px; }}
  td.resultado {{ font-weight: 700; font-size: 0.8rem; }}
  td.resultado.green {{ color: #4caf50; }}
  td.resultado.red   {{ color: #f44336; }}
  td.resultado.void  {{ color: #888; }}
  tr.res-green {{ background: rgba(76,175,80,.04); }}
  tr.res-red   {{ background: rgba(244,67,54,.04); }}
  .badge {{ display: inline-block; padding: 2px 6px; border-radius: 4px;
            font-size: 0.7rem; font-weight: 700; }}
  .badge.PL  {{ background: #3700b3; color: #bb86fc; }}
  .badge.PPL {{ background: #005c2b; color: #69f0ae; }}
  .badge.BRA {{ background: #004d40; color: #64ffda; }}
  footer {{ margin-top: 20px; font-size: 0.75rem; color: #555; }}
</style>
</head>
<body>
<h1>⚽ Lançamentos Laterais — Dashboard</h1>
<div class="sub">Gerado em {agora} · {total} entradas · <a href="lançamentos_registo.html" style="color:#64b5f6">Ver registo HTML →</a></div>

<div class="stats">
  <div class="stat"><div class="label">Total</div><div class="val">{total}</div></div>
  <div class="stat"><div class="label">Fechados</div><div class="val">{len(fechados)}</div></div>
  <div class="stat green"><div class="label">GREEN</div><div class="val">{greens}</div></div>
  <div class="stat red"><div class="label">RED</div><div class="val">{reds}</div></div>
  <div class="stat"><div class="label">Win Rate</div><div class="val">{win_rate}</div></div>
  <div class="stat roi"><div class="label">ROI (u)</div><div class="val">{roi_val or "—"}</div></div>
</div>

<table>
<thead>
  <tr>
    <th>Liga</th><th>Casa</th><th>Fora</th><th>KO</th>
    <th>Linha</th><th>Modelo</th><th>Signal</th>
    <th>Odd O</th><th>Odd U</th><th>Real</th><th>Resultado</th>
  </tr>
</thead>
<tbody>
{linhas_html}
</tbody>
</table>

<footer>DB: {DB_PATH} · Corre <code>python gerar_dashboard.py</code> para actualizar</footer>
</body>
</html>'''

    OUT_PATH.write_text(html, encoding='utf-8')
    print(f'Dashboard gerado: {OUT_PATH}')
    print(f'  {total} jogos | {greens} GREEN | {reds} RED | win rate {win_rate}')

if __name__ == '__main__':
    gerar()

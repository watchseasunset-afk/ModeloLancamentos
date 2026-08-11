"""
Modelo de Previsão — Lançamentos Laterais por Jogo
====================================================
Suporta: PL 2025/26 | PPL 2025/26 | BRA 2025

Uso:
    python modelo_previsao.py                              → menu interativo
    python modelo_previsao.py "Liverpool" "Arsenal"        → PL (default)
    python modelo_previsao.py "Benfica" "Porto" --liga PPL
    python modelo_previsao.py "Flamengo" "Palmeiras" 36.5 --liga BRA
    python modelo_previsao.py --retreinar                  → força re-treino
"""
import sys, re
sys.path.insert(0, '/sessions/charming-hopeful-davinci/.local/lib/python3.10/site-packages')
for pkg in ['pandas','openpyxl','scipy','scikit-learn','numpy']:
    try:
        __import__(pkg.replace('scikit-learn','sklearn'))
    except ImportError:
        import subprocess
        subprocess.check_call([sys.executable,'-m','pip','install',pkg,'--break-system-packages'])

import numpy as np
import pandas as pd
from scipy import stats
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import mean_absolute_error, r2_score
from sklearn.model_selection import cross_val_score
from pathlib import Path
import pickle, warnings
warnings.filterwarnings('ignore')

DATA_DIR = Path(__file__).parent / 'data'
BASE_DIR = Path(__file__).parent

# ── Configuração por liga ───────────────────────────────────────────────────────
LIGAS_CFG = {
    'PL': {
        'nome':      'Premier League 2025/26',
        'stats':     DATA_DIR / 'stats_completas_PL_2025_26.xlsx',
        'teams_map': BASE_DIR / 'teams_map.csv',
        'model_pkl': DATA_DIR / 'modelo_lancamentos_PL.pkl',
    },
    'PPL': {
        'nome':      'Primeira Liga 2025/26',
        'stats':     DATA_DIR / 'stats_completas_PPL_2025_26.xlsx',
        'teams_map': BASE_DIR / 'teams_map_ppl.csv',
        'model_pkl': DATA_DIR / 'modelo_lancamentos_PPL.pkl',
    },
    'BRA': {
        'nome':      'Brasileirão Série A 2025',
        'stats':     DATA_DIR / 'stats_completas_BRA_2025.xlsx',
        'teams_map': BASE_DIR / 'teams_map_brazil.csv',
        'model_pkl': DATA_DIR / 'modelo_lancamentos_BRA.pkl',
    },
}

# ── Features ────────────────────────────────────────────────────────────────────
HOME_FEATS = [
    'Alívios_casa',
    'Passes longos_casa',
    'Cruzamentos_casa',
    'Posse de bola_casa',
    'Lançamentos_casa',
    'Faltas_fora',
    'Interceções_fora',
    'Desarmes_fora',
]
AWAY_FEATS = [
    'Alívios_fora',
    'Passes longos_fora',
    'Cruzamentos_fora',
    'Posse de bola_fora',
    'Lançamentos_fora',
    'Faltas_casa',
    'Interceções_casa',
    'Desarmes_casa',
]

# ── Parsing ─────────────────────────────────────────────────────────────────────
def parse_stat(val):
    if pd.isna(val): return np.nan
    s = str(val).strip()
    m = re.search(r'\((\d+)/(\d+)\)', s)
    if m: return float(m.group(2))
    try: return float(s.replace('%','').strip())
    except: return np.nan

# ════════════════════════════════════════════════════════════════════════════════
def carregar_dados(liga='PL'):
    cfg = LIGAS_CFG[liga]
    df    = pd.read_excel(cfg['stats'])
    teams = pd.read_csv(cfg['teams_map'])
    df    = df.merge(teams, on='mid', how='left')
    for c in df.columns:
        if c not in ('mid','home','away'):
            df[c] = df[c].apply(parse_stat)
    return df

def calcular_medias_equipa(df):
    all_stats = list(set(HOME_FEATS + AWAY_FEATS))
    medias = {}
    for team in sorted(df['home'].dropna().unique()):
        mh = df['home'] == team
        mf = df['away'] == team
        medias[team] = {}
        for stat in all_stats:
            if stat.endswith('_casa'):
                medias[team][stat] = df.loc[mh, stat].mean()
            else:
                medias[team][stat] = df.loc[mf, stat].mean()
    return medias

def construir_dataset_treino(df, medias):
    rows = []
    for _, r in df.iterrows():
        home, away = r['home'], r['away']
        if pd.isna(home) or pd.isna(away): continue
        if home not in medias or away not in medias: continue
        feats = {}
        for stat in HOME_FEATS:
            feats[f'H_{stat}'] = medias[home].get(stat, np.nan)
        for stat in AWAY_FEATS:
            feats[f'A_{stat}'] = medias[away].get(stat, np.nan)
        total = r.get('Lançamentos_casa', np.nan) + r.get('Lançamentos_fora', np.nan)
        if pd.isna(total): continue
        feats['target'] = total
        feats['home']   = home
        feats['away']   = away
        rows.append(feats)
    return pd.DataFrame(rows)

def treinar_modelo(df_treino, liga):
    feat_cols = [c for c in df_treino.columns if c.startswith('H_') or c.startswith('A_')]
    X = df_treino[feat_cols].values
    y = df_treino['target'].values
    mask = ~np.isnan(X).any(axis=1) & ~np.isnan(y)
    X, y = X[mask], y[mask]

    scaler = StandardScaler()
    X_sc   = scaler.fit_transform(X)
    lr     = LinearRegression()
    lr.fit(X_sc, y)

    cv_scores = cross_val_score(lr, X_sc, y, cv=5, scoring='neg_mean_absolute_error')
    mae_cv    = -cv_scores.mean()
    y_pred    = lr.predict(X_sc)
    r2        = r2_score(y, y_pred)
    std_resid = (y - y_pred).std()

    print(f"  [{liga}] R²: {r2:.3f} | MAE CV: {mae_cv:.2f} | σ resíduos: {std_resid:.2f} | N: {len(y)}")
    return {
        'model':         lr,
        'scaler':        scaler,
        'feat_cols':     feat_cols,
        'std_resid':     std_resid,
        'mae_cv':        mae_cv,
        'r2':            r2,
        'medias_global': {'total': float(np.mean(y)), 'std': float(np.std(y))},
    }

# ════════════════════════════════════════════════════════════════════════════════
def carregar_pacote(liga='PL'):
    cfg = LIGAS_CFG[liga]
    if cfg['model_pkl'].exists():
        with open(cfg['model_pkl'], 'rb') as f:
            return pickle.load(f)
    return _treinar_e_guardar(liga)

def _treinar_e_guardar(liga='PL'):
    print(f"A treinar modelo [{liga}]...")
    df     = carregar_dados(liga)
    medias = calcular_medias_equipa(df)
    df_tr  = construir_dataset_treino(df, medias)
    info   = treinar_modelo(df_tr, liga)
    pacote = {**info, 'medias': medias, 'liga': liga}
    cfg    = LIGAS_CFG[liga]
    with open(cfg['model_pkl'], 'wb') as f:
        pickle.dump(pacote, f)
    print(f"  Modelo [{liga}] guardado.\n")
    return pacote

def retreinar(liga=None):
    """Força re-treino. liga=None → todas as ligas."""
    ligas = [liga] if liga else list(LIGAS_CFG.keys())
    pacotes = {}
    for l in ligas:
        cfg = LIGAS_CFG[l]
        if cfg['model_pkl'].exists(): cfg['model_pkl'].unlink()
        pacotes[l] = _treinar_e_guardar(l)
    return pacotes[ligas[0]] if len(ligas) == 1 else pacotes

# ════════════════════════════════════════════════════════════════════════════════
def prever(home_team, away_team, linha_bookie=None, liga='PL', pacote=None, verbose=True):
    if pacote is None:
        pacote = carregar_pacote(liga)

    model     = pacote['model']
    scaler    = pacote['scaler']
    feat_cols = pacote['feat_cols']
    medias    = pacote['medias']
    std_r     = pacote['std_resid']
    mu_glob   = pacote['medias_global']['total']
    liga_nome = LIGAS_CFG.get(pacote.get('liga', liga), {}).get('nome', liga)

    home_team = _match_team(home_team, list(medias.keys()))
    away_team = _match_team(away_team, list(medias.keys()))
    if home_team is None or away_team is None:
        return None

    feats = {}
    for stat in HOME_FEATS:
        feats[f'H_{stat}'] = medias[home_team].get(stat, np.nan)
    for stat in AWAY_FEATS:
        feats[f'A_{stat}'] = medias[away_team].get(stat, np.nan)

    x = np.array([feats.get(c, np.nan) for c in feat_cols]).reshape(1,-1)
    if np.isnan(x).any():
        x = np.nan_to_num(x, nan=np.nanmean(x))

    x_sc = scaler.transform(x)
    pred = max(float(model.predict(x_sc)[0]), 15)

    ci68_lo = pred - std_r;       ci68_hi = pred + std_r
    ci90_lo = pred - 1.645*std_r; ci90_hi = pred + 1.645*std_r

    prob_over = prob_under = None
    if linha_bookie is not None:
        prob_over  = float(1 - stats.norm.cdf(linha_bookie, loc=pred, scale=std_r))
        prob_under = float(stats.norm.cdf(linha_bookie, loc=pred, scale=std_r))

    diff       = (pred - linha_bookie) if linha_bookie is not None else None
    edge_signal = None
    if diff is not None:
        if abs(diff) > 4:   edge_signal = 'OVER FORTE' if diff > 0 else 'UNDER FORTE'
        elif abs(diff) > 2: edge_signal = 'OVER MODERADO' if diff > 0 else 'UNDER MODERADO'
        else:               edge_signal = 'sem edge'

    resultado = {
        'liga':        pacote.get('liga', liga),
        'home':        home_team,
        'away':        away_team,
        'previsao':    round(pred, 1),
        'ci68':        (round(ci68_lo,1), round(ci68_hi,1)),
        'ci90':        (round(ci90_lo,1), round(ci90_hi,1)),
        'media_liga':  round(mu_glob, 1),
        'linha':       linha_bookie,
        'prob_over':   round(prob_over*100,1) if prob_over is not None else None,
        'prob_under':  round(prob_under*100,1) if prob_under is not None else None,
        'lanc_H_avg':  round(medias[home_team].get('Lançamentos_casa', 0), 1),
        'lanc_A_avg':  round(medias[away_team].get('Lançamentos_fora', 0), 1),
        'edge_signal': edge_signal,
    }

    if verbose:
        _print_resultado(resultado, liga_nome)
    return resultado

# ════════════════════════════════════════════════════════════════════════════════
def _print_resultado(r, liga_nome=''):
    sep = "─" * 58
    print(f"\n{sep}")
    print(f"  [{liga_nome}]  {r['home']}  🏠 vs ✈  {r['away']}")
    print(sep)
    print(f"  {'Previsão total lançamentos':<33} {r['previsao']:>6.1f}")
    print(f"  {'IC 68%':<33} {r['ci68'][0]:>5.1f} – {r['ci68'][1]:.1f}")
    print(f"  {'IC 90%':<33} {r['ci90'][0]:>5.1f} – {r['ci90'][1]:.1f}")
    print(f"  {'Média liga':<33} {r['media_liga']:>6.1f}")
    print(f"  {'Média {h} (casa)'.format(h=r['home']):<33} {r['lanc_H_avg']:>6.1f}")
    print(f"  {'Média {a} (fora)'.format(a=r['away']):<33} {r['lanc_A_avg']:>6.1f}")
    if r['linha'] is not None:
        print(sep)
        diff = r['previsao'] - r['linha']
        emoji = '🟢' if abs(diff)>3 else '🟡' if abs(diff)>1.5 else '⚪'
        print(f"  Linha: {r['linha']:.1f}  →  {emoji} {r['edge_signal']}  ({diff:+.1f})")
        print(f"  P(Over {r['linha']:.1f}):  {r['prob_over']:>5.1f}%")
        print(f"  P(Under {r['linha']:.1f}): {r['prob_under']:>5.1f}%")
    print(sep)

def _match_team(name, team_list):
    name = name.strip()
    if name in team_list: return name
    for t in team_list:
        if t.lower() == name.lower(): return t
    matches = [t for t in team_list if name.lower() in t.lower() or t.lower() in name.lower()]
    if len(matches) == 1: return matches[0]
    if matches:
        print(f"  '{name}' → usando: {matches[0]}")
        return matches[0]
    print(f"  Equipa '{name}' não encontrada. Disponíveis: {sorted(team_list)}")
    return None

# ════════════════════════════════════════════════════════════════════════════════
def prever_jornada(jogos, linha_default=None, liga='PL', pacote=None):
    if pacote is None: pacote = carregar_pacote(liga)
    rows = []
    for jogo in jogos:
        h, a = jogo[0], jogo[1]
        l = jogo[2] if len(jogo) >= 3 else linha_default
        r = prever(h, a, l, liga=liga, pacote=pacote, verbose=False)
        if r: rows.append(r)
    df = pd.DataFrame(rows)
    print(f"\n{'─'*82}")
    print(f"  {'Casa':<20} {'Fora':<20} {'Prev':>6} {'CI68':>10} {'Linha':>7} {'Δ':>6} {'Signal'}")
    print(f"{'─'*82}")
    for _, r in df.iterrows():
        diff  = (r['previsao'] - r['linha']) if r['linha'] else None
        linha = f"{r['linha']:.1f}" if r['linha'] else '  —  '
        delta = f"{diff:+.1f}" if diff is not None else '  —  '
        sig   = r.get('edge_signal', '') or ''
        print(f"  {r['home']:<20} {r['away']:<20} {r['previsao']:>6.1f} "
              f"  {r['ci68'][0]:.0f}–{r['ci68'][1]:.0f}  {linha:>7} {delta:>6}  {sig}")
    print(f"{'─'*82}")
    return df

# ════════════════════════════════════════════════════════════════════════════════
# Interface usada pelo monitor_lancamentos.py
def carregar_modelo(liga='PL'):
    """Alias para carregar_pacote — compatibilidade com monitor."""
    return carregar_pacote(liga)

def prever_jogo(home, away, linha=None, liga='PL', pacote=None):
    """Alias simplificado para uso no monitor."""
    return prever(home, away, linha, liga=liga, pacote=pacote, verbose=False)

def calcular_edge(previsao, linha):
    if previsao is None or linha is None: return 'sem edge'
    diff = previsao - linha
    if abs(diff) > 4:   return 'OVER FORTE' if diff > 0 else 'UNDER FORTE'
    if abs(diff) > 2:   return 'OVER MODERADO' if diff > 0 else 'UNDER MODERADO'
    return 'sem edge'

# ════════════════════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    liga_arg = 'PL'
    if '--liga' in sys.argv:
        idx = sys.argv.index('--liga')
        liga_arg = sys.argv[idx+1].upper()
        sys.argv.pop(idx); sys.argv.pop(idx)

    if '--retreinar' in sys.argv:
        sys.argv.remove('--retreinar')
        if liga_arg == 'ALL':
            retreinar()
        else:
            retreinar(liga_arg)

    pacote = carregar_pacote(liga_arg)
    args   = [a for a in sys.argv[1:] if not a.startswith('--')]

    if len(args) >= 2:
        home  = args[0]
        away  = args[1]
        linha = float(args[2]) if len(args) >= 3 else None
        prever(home, away, linha, liga=liga_arg, pacote=pacote)
    else:
        # Menu interativo
        teams = sorted(pacote['medias'].keys())
        cfg   = LIGAS_CFG[liga_arg]
        print("\n" + "═"*58)
        print(f"  MODELO LANÇAMENTOS — {cfg['nome']}")
        print("═"*58)
        print(f"  Ligas disponíveis: {list(LIGAS_CFG.keys())}  (--liga PPL / BRA)")
        print(f"  Equipas disponíveis [{liga_arg}] ({len(teams)}):")
        for i, t in enumerate(teams, 1):
            print(f"    {i:2}. {t}")
        print("\n  Exemplo: Benfica Porto 36.5")
        print("  Exemplo jornada: 'jornada'\n")

        entrada = input("  > ").strip()
        partes  = entrada.split()

        if entrada.lower() == 'jornada':
            jogos = []
            print("  Insere jogos (Casa Fora [Linha]) — linha vazia termina:")
            while True:
                lj = input("  + ").strip()
                if not lj: break
                p = lj.split()
                if len(p) >= 2:
                    l = float(p[2]) if len(p) >= 3 else None
                    jogos.append((p[0], p[1], l) if l else (p[0], p[1]))
            prever_jornada(jogos, liga=liga_arg, pacote=pacote)
        else:
            def resolve(s):
                try:
                    i = int(s)-1
                    return teams[i] if 0 <= i < len(teams) else s
                except: return s
            home  = resolve(partes[0]) if partes else None
            away  = resolve(partes[1]) if len(partes) > 1 else None
            linha = float(partes[2]) if len(partes) > 2 else None
            if home and away:
                prever(home, away, linha, liga=liga_arg, pacote=pacote)

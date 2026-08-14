"""
debug_22bet_live.py
Inspeciona o payload completo do GetGameZip para um jogo ao vivo na 22bet.
Uso: python debug_22bet_live.py [CI_do_jogo]
Se não der CI, tenta encontrar PSG vs Aston Villa automaticamente.
"""
import json, urllib.request, sys, re
from datetime import datetime

BET_BASE = 'https://22bet4me.com/service-api/LineFeed'
HDR = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Accept': 'application/json, text/plain, */*',
    'Referer': 'https://22bet4me.com/line/football/',
}

def api(url):
    req = urllib.request.Request(url, headers=HDR)
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read())

def search_game(home_kw, away_kw):
    """Procura jogo ao vivo pelo nome das equipas."""
    print(f'[SEARCH] A procurar {home_kw!r} vs {away_kw!r} nos jogos ao vivo...')
    url = (f'{BET_BASE}/GetSports?sports=1&lng=pt_PT&partner=151'
           f'&country=148&fcountry=148&isNewBuilder=true&grMode=4')
    try:
        data = api(url)
        # Estrutura: Value → champs → jogos
        champs = data.get('Value', [])
        for champ in champs:
            for game in champ.get('G', []):
                home = game.get('HN', '') or ''
                away = game.get('AN', '') or ''
                if home_kw.lower() in home.lower() or away_kw.lower() in away.lower():
                    print(f'  ✅ Encontrado: {home} vs {away} — CI={game.get("CI")}')
                    return game.get('CI')
    except Exception as e:
        print(f'  ❌ Erro em GetSports: {e}')

    # Fallback: GetChampsZip com futebol ao vivo
    print('[SEARCH] A tentar GetChampsZip...')
    try:
        url2 = (f'{BET_BASE}/GetChampsZip?sports=1&lng=pt_PT&partner=151'
                f'&country=148&fcountry=148&isNewBuilder=true&grMode=4&mode=4')
        data2 = api(url2)
        for item in data2.get('Value', []):
            for g in item.get('G', []):
                home = g.get('HN',''); away = g.get('AN','')
                if home_kw.lower() in home.lower() or away_kw.lower() in away.lower():
                    print(f'  ✅ Encontrado: {home} vs {away} — CI={g.get("CI")}')
                    return g.get('CI')
    except Exception as e:
        print(f'  ❌ Erro em GetChampsZip: {e}')
    return None

def inspect_game(ci):
    print(f'\n[GAMZIP] CI={ci}')
    url = (f'{BET_BASE}/GetGameZip?id={ci}&lng=pt_PT&tzo=1&cfview=0'
           f'&isSubGames=true&GroupEvents=true&countevents=250'
           f'&country=148&fcountry=148&isNewBuilder=true&partner=151&grMode=4')
    data = api(url)
    val  = data.get('Value', {})

    # ── Info básica do jogo ──────────────────────────────────────────────────
    print(f'\n=== INFO JOGO ===')
    for k in ['HN','AN','SC','Timer','S','PS','TN','CI']:
        if k in val:
            print(f'  {k}: {val[k]}')

    # ── Score / Stats ao vivo ────────────────────────────────────────────────
    print(f'\n=== SCORE / STATS ===')
    score = val.get('SC') or val.get('Score') or val.get('LS')
    print(f'  Score (SC): {score}')
    timer = val.get('Timer') or val.get('T')
    print(f'  Timer: {timer}')

    # Procura campos de stats
    stat_keys = [k for k in val.keys() if any(x in k.upper() for x in
                 ['STAT','CORNER','FOUL','THROW','SHOT','ATTACK','POSS','YELLOW'])]
    if stat_keys:
        print(f'  Stats encontradas: {stat_keys}')
        for k in stat_keys:
            print(f'    {k}: {val[k]}')
    else:
        print('  ⚠ Nenhum campo de stats óbvio no nível raiz')

    # ── Sub-jogos (TIs disponíveis) ──────────────────────────────────────────
    sg = val.get('SG', [])
    print(f'\n=== SUB-JOGOS (SG) — {len(sg)} entradas ===')
    for e in sg[:60]:
        ti   = e.get('TI','?')
        name = e.get('N') or e.get('NA','?')
        ci2  = e.get('CI','?')
        print(f'  TI={ti:>4}  CI={ci2}  "{name}"')

    # ── GE (odds) ────────────────────────────────────────────────────────────
    ge = val.get('GE', [])
    if ge:
        print(f'\n=== GE (odds grupos) — {len(ge)} grupos ===')
        for g in ge[:5]:
            print(f'  G={g.get("G")}  N={g.get("N","?")}  entries={len(g.get("E",[]))}')

    # ── Payload completo (para análise manual) ───────────────────────────────
    out = f'payload_22bet_{ci}_{datetime.now().strftime("%H%M%S")}.json'
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f'\n💾 Payload completo guardado em: {out}')

    return val, sg

# ── Main ─────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    if len(sys.argv) > 1:
        ci = sys.argv[1]
        print(f'CI fornecido: {ci}')
    else:
        ci = search_game('PSG', 'Aston Villa') or search_game('Paris', 'Villa')

    if not ci:
        print('\n❌ Jogo não encontrado. Tenta:')
        print('   1. Confirmar que o jogo está ao vivo na 22bet')
        print('   2. Correr: python debug_22bet_live.py <CI_do_jogo>')
        sys.exit(1)

    inspect_game(ci)

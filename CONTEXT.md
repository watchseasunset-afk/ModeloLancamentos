# ModLancamentos — Contexto do Projeto

## Arquitectura

Dois projetos separados:
- **Faltas** (`C:\Claude_Mod_Faltas\Faltas_206_2027\faltas_app_v5\faltas_app`) — READ ONLY. Gera `model_data.json` com equipas, médias de faltas e árbitros.
- **ModLancamentos** (`C:\ModLancamentos`) — Flask/Railway. Lê `model_data.json` do Faltas como fonte master para faltas e árbitros. Lê Excel próprio (`data/medias_equipas_MULTI_LIGA.xlsx`) para lançamentos.

**REGRA CRÍTICA**: Nunca modificar nada no projeto Faltas.

## Deploy

- Railway: `https://lancamentos-laterais.up.railway.app`
- Git push → deploy automático
- Filesystem efémero — dados persistem em SQLite (`data/lancamentos_lines.db`, `data/live_stats.db`)

## Ficheiros-chave

| Ficheiro | Função |
|---|---|
| `app.py` | App principal Flask + APScheduler |
| `model_data.json` | Cópia do Faltas para Railway (bundled) — atualizar quando o Faltas gera novo |
| `data/medias_equipas_MULTI_LIGA.xlsx` | Médias de lançamentos por equipa/liga |
| `repush.ps1` | Push manual dos 3 jogos LEQ de hoje (Anderlecht/PAOK, Hearts/Benfica, Dinamo Minsk/Braga) |
| `repush_statscore.ps1` | Versão com variáveis para statscore_id (preencher via Yonibet) |

## Estrutura de dados em app.py

```python
_MEDIAS   # {norm_nome: {liga, equipa, faltas_media, ffh, fsh, ffa, fsa, jogos, lanc_media, posse_media}}
_ARBITROS # {'PPL': {norm_nome: {nome, jogos, media_faltas, last_totals}}, 'ESP': {...}, 'BRA': {...}}

_JSON_LIGA_MAP = {'Tugao':'PPL', 'Spain':'ESP', 'Brazil':'BRA', 'UK':'UK', 'Italia':'ITA', 'France':'FRA'}
```

## Fontes de dados (startup)

```python
_load_medias()       # faltas das equipas — lê model_data.json (Faltas master)
_load_arbitros()     # árbitros — lê model_data.json (Faltas master)
_load_lanc_medias()  # lançamentos — lê Excel ModLancamentos, enriquece _MEDIAS
```

## model_data.json — estrutura

```json
{
  "teams": {"Tugao": [{"name", "ffh", "fsh", "ffa", "fsa", "games_home", "games_away"}]},
  "refs":  {"Tugao": [{"name", "n_games", "avg_total", "last_totals"}]},
  "league_avg": {"Tugao": 26.78},
  "coefs": {...}
}
```

## Matching árbitros (_get_arbitro_stats)

Word-based fuzzy: "Bessa J." → word "bessa" → encontra "jose bessa".
Funciona para PPL, ESP, BRA.

## Popup JOGOS — dados mostrados

- **Equipas**: FF Casa (FFH) / FF Fora (FFA), FS Casa/Fora, Lanç/jogo, Posse
- **Árbitro**: nome, N jogos · avg X.XX (fonte maior, branco), last_totals tiles (verde=OVER avg árbitro, azul=UNDER)
- **Previsão**: totLanc (laranja) + faltFinal (roxo)
- Faltas contextuais: FFH para equipa da casa, FFA para equipa de fora

## Modelo de previsão de faltas (live)

```
combined = 0.4 * reg_pred + 0.4 * (FFH + FFA) + 0.2 * ref_pred
```
Coefs por liga em `_FALTAS_COEFS` (PPL, ESP, BRA, PL, SCOPA).

## Janela live

- Recolhe jogos com `kickoff <= now + 3h` (pré-jogo) e `kickoff >= now - 115min` (não terminados)
- Ciclo: 30 segundos
- Status: `pending` → `live` (quando kickoff passou) → `done` (>115min após kickoff)

## Jogos LEQ hoje (2026-08-13)

| Jogo | bet_ci | statscore_id | Kickoff UTC |
|---|---|---|---|
| Anderlecht vs PAOK | 358361500 | PREENCHER via Yonibet | 19:30 |
| Hearts vs Benfica | 357962992 | PREENCHER via Yonibet | 19:45 |
| Dinamo Minsk vs SC Braga | 358387744 | PREENCHER via Yonibet | 18:00 |

**Para encontrar statscore_id**: Yonibet → clicar no jogo live → DevTools F12 → Network → filtrar "statscore" → copiar eventId da URL.

## Debug endpoints

- `GET /debug/arbitro?nome=X&liga=PPL` — testa lookup árbitro
- `GET /debug/medias` — mostra `_ARBITROS` e `_MEDIAS`

## Pendentes

- [ ] statscore_id dos 3 jogos LEQ (encontrar no Yonibet quando ficarem live)
- [ ] Deploy das últimas alterações ao Railway (lançamentos no popup, fonte maior árbitro)

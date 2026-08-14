# Resumo Sessão — 14 Agosto 2026

## Objetivo
Corrigir ausência de médias de lançamentos no popup live (PPL) para Vitória Guimarães, Marítimo e Académico Viseu.

---

## O que foi feito

### 1. Vitória Guimarães
- **Problema**: `teams_map_ppl.csv` só tem Guimarães como "away" (nunca home) → `calc_medias_multi_liga.py` usava só `df['home'].unique()` → excluía Guimarães.
- **Fix `calc_medias_multi_liga.py`**: `teams_list = sorted(set(df['home'] + df['away']))` (home ∪ away).
- **Dados recolhidos** via Chrome/Flashscore PPL 2025/26 (Vitória SC = nome no Flashscore):
  - 5H + 6A = 11 jogos, **avg = 19.91**
- **Excel**: `Vitoria Guimaraes` → `lanc_media=19.91`, `Jogos_Total=11`
- **Alias correto** já existente: `'vitoria sc': 'guimaraes'` em `_LANC_ALIASES` e `_TSDB_ALIASES`

### 2. Marítimo
- **Situação**: Marítimo estava na Liga 2 2025/26 (não no PPL) → sem dados no `stats_completas_PPL_2025_26.xlsx`.
- **Dados recolhidos** via Chrome/Flashscore Liga 2 2025/26:
  - 5H + 6A = 11 jogos, **avg = 23.45**
  - Home: [20, 19, 30, 24, 17] → avg 22.00
  - Away: [22, 30, 19, 22, 24, 31] → avg 24.67
- **Excel**: `Maritimo` adicionado → `lanc_media=23.45`, `Jogos_Total=11`
- **Alias**: `'maritimo'` bate direto em `_MEDIAS['maritimo']` (sem alias necessário)

### 3. Académico Viseu
- **Situação**: Subiu da Liga 2 2025/26 para PPL 2026/27. AFS ≠ AVS ≠ Académico Viseu (equipas distintas).
- **Dados recolhidos** via Chrome/Flashscore Liga 2 2025/26 ("Académico Viseu"):
  - 5H + 6A = 11 jogos, **avg = 24.00**
  - Home: [23, 26, 24, 18, 33] → avg 24.80
  - Away: [9, 26, 26, 26, 20, 33] → avg 23.33
- **Excel**: `Academico Viseu` adicionado → `lanc_media=24.00`, `Jogos_Total=11`
- **Alias**: normaliza para `'academico viseu'` → bate em `_MEDIAS['academico viseu']` direto

### 4. AFS / AVS
- **Esclarecimento**: AFS (PPL 2025/26, desceu) ≠ AVS (equipa distinta em model_data) ≠ Académico Viseu (subiu Liga 2)
- **Dados AFS PPL 2025/26** recolhidos: 5H + 6A = 11 jogos, **avg = 19.82**
- **Excel**: `Afs` atualizado → `lanc_media=19.82`, `Jogos_Total=11`
- **Alias**: `'afs': 'avs'` em `_LANC_ALIASES` (mapeia Excel 'Afs' → `_MEDIAS['avs']`)

### 5. Fix `_TSDB_ALIASES` no app.py
- Adicionado `'afs': 'avs'` para resolver lookup live da AVS (3 chars → não apanhado pelo word-length match)

---

## Estado do Excel (`data/medias_equipas_MULTI_LIGA.xlsx` — PPL-Geral)

| Equipa             | Jogos | lanc_media |
|--------------------|-------|------------|
| Vitoria Guimaraes  | 11    | 19.91      |
| Maritimo           | 11    | 23.45      |
| Academico Viseu    | 11    | 24.00      |
| Afs                | 11    | 19.82      |

---

## Pendente

- [ ] **`railway up`** — fazer deploy de todas as alterações (correr no terminal em `C:\ModLancamentos`)
- [ ] Verificar popup live no próximo jogo do PPL com estas equipas
- [ ] (Opcional) Recolher mais jogos do Académico Viseu e Marítimo quando houver mais dados PPL 2026/27
- [ ] (Opcional) Clarificar relação AFS ↔ AVS — confirmar se `'afs': 'avs'` em `_TSDB_ALIASES` está correto para jogos live do AVS

---

## Ficheiros modificados

| Ficheiro | Alteração |
|---|---|
| `app.py` | `_TSDB_ALIASES`: adicionado `'afs': 'avs'` |
| `calc_medias_multi_liga.py` | `teams_list` usa home ∪ away |
| `data/medias_equipas_MULTI_LIGA.xlsx` | +Maritimo, +Academico Viseu, updated Vitoria Guimaraes + Afs |

**Nota**: `C:\Claude_Mod_Faltas\...` é READ ONLY — nunca modificar.

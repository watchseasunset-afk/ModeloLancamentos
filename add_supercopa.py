"""
add_supercopa.py — Regista PSG vs Aston Villa (UEFA Super Cup) na live_stats.db

Uso:
    python add_supercopa.py            # insere com kickoff 20:00 Lisboa (19:00 UTC)
    python add_supercopa.py --status   # mostra estado do jogo na DB
    python add_supercopa.py --remove   # remove da DB (para recomeçar)
"""
import sqlite3, sys, os
from datetime import datetime, timezone
from pathlib import Path

BASE     = Path(__file__).parent
DATA_DIR = Path(os.environ.get('DATA_DIR', str(BASE)))
DB_PATH  = DATA_DIR / 'live_stats.db'

FLASH_MID = 'GnSGwOrJ'          # Flashscore: Aston Villa vs PSG
LEAGUE    = 'SCOPA'
HOME      = 'Paris SG'           # 22bet: PSG é "casa"
AWAY      = 'Aston Villa'
KICKOFF   = '2026-08-12T19:00:00+00:00'   # 20:00 Lisboa = 19:00 UTC
BET_CI    = '191946'             # 22bet game ID (PSG vs Villa, Super Cup)
LANC_BASE = 37.0                 # média europeia lançamentos
FALT_BASE = 22.0                 # média europeia faltas

def init_db():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    # Garantir coluna bet_ci existe
    try:
        con.execute('ALTER TABLE live_games ADD COLUMN bet_ci TEXT')
        con.commit()
    except: pass
    return con

def insert_game():
    con = init_db()
    now = datetime.now(timezone.utc).isoformat()
    try:
        con.execute('''
            INSERT OR REPLACE INTO live_games
              (flash_mid, league, home, away, kickoff,
               lanc_baseline, faltas_baseline, status, bet_ci, added_at)
            VALUES (?,?,?,?,?,?,?,'pending',?,?)
        ''', (FLASH_MID, LEAGUE, HOME, AWAY, KICKOFF,
              LANC_BASE, FALT_BASE, BET_CI, now))
        con.commit()
        print(f'✅ Jogo registado:')
        print(f'   {HOME} vs {AWAY} ({LEAGUE})')
        print(f'   flash_mid={FLASH_MID}  bet_ci={BET_CI}')
        print(f'   kickoff={KICKOFF}  lanc_base={LANC_BASE}  falt_base={FALT_BASE}')
    except Exception as e:
        print(f'❌ Erro: {e}')
    finally:
        con.close()

def show_status():
    con = init_db()
    row = con.execute('SELECT * FROM live_games WHERE flash_mid=?', (FLASH_MID,)).fetchone()
    if row:
        print('Estado na DB:')
        for k in row.keys():
            print(f'  {k}: {row[k]}')
    else:
        print('⚠ Jogo não encontrado na DB.')
    con.close()

def remove_game():
    con = init_db()
    con.execute('DELETE FROM live_games WHERE flash_mid=?', (FLASH_MID,))
    con.execute('DELETE FROM live_snapshots WHERE flash_mid=?', (FLASH_MID,))
    con.commit()
    print(f'🗑 Jogo {FLASH_MID} removido.')
    con.close()

if __name__ == '__main__':
    if '--status' in sys.argv:
        show_status()
    elif '--remove' in sys.argv:
        remove_game()
    else:
        insert_game()
        print()
        show_status()

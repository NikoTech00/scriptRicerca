"""Script diagnostico TEMPORANEO — da eseguire una volta su Windows con:
    python debug_cv_state.py
Ispeziona il DB reale stato_massivo/ricerca.sqlite per capire perche' i CV
utili sono crollati da 2109 a 336 dopo il bump di DISCOVERY_VERSION, mentre
la coda (IN_CODA) e' ormai completamente drenata (0).

Non modifica nulla: sola lettura. Va rimosso con un commit successivo una
volta ottenuto l'output.
"""
import json
import sqlite3
from pathlib import Path

DB = Path('stato_massivo') / 'ricerca.sqlite'


def main():
    if not DB.exists():
        print(f'DB non trovato: {DB.resolve()}')
        return
    con = sqlite3.connect(str(DB))
    con.row_factory = sqlite3.Row

    def one(sql, params=()):
        return con.execute(sql, params).fetchone()[0]

    print('--- Conteggi grezzi ---')
    print('candidates totali:', one('SELECT COUNT(*) FROM candidates'))
    print('evidence totali:', one('SELECT COUNT(*) FROM evidence'))
    print('urls totali:', one('SELECT COUNT(*) FROM urls'))
    print('urls per stato:')
    for row in con.execute("SELECT state, COUNT(*) n FROM urls GROUP BY state"):
        print(' ', row['state'], row['n'])

    print()
    print('--- candidates per source (top 15) ---')
    for row in con.execute(
        'SELECT source, COUNT(*) n FROM candidates GROUP BY source ORDER BY n DESC LIMIT 15'
    ):
        print(' ', row['source'], row['n'])

    cv_da_profilo_n = one("SELECT COUNT(*) FROM candidates WHERE source='cv_da_profilo'")
    print()
    print("candidates source='cv_da_profilo':", cv_da_profilo_n)

    print()
    print('--- evidence con cv=true (scansione data JSON) ---')
    cv_true = 0
    cv_true_by_source = {}
    cv_true_identity = {}
    total_evidence = 0
    for row in con.execute('SELECT e.pid,e.url,e.data,c.source FROM evidence e LEFT JOIN candidates c ON c.pid=e.pid AND c.url=e.url'):
        total_evidence += 1
        try:
            data = json.loads(row['data'])
        except (TypeError, ValueError):
            continue
        if data.get('cv'):
            cv_true += 1
            src = row['source'] or '(nessun candidate corrispondente!)'
            cv_true_by_source[src] = cv_true_by_source.get(src, 0) + 1
            ident = data.get('identity')
            cv_true_identity[ident] = cv_true_identity.get(ident, 0) + 1
    print('evidence totali scansionate:', total_evidence)
    print('evidence con cv=true:', cv_true)
    print('per source:')
    for src, n in sorted(cv_true_by_source.items(), key=lambda kv: -kv[1]):
        print(' ', src, n)
    print('per identity:')
    for ident, n in sorted(cv_true_identity.items(), key=lambda kv: -kv[1]):
        print(' ', ident, n)

    print()
    print('--- cv_da_profilo: candidates senza evidence corrispondente ---')
    orphan = one(
        "SELECT COUNT(*) FROM candidates c WHERE c.source='cv_da_profilo' "
        "AND NOT EXISTS (SELECT 1 FROM evidence e WHERE e.pid=c.pid AND e.url=c.url)"
    )
    print('candidates cv_da_profilo SENZA riga evidence:', orphan, '/', cv_da_profilo_n)

    print()
    print('--- cv_da_profilo: stato delle relative urls ---')
    for row in con.execute(
        "SELECT u.state, COUNT(*) n FROM candidates c JOIN urls u ON u.url=c.url "
        "WHERE c.source='cv_da_profilo' GROUP BY u.state"
    ):
        print(' ', row['state'], row['n'])

    con.close()


if __name__ == '__main__':
    main()

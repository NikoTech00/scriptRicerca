#!/usr/bin/env python3
"""Genera il report fisso 'Report_specialita_medici.xlsx' per la consegna aziendale.

Legge i risultati della modalita' massiva (output/risultato_massivo_recuperato.xlsx e
il relativo .summary.json) e produce un workbook con tre fogli, nello stesso formato
del report di riferimento fornito dall'utente (generato in precedenza con
build_report_medici.py/mjs in un altro ambiente, non disponibili in questo repository):

- Riepilogo: numeri chiave e criteri di lettura (Alta / Media / Disciplina dichiarata).
- Specialita trovate: un record per persona con specialita' o disciplina utilizzabile
  (le tre categorie SPECIALITA_DOCUMENTATA, SPECIALITA_CON_IDENTITA_NOMINALE,
  DISCIPLINA_DICHIARATA_NEL_PROFILO: il "totale utile" del progetto).
- CV utili: lo stesso sottoinsieme, filtrato a chi ha anche un CV acquisito.

Uso:
    python3 build_report_medici.py \
        --massivo-xlsx output/risultato_massivo_recuperato.xlsx \
        --summary-json output/risultato_massivo_recuperato.summary.json \
        --output output/Report_specialita_medici.xlsx
"""
from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill('solid', fgColor='17365D')
HEADER_FONT = Font(name='Arial', size=11, bold=True, color='FFFFFF')
TITLE_FONT = Font(name='Arial', size=16, bold=True, color='17365D')
SUBTITLE_FONT = Font(name='Arial', size=10, color='595959')
LABEL_FONT = Font(name='Arial', size=11, bold=True)
BODY_FONT = Font(name='Arial', size=10)

# Le tre categorie che compongono il "totale utile" del progetto (vedi
# RELAZIONE_PROGETTO_MEDICI.md): non sommare le altre categorie qui.
USEFUL_STATES = {
    'SPECIALITA_DOCUMENTATA': ('Specialita documentata', 'Alta'),
    'SPECIALITA_CON_IDENTITA_NOMINALE': ('Specialita (identita nominale)', 'Media'),
    'DISCIPLINA_DICHIARATA_NEL_PROFILO': ('Disciplina dichiarata', 'Disciplina dichiarata'),
}

COLUMNS = ['Pers_Id', 'Medico_Id', 'Cognome', 'Nome', 'Città', 'Data nascita',
           'Codice fiscale', 'Specialità / disciplina', 'Tipo di dato', 'Affidabilità',
           'Verifica identità', 'Evidenza', 'Tipo fonte', 'Fonti', 'Stato CV',
           'File CV', 'Note']

COL_WIDTHS = [13, 12, 18, 18, 18, 13, 18, 30, 28, 13, 24, 42, 18, 42, 30, 42, 42]

MESI_IT = ['gennaio', 'febbraio', 'marzo', 'aprile', 'maggio', 'giugno', 'luglio',
           'agosto', 'settembre', 'ottobre', 'novembre', 'dicembre']


def fmt_data_ora_italiana(dt_iso, tz_offset_ore=2):
    dt = datetime.fromisoformat(str(dt_iso).replace('Z', '+00:00'))
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    dt = dt.astimezone(timezone(timedelta(hours=tz_offset_ore)))
    return f'{dt.day} {MESI_IT[dt.month - 1]} {dt.year}, ore {dt:%H:%M}'


def join_nonempty(*parts, sep='\n'):
    seen = []
    for p in parts:
        if p is None:
            continue
        p = str(p).strip()
        if p and p not in seen:
            seen.append(p)
    return sep.join(seen)


REQUIRED_COLUMNS = [
    'Pers_Id', 'Medico_Id', 'Pers_Cognome', 'Pers_Nome', 'Indirizzi_Citta',
    'Pers_DataNascita', 'Pers_CodFis', 'Massivo_Stato', 'Identita_Verifica',
    'Specialita_Documentata', 'Specialita_Proposta', 'Disciplina_Dichiarata_Profilo',
    'Specialita_Evidenza_Massivo', 'Disciplina_Evidenza', 'Tipo_Fonte',
    'CV_Stato_Massivo', 'CV_File_Massivo', 'Fonti_Massivo', 'Motivo_Massivo',
]


def load_rows(massivo_path):
    wb = load_workbook(massivo_path, read_only=True, data_only=True)
    ws = wb['Foglio1']
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    idx = {name: i for i, name in enumerate(header)}
    missing = [c for c in REQUIRED_COLUMNS if c not in idx]
    if missing:
        raise SystemExit(f'Colonne mancanti in Foglio1 di {massivo_path}: {missing}')

    specialita_rows, cv_rows = [], []
    for row in rows:
        stato = row[idx['Massivo_Stato']]
        info = USEFUL_STATES.get(stato)
        if not info:
            continue
        tipo_dato, affidabilita = info
        specialita = join_nonempty(
            row[idx['Specialita_Documentata']],
            row[idx['Specialita_Proposta']],
            row[idx['Disciplina_Dichiarata_Profilo']],
            sep='; ',
        )
        evidenza = join_nonempty(row[idx['Specialita_Evidenza_Massivo']], row[idx['Disciplina_Evidenza']])
        record = [
            row[idx['Pers_Id']], row[idx['Medico_Id']], row[idx['Pers_Cognome']],
            row[idx['Pers_Nome']], row[idx['Indirizzi_Citta']], row[idx['Pers_DataNascita']],
            row[idx['Pers_CodFis']], specialita, tipo_dato, affidabilita,
            row[idx['Identita_Verifica']], evidenza, row[idx['Tipo_Fonte']],
            row[idx['Fonti_Massivo']], row[idx['CV_Stato_Massivo']],
            row[idx['CV_File_Massivo']], row[idx['Motivo_Massivo']],
        ]
        specialita_rows.append(record)
        cv_stato = row[idx['CV_Stato_Massivo']]
        if cv_stato and cv_stato != 'CV non acquisito':
            cv_rows.append(record)
    return specialita_rows, cv_rows


def write_detail_sheet(wb, name, rows):
    ws = wb.create_sheet(name)
    ws.append(COLUMNS)
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    for i, width in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = 'E2'
    # Il font di corpo e' impostato una sola volta a livello di workbook
    # (stile 'Normal', vedi build()): applicarlo cella per cella su decine
    # di migliaia di righe x 17 colonne renderebbe la generazione troppo
    # lenta senza benefici visibili.
    for record in rows:
        ws.append(record)
    return ws


def build(massivo_path, summary_path, output_path):
    summary = json.loads(Path(summary_path).read_text(encoding='utf-8'))
    states = summary['states']
    totale_archivio = summary['total']
    doc = states.get('SPECIALITA_DOCUMENTATA', 0)
    nominale = states.get('SPECIALITA_CON_IDENTITA_NOMINALE', 0)
    disciplina = states.get('DISCIPLINA_DICHIARATA_NEL_PROFILO', 0)

    import time
    t0 = time.monotonic()
    specialita_rows, cv_rows = load_rows(massivo_path)
    print(f'[build_report] lettura completata in {time.monotonic() - t0:.1f}s: '
          f'{len(specialita_rows)} righe utili, {len(cv_rows)} con CV', flush=True)

    wb = Workbook()
    wb.remove(wb.active)
    for style in wb._named_styles:
        if style.name == 'Normal':
            style.font = BODY_FONT

    ws = wb.create_sheet('Riepilogo')
    ws.column_dimensions['A'].width = 44
    for col in 'BCDEF':
        ws.column_dimensions[col].width = 18
    ora_generazione = fmt_data_ora_italiana(datetime.now(timezone.utc).isoformat())
    ora_acquisizione = fmt_data_ora_italiana(summary['updated'])
    ws['A2'] = 'Report specialità medici'
    ws['A2'].font = TITLE_FONT
    ws['A3'] = (f'Report aggiornato: {ora_generazione}. Ultima acquisizione: '
                f'{ora_acquisizione}. Archivio di {totale_archivio} medici.')
    ws['A3'].font = SUBTITLE_FONT
    ws['A5'], ws['B5'] = 'Indicatore', 'Medici'
    for coord in ('A5', 'B5'):
        ws[coord].font = HEADER_FONT
        ws[coord].fill = HEADER_FILL
    ws['A6'], ws['B6'] = 'Specialità documentata con identità forte', doc
    ws['A7'], ws['B7'] = 'Specialità da profilo nominale', nominale
    ws['A8'], ws['B8'] = 'Disciplina dichiarata nel profilo', disciplina
    ws['A9'], ws['B9'] = 'Totale con specialità o disciplina', '=SUM(B6:B8)'
    ws['A10'], ws['B10'] = 'Quota sul totale archivio', f'=B9/{totale_archivio}'
    ws['B10'].number_format = '0.0%'
    ws['A12'], ws['B12'] = 'CV acquisiti con specialità o disciplina', len(cv_rows)
    ws['A14'] = 'Criteri di lettura'
    ws['A14'].font = LABEL_FONT
    ws['A15'], ws['B15'] = 'Livello', 'Significato'
    for coord in ('A15', 'B15'):
        ws[coord].font = HEADER_FONT
        ws[coord].fill = HEADER_FILL
    ws['A16'], ws['B16'] = 'Alta', 'Identità confermata tramite dato anagrafico e specialità esplicita nella fonte.'
    ws['A17'], ws['B17'] = 'Media', 'Nome e cognome coincidono con il profilo pubblico; manca un discriminante anagrafico.'
    ws['A18'], ws['B18'] = 'Disciplina dichiarata', 'Area professionale indicata dal profilo; può non equivalere a un diploma di specializzazione.'
    ws['A20'] = ('Fonte dati: ' + Path(massivo_path).name + '. Il report conserva separati i livelli di '
                 'prova e non trasforma una disciplina dichiarata in specializzazione certificata.')
    ws['A20'].font = SUBTITLE_FONT

    t1 = time.monotonic()
    write_detail_sheet(wb, 'Specialita trovate', specialita_rows)
    write_detail_sheet(wb, 'CV utili', cv_rows)
    print(f'[build_report] fogli scritti in {time.monotonic() - t1:.1f}s', flush=True)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    t2 = time.monotonic()
    wb.save(output_path)
    print(f'[build_report] salvataggio completato in {time.monotonic() - t2:.1f}s', flush=True)
    return {
        'totale_utile': len(specialita_rows),
        'cv_utili': len(cv_rows),
        'output': str(Path(output_path).resolve()),
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--massivo-xlsx', default='output/risultato_massivo_recuperato.xlsx')
    parser.add_argument('--summary-json', default='output/risultato_massivo_recuperato.summary.json')
    parser.add_argument('--output', default='output/Report_specialita_medici.xlsx')
    args = parser.parse_args()
    result = build(args.massivo_xlsx, args.summary_json, args.output)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()

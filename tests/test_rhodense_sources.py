import csv
import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import aggiorna_fonti_rhodense as sources


class SourceIndexTests(unittest.TestCase):
    def test_medical_rows_only_and_public_pdf_links(self):
        html = b'''<table><tr><th>Cognome</th><th>Nome</th><th>Profilo</th><th>Mail</th><th>CV</th></tr>
        <tr><td>ROSSI</td><td>ANNA</td><td>MEDICI</td><td></td><td><a href="CVDirigenti/1.pdf">CV</a></td></tr>
        <tr><td>BIANCHI</td><td>MARIO</td><td>PSICOLOGI</td><td></td><td><a href="CVDirigenti/2.pdf">CV</a></td></tr>
        <tr><td>VERDI</td><td>LUCA</td><td>MEDICI</td><td></td><td><a href="https://other.invalid/a.pdf">CV</a></td></tr></table>'''
        result = sources.parse_index(html)
        self.assertEqual(list(result), [('rossi', 'anna')])
        self.assertTrue(result[('rossi', 'anna')][0].endswith('/CVDirigenti/1.pdf'))

    def test_unknown_layout_fails(self):
        with self.assertRaises(ValueError):
            sources.parse_index(b'<html>Temporarily unavailable</html>')

    def test_merge_preserves_existing_and_adds_next_batch(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory)/'sources.csv'
            path.write_text('Pers_Id,URL,Nota,Custom\n1,https://example.org/1.pdf,originale,keep\n')
            candidates = [{'Pers_Id': str(i), 'URL':f'https://example.org/{i}.pdf', 'Nota':'candidate'} for i in range(1,4)]
            self.assertEqual(sources.merge_manifest(path, candidates, 1), 1)
            self.assertEqual(sources.merge_manifest(path, candidates, 1), 1)
            self.assertEqual(sources.merge_manifest(path, candidates, 1), 0)
            with path.open() as stream:
                rows = list(csv.DictReader(stream))
            self.assertEqual(len(rows), 3)
            self.assertEqual(rows[0]['Custom'], 'keep')
            self.assertEqual(rows[0]['Nota'], 'originale')

    def test_matching_uses_both_names_and_keeps_homonyms_as_candidates(self):
        from openpyxl import Workbook
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory)/'input.xlsx'
            book = Workbook()
            book.active.title = 'Foglio1'
            book.active.append(['Pers_Id','Pers_Cognome','Pers_Nome'])
            for row in [('1','Rossi','Anna'),('2','Rossi','Anna Maria'),('3','Rossi','Anna')]:
                book.active.append(row)
            book.save(path)
            index = {('rossi','anna'): ['https://example.org/1.pdf']}
            result = sources.match_people(path, 'Foglio1', index)
            self.assertEqual([r['Pers_Id'] for r in result], ['1','3'])


if __name__ == '__main__':
    unittest.main()

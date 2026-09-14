"""Diagnostica una tantum: mostra come pypdf estrae il testo di un PDF esterno
e se le righe vengono riconosciute come intestazione di specialita' dall'algoritmo
pdf_activity_roster (stessa logica di discover_one in medici_massivo.py).

Uso:  python debug_pdf_extract.py <url_pdf> [n_righe]

Non tocca stato_massivo, non scrive nulla: solo stampa a schermo. Cancellare
questo file dopo l'uso (non serve nel repository a lungo termine).
"""
import io
import sys

import requests

import medici_massivo as m

url = sys.argv[1] if len(sys.argv) > 1 else \
    "https://www.gaslini.org/wp-content/uploads/2026/04/SITO-Elenco-Medici-Esterno.pdf"
n = int(sys.argv[2]) if len(sys.argv) > 2 else 60

print(f"Scarico: {url}")
raw = requests.get(url, timeout=30).content
print(f"Bytes scaricati: {len(raw)}")

reader = m.PdfReader(io.BytesIO(raw))
aliases = {m.key(label): value for label, value in m.core.SPECIALTY_ALIASES.items()}
aliases.update({m.key(value): value for value in m.core.SPECIALTY_ALIASES.values()})

shown = 0
for page_num, page in enumerate(reader.pages):
    text = page.extract_text() or ''
    lines = text.splitlines()
    print(f"\n--- Pagina {page_num + 1}: {len(lines)} righe estratte ---")
    for line in lines:
        line_key = m.key(line)
        tag = "  <== MATCH SPECIALITA'" if line_key in aliases else ""
        print(f"{shown:3d} | {line!r}{tag}")
        shown += 1
        if shown >= n:
            break
    if shown >= n:
        break

print(f"\nTotale righe mostrate: {shown}")
print(f"Specialita' note nel catalogo (per confronto): {len(m.core.SPECIALTY_ALIASES)}")

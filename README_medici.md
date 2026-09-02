# Ricerca Medici - Specialità + CV

Nuovo progetto ripartito da zero.

## Obiettivo

Per ogni medico presente nell'Excel:

1. cercare un CV attendibile;
2. verificare che il CV appartenga davvero al medico;
3. scaricare il CV come:
   `<Medico_Id>-<Cognome>-<Nome>.pdf`
   oppure, se `Medico_Id` è `0`/vuoto:
   `<Pers_Id>-<Cognome>-<Nome>.pdf`;
4. ricavare la specialità medica dal CV quando possibile;
5. se il CV non basta, cercare la specialità su fonti web affidabili;
6. usare Gemini solo come supporto sulle evidenze già raccolte;
7. creare un nuovo Excel con le colonne aggiunte.

## Colonne dell'input

Lo script è stato progettato sul file `scriptMedici.xlsx`.
Le colonne minime obbligatorie sono:

- `Pers_Id`
- `Pers_Cognome`
- `Pers_Nome`

Utilizza inoltre, quando presenti:

- `Medico_Id`
- `Pers_DataNascita`
- `Pers_CodFis`
- `Indirizzi_Citta`
- `Email` / `emailPredefinita`

## Installazione

Consigliato Python 3.12.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements_medici.txt
```

## .env

Crea un file `.env` nella root.

Configurazione minima gratuita:

```env
GEMINI_API_KEY=la_tua_chiave
GEMINI_MODEL=gemini-3.7-flash
```

Lo script funziona anche senza Gemini, usando `--no-ai`.

### Search API opzionali

Per aumentare stabilità/qualità delle ricerche puoi configurare UNO di questi:

```env
SERPER_API_KEY=...
```

oppure:

```env
BRAVE_SEARCH_API_KEY=...
```

Priorità automatica:

1. Serper/Google, se configurato
2. Brave Search API, se configurato
3. DDGS/DuckDuckGo come fallback gratuito

## Primo test consigliato

```powershell
python ricerca_medici_specialita_cv.py "scriptMedici.xlsx" --output "risultati_medici_test.xlsx" --limit 20 --workers 6 --search-concurrency 3
```

Controlla manualmente i primi 20 record.

Poi test da 100:

```powershell
python ricerca_medici_specialita_cv.py "scriptMedici.xlsx" --output "risultati_medici.xlsx" --limit 100 --workers 8 --search-concurrency 4
```

E infine tutto il file:

```powershell
python ricerca_medici_specialita_cv.py "scriptMedici.xlsx" --output "risultati_medici.xlsx" --workers 8 --search-concurrency 4
```

## Output Excel

Vengono aggiunte:

- `Ricerca_Stato`
- `Specialita`
- `Specialita_Confidenza`
- `CV_Salvato`
- `CV_URL`
- `Fonti_Ricerca`
- `Ricerca_Note`
- `Ultimo_Aggiornamento_UTC`
- `Ricerca_Errore`

## Stati

- `COMPLETATO`: specialità + CV
- `SPECIALITA_TROVATA`: specialità trovata, CV non trovato
- `CV_TROVATO_SPECIALITA_DA_VERIFICARE`: CV verificato, ma specialità non estratta con sufficiente certezza
- `DA_VERIFICARE`: trovate evidenze, ma non sufficienti
- `NESSUN_RISULTATO`
- `ERRORE`

## CV

Cartella predefinita:

```text
cv_medici/
```

Formato:

```text
CODICE-COGNOME-NOME.pdf
```

Lo script non salva automaticamente graduatorie, bandi o PDF generici come CV.
Verifica nome/cognome, eventuale codice fiscale/data di nascita, struttura da curriculum e contesto medico.

## Cache / Resume

Ogni medico elaborato viene memorizzato in:

```text
cache_medici/
```

Questo evita di rifare gratuitamente le stesse ricerche dopo un'interruzione.

L'Excel viene salvato a checkpoint ogni 20 risultati per default.

## Parametri utili

```text
--limit 20
--workers 8
--search-concurrency 4
--search-results 6
--save-every 20
--no-ai
--retry-completed
--retry-errors
```

Per evitare blocchi dei motori di ricerca non conviene impostare `--search-concurrency`
troppo alto. È possibile avere 8-12 worker, mantenendo 3-5 ricerche web contemporanee.

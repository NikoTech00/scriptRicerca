# Ricerca medici e recupero CV

`scriptMedici.py` ricerca specialità e CV a partire da un Excel. La modalità locale permette di lavorare senza chiavi API e senza query web sui documenti già scaricati. Le modifiche operative sono compatibili con la pipeline V6.0.

## Installazione

Usare Python 3.10 o successivo nell'ambiente virtuale del progetto:

```powershell
python -m pip install -r requirements.txt
```

Su macOS usare `python3` se `python` non è disponibile. Eseguire i comandi dalla cartella che contiene lo script. Chiudere l'Excel di output prima del salvataggio.

## Recupero locale senza API

```powershell
python scriptMedici.py "input/scriptMedici.xlsx" --offline --output "output/recupero_locale.xlsx"
```

L'output deve essere un file nuovo. Lo script legge `cv_medici` e `cv_medici_da_verificare`, seleziona i candidati tramite codice e nome del file, poi ne verifica il contenuto. Il report contiene solo i medici per cui sono disponibili documenti locali. Gli altri nominativi non sono classificati come risultati negativi. L'input, i CV e i risultati precedenti sono conservati.

Questa modalità non crea un client API, non usa la cache delle ricerche e non effettua richieste HTTP. Per i DOC legacy utilizza i convertitori locali disponibili (antiword, catdoc, LibreOffice, Word su Windows). Se rimane disponibile solo l'estrazione euristica, il documento resta da verificare. Un PDF senza testo estratto richiede verifica o OCR separato.

La specialità viene proposta solo da diciture esplicite nel CV verificato; riferimenti a reparti o pubblicazioni non bastano. Le specialità in conflitto restano da verificare. La verifica automatica non sostituisce il controllo umano dei risultati aziendali, soprattutto per gli omonimi.

## Recupero da fonti pubbliche dirette, senza Search API

Il file `fonti_campione.csv` contiene URL candidati con le colonne `Pers_Id,URL,Nota`. Le note descrivono la provenienza; non autorizzano il programma a considerare verificata una persona. Non occorrono account o carte.

```powershell
python scriptMedici.py "input/scriptMedici.xlsx" --sources-file "fonti_campione.csv" --max-direct-downloads 4 --output "output/test_fonti_dirette.xlsx"
```

Il report include i medici con CV locali e quelli presenti nel CSV. Viene letto l'input originale e il report deve essere un file nuovo. Sono ammessi solo URL HTTP(S) pubblici; anche i redirect sono controllati. Non vengono interrogati motori di ricerca né seguiti automaticamente collegamenti nelle pagine. Un URL candidato va quindi aggiunto al CSV, oppure fornito da un indice pubblico analizzato separatamente.

Le copie scaricate rimangono in `fonti_dirette` insieme a un file `.source.json` contenente URL iniziale/finale, data di acquisizione e impronta SHA-256. Le copie integre sono riutilizzate senza un nuovo download. Il limite riguarda gli URL tentati, esclusi quelli già in cache; ciascun URL può richiedere fino a cinque redirect. Un errore di download è registrato nelle note, non diventa `NESSUN_RISULTATO`. Copiare anche `fonti_dirette` sull'altro PC evita nuovi download delle stesse fonti.

Questa modalità verifica i CV; non interpreta automaticamente graduatorie, pubblicazioni o elenchi come specializzazioni. Il CSV del campione contiene anche riferimenti utili alla verifica manuale: non tutti produrranno un CV o una specialità. Per aggiornare una fonte salvata, conservarne la vecchia copia altrove e rimuovere la relativa coppia documento/metadata dalla cache.

## Verifiche corrette il 4 settembre 2026

- Date di nascita numeriche con o senza zeri, ISO e mesi italiani in lettere.
- Il CV DOC di Ornella Abate presente fra quelli precedentemente verificati appartiene a un'omonima: il contenuto ora viene scartato per nascita incompatibile. Il documento originale è conservato.
- Il CV di Marco Adorni descrive una specializzazione in corso: quel documento non prova il titolo conseguito.
- Scuole di specializzazione e insegnamenti non valgono come diplomi del titolare del CV.
- Più diplomi espliciti nello stesso CV verificato possono essere riportati insieme.
- Un generico profilo biografico HTML non viene chiamato CV sulla base delle sole voci di menu.

Usare `output/report_medici_verificato_20260904.xlsx` per l'ultima verifica del campione. I precedenti `recupero_locale_20260904.xlsx` e `recupero_fonti_dirette_20260904.xlsx` sono conservati come risultati intermedi e sono superati.

## Eventuale uso futuro delle Search API

Serper rimane disponibile quando il vostro account dispone di crediti. Brave richiede una verifica ulteriore: la documentazione ufficiale richiede diritti espliciti di conservazione per salvare i risultati API. Non è stata verificata la disponibilità di tali diritti con i crediti gratuiti. Non usare Brave in questo flusso persistente finché il piano aziendale non consente cache e conservazione dei risultati.

Riferimenti ufficiali: https://brave.com/search/api/ e https://api-dashboard.search.brave.com/documentation/resources/help-feedback

Il limite predefinito è 100 chiamate API per esecuzione, compresi errori e fallback. `--max-api-requests` cambia questo limite; non controlla il saldo o la fatturazione dell'account. `--limit 40` limita il numero di medici della prova. Non avviare più esecuzioni contemporanee sullo stesso output.

## Cache e ripresa

- `cache_ricerche`: conserva le risposte API riuscite, anche quelle realmente vuote. La stessa query, provider e numero di risultati viene riutilizzata senza nuove chiamate. Gli errori non sono salvati come risposte vuote.
- `cache_medici`: conserva i risultati finali per persona. Le vecchie cache prive del nuovo schema non vengono accettate automaticamente, perché potrebbero contenere falsi negativi prodotti dagli errori delle versioni precedenti. Conservarle per un'eventuale analisi separata.
- `--ignore-cache`: ricalcola il risultato per persona; continua a riutilizzare le risposte in `cache_ricerche`.
- `--refresh-search-cache`: ripete anche le query salvate, consumando nuove richieste. Usarlo quando occorre aggiornare le fonti; la cache non ha scadenza automatica.
- `--retry-all`: rielabora anche le righe già completate. Per le righe nello stato `ERRORE` serve anche `--retry-errors`.

Per ripetere o riprendere le prove, conservare entrambe le cartelle di cache. I percorsi dei CV salvati su un altro computer potrebbero non essere validi; la modalità locale ricostruisce i riferimenti ai documenti effettivamente presenti.

## Stati e interruzioni

- `COMPLETATO`: CV verificato automaticamente e specialità proposta con evidenza.
- `SPECIALITA_TROVATA`: specialità proposta, CV non trovato.
- `CV_TROVATO_SPECIALITA_DA_VERIFICARE`: CV verificato, specialità assente o in conflitto.
- `DA_VERIFICARE`: evidenze insufficienti o documenti non verificati.
- `NESSUN_RISULTATO`: nessuna evidenza dopo una ricerca API riuscita; non viene usato nel recupero locale.
- `BLOCCATO_QUOTA_RICERCA`: quota o crediti del provider esauriti, incluso il messaggio Serper HTTP 400 `Not enough credits`.
- `BLOCCATO_LIMITE_RICHIESTE`: raggiunto il tetto locale di chiamate.
- `BLOCCATO_AUTENTICAZIONE_API` / `BLOCCATO_ERRORE_API`: errore del servizio, esecuzione interrotta e record da riprovare.

Le interruzioni API terminano con codice di uscita 1. Se la sostituzione dell'Excel fallisce, il file completo di recupero viene conservato e il suo percorso compare nell'errore. Non cancellarlo prima di aver recuperato i risultati.

## Test locali

```powershell
python -B tests/test_failure_handling.py
```

I test simulano le API e i blocchi di salvataggio senza consumare crediti. Coprono cache, budget, errori quota, selezione dei documenti, identità incompatibili e report offline.

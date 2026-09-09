# Ricerca sull'intero archivio

Eseguire dalla cartella del progetto con Python 3.10 o successivo:

```powershell
python -m pip install -r requirements.txt
python scriptMedici.py "input/scriptMedici.xlsx" --massivo --output "output/risultati_massivi.xlsx"
```

Questa modalità importa tutte le righe dell'Excel originale e produce un report completo. Non occorrono Serper, Brave, chiavi o servizi a pagamento. I dati anagrafici rimangono locali: vengono richiesti gli indici pubblici degli enti e i profili candidati, senza inviare il file Excel o i codici fiscali a servizi esterni.

Il programma incrocia gli indici con i nomi dell'archivio, acquisisce i profili corrispondenti e segue i collegamenti ai CV. Le copie e la coda rimangono in `stato_massivo`. Ogni URL è scaricato una volta e verificato tramite SHA-256 prima del riutilizzo. I CV già presenti nelle cartelle del progetto sono inclusi. Fonti, controlli d'identità ed evidenze rimangono tracciabili.

## Continuare il lavoro

Ripetere **lo stesso comando**: riparte dalla coda salvata e aggiorna lo stesso report. Non serve inventare un nome nuovo a ogni esecuzione. Un file esistente che non appartiene a questa modalità non viene sovrascritto. Chiudere il report in Excel prima di aggiornarlo.

I limiti predefiniti per esecuzione sono 5.000 documenti e 6.000 richieste HTTP, con sei worker. Le richieste a un singolo sito sono distanziate. I limiti comprendono i redirect e robots.txt, ma non le letture dalla cache. Non sono crediti API né garanzie di completamento entro un certo tempo.

Per una sessione più lunga:

```powershell
python scriptMedici.py "input/scriptMedici.xlsx" --massivo --max-documents 30000 --max-http-requests 40000 --output "output/risultati_massivi.xlsx"
```

Durante una run lunga il report viene aggiornato automaticamente ogni 1.000 documenti.
Il terminale mostra `Checkpoint salvato` quando il file Excel intermedio è pronto.
La frequenza si può cambiare, per esempio con `--mass-checkpoint-every 500`.

Il nome indicato con `--output` viene controllato prima dei download. Se esiste già
un file che non è stato creato dalla modalità massiva, il programma si ferma subito
e chiede un nome nuovo, senza lavorare inutilmente per ore.

## Seconda fase: Top Doctors e iDoctors

Il catalogo include anche Top Doctors e iDoctors. Top Doctors permette l'incrocio
diretto del nome dalla sitemap. Su iDoctors gli URL sono numerici: la pagina viene
prima scaricata e poi collegata soltanto quando l'intestazione contiene esattamente
un nome dell'archivio. Le schede non corrispondenti non generano risultati.

Dopo aver aggiornato il codice, il primo ciclo della seconda fase va eseguito senza
`--skip-discovery`:

```powershell
python .\scriptMedici.py "input\scriptMedici.xlsx" --massivo --max-documents 2500 --max-http-requests 3000 --mass-checkpoint-every 500 --output "output\risultato_massivo_recuperato.xlsx"
```

Nei cicli successivi si può aggiungere `--skip-discovery`. La coda SQLite consente
di distribuire il lavoro su più sessioni senza ripetere le schede completate.

Si può interrompere con Ctrl+C. I documenti completati sono già registrati in SQLite; l'interruzione ordinaria esporta il report dopo la chiusura delle richieste in corso. Dopo una chiusura forzata del PC, rilanciare il comando. Una seconda esecuzione contemporanea sullo stesso stato viene bloccata.

Per esportare lo stato raggiunto senza richieste di rete:

```powershell
python scriptMedici.py "input/scriptMedici.xlsx" --massivo --export-only --output "output/risultati_massivi.xlsx"
```

Per riprovare fonti non accessibili aggiungere `--retry-errors`. Non vengono aggirati login, CAPTCHA o blocchi HTTP. Gli errori del sito non sono classificati come assenza del medico.

`--refresh-sources` rilegge gli indici online e aggiunge le nuove associazioni. `--skip-discovery` lavora sulla coda esistente. Per un Excel diverso usare un'altra cartella, ad esempio `--state-dir "stato_nuovo_archivio"`, e un altro report.

## Leggere i risultati

Il foglio originale contiene tutte le righe di input e nuove colonne `Massivo_*`, oltre a quelle relative a identità, specialità, disciplina e CV. Le colonne originarie non vengono usate per nascondere gli esiti nuovi. Il foglio `Copertura` calcola i conteggi sull'intero archivio; `Fonti` mostra l'esito dell'acquisizione di ciascun indice.

- `SPECIALITA_DOCUMENTATA`: titolo esplicito e identità sostenuta da data di nascita o codice fiscale concordante.
- `SPECIALITA_CON_IDENTITA_NOMINALE`: titolo esplicito nel profilo/CV, ma il collegamento con la persona dell'Excel è basato sul nome completo. La proposta resta separata dalla specialità confermata.
- `DISCIPLINA_DICHIARATA_NEL_PROFILO`: attività indicata nel campo professionale del profilo. Non prova un diploma di specializzazione.
- `IDENTITA_CONFERMATA_SENZA_TITOLO`: identità concordante, titolo non ricavato.
- `PROFILO_CON_IDENTITA_NOMINALE`: nome concordante, informazioni sul titolo insufficienti.
- `EVIDENZA_NON_CONFERMATA`: omonimia, dati incompatibili, contenuto insufficiente o documento da convertire/OCR.
- `IN_CODA`: fonte candidata acquisita nell'indice, contenuto ancora da elaborare.
- `FONTE_NON_ACCESSIBILE`: tentativo non riuscito o sito sospeso per errore/accesso.
- `NON_COPERTO_DALLE_FONTI`: nessun candidato negli indici finora acquisiti. Non significa che il medico o il CV non esistano online.

La presenza del CV e la verifica della specialità sono separate. L'identità nominale non diventa anagrafica solo perché il nome compare su più siti. Gli omonimi sono verificati rispetto a tutto l'input, non solo al lotto in elaborazione. Reparti, pubblicazioni, iscrizioni a scuole e occupazioni desiderate non valgono come titoli conseguiti.

## Ampliare la copertura

`fonti_massive.json` è il catalogo modificabile: supporta sitemap XML, indici di sitemap e tabelle/pagine con collegamenti ai profili o ai CV. Ogni fonte dichiara domini ammessi e regole per selezionare gli URL. Non richiede modifiche al motore per aggiungere un ente con uno di questi formati.

Il catalogo non equivale a un'anagrafe nazionale esaustiva: non tutti i medici hanno un profilo pubblico e non tutti gli enti espongono indici scaricabili. Gli indici con URL esclusivamente numerici richiedono un adattatore con un elenco nominativo; non si inventano nomi dal codice URL. Le sitemap dei portali commerciali sono fonti di profili pubblicati, non registri ufficiali dei titoli.

Il limite residuo si misura dai record non coperti e dalle identità non disambiguate. Nessun cambio di linguaggio o incremento dei worker può creare evidenze mancanti. Per arrivare a una copertura molto più elevata servono altri elenchi nominativi utilizzabili o una fonte nazionale accessibile: il programma può essere esteso per importarli senza rifare la pipeline.

## Aggiornamento tramite GitHub

Caricare le modifiche a `scriptMedici.py`, `medici_massivo.py`, `fonti_massive.json`, `.gitignore`, questa guida e i test. Il computer Windows esegue il comando nella propria copia del repository. La cartella `stato_massivo` e le credenziali sono escluse da Git; la nuova macchina costruirà la propria cache e la propria coda.

Test locali, senza richieste reali a provider:

```powershell
python -B -m unittest discover -s tests -q
```

## Vocabolario delle discipline

Il vocabolario include denominazioni del [catalogo ufficiale delle scuole di specializzazione dell’Università degli Studi di Milano](https://www.unimi.it/it/corsi/corsi-post-laurea-e-formazione-continua/scuole-di-specializzazione-catalogo-corsi/catalogo-scuole-di-specializzazione-di-area-medica-e-sanitaria), consultato il 4 settembre 2026, oltre alle varianti già gestite dal progetto. Le etichette servono al raggruppamento dei risultati; non certificano equipollenze legali fra diplomi. Il testo originale del titolo è conservato nell’evidenza. I titoli fuori dal vocabolario restano nelle note per la normalizzazione, senza essere promossi automaticamente a specializzazioni mediche.

## Verifica eseguita il 4 settembre 2026

Il confronto integrale ha confermato la conservazione di tutte le 81.331 righe e di tutti i valori delle colonne originarie; l’impronta dell’input è invariata. Sono passati 55 test locali. La versione finale ha rianalizzato 1.540 documenti e profili già acquisiti con zero richieste HTTP, verificando la cache. Sono state provate anche l’interruzione ordinaria, la ripresa e la sostituzione del report sullo stesso percorso.

Il report corrente contiene 17.495 record con almeno una fonte candidata (21,5% dell’input). La verifica dei documenti è ancora parziale: 15.539 record sono in coda, 63.836 non hanno candidati nelle fonti integrate. Sono presenti 2 specialità con anagrafica concordante, 368 proposte di specialità con identità nominale e 284 record con disciplina dichiarata nel profilo. Queste tre categorie non sono intercambiabili.

L’arricchimento dell’intero archivio resta incompleto. I conteggi delle fonti candidate non devono essere presentati come medici verificati. La prova sul computer Windows rimane da eseguire nella sua copia aggiornata del progetto.

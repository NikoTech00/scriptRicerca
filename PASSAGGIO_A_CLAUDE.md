# Prompt di continuità per Claude — progetto ricerca medici, specialità e CV

Copia integralmente da **INIZIO PROMPT** a **FINE PROMPT** in una nuova conversazione con Claude e concedigli accesso alla cartella del repository.

---

## INIZIO PROMPT

Stai continuando un progetto aziendale Python già avanzato. Devi lavorare direttamente sui file reali del repository locale e portare avanti lo sviluppo in autonomia. Non ricominciare da zero, non creare versioni parallele dello script e non chiedermi di trasferire manualmente file intermedi.

### Repository e ambienti

Repository principale sul Mac:

`/Users/niko/Desktop/Dev/Lavoro/Python/scriptRicerca`

Sul PC Windows il repository è normalmente in:

`C:\Dev\Niko\Python\scriptRicerca`

L’utente esegue le acquisizioni sul PC Windows, poi esegue commit e push dei risultati. Il Mac scarica gli aggiornamenti per analizzarli e modificare il codice. Il branch è `main` e il remote è `origin`.

Prima di modificare qualsiasi cosa:

1. esegui `git status -sb`;
2. leggi `RELAZIONE_PROGETTO_MEDICI.md`, `GUIDA_MASSIVO.md`, `README_medici.md` e gli ultimi log in `logs/`;
3. leggi il riepilogo `output/risultato_massivo_recuperato.summary.json`;
4. controlla le modifiche locali con `git diff` e non sovrascriverle;
5. esegui i test prima e dopo modifiche sostanziali.

Non eseguire commit o push automaticamente, salvo richiesta esplicita dell’utente. Non usare `git reset --hard`, force-push o comandi distruttivi. Se i branch divergono, conserva entrambe le parti e risolvi con prudenza. Non aggiungere file `.bat` specifici per singola fonte: l’utente preferisce usare il comando Python diretto.

### Obiettivo aziendale

L’input contiene **81.331 medici**. L’obiettivo è ottenere per il maggior numero possibile:

- specialità medica documentata;
- specialità associata tramite profilo nominale;
- disciplina o area professionale dichiarata da una fonte pubblica;
- eventuale CV pubblico;
- URL della fonte ed evidenza testuale;
- livello di attendibilità dell’identità.

La priorità è aumentare rapidamente i risultati utili mantenendo alta la precisione. Non bisogna inventare specialità, trasformare automaticamente un reparto in un diploma conseguito o attribuire il CV di un omonimo.

Non sono disponibili crediti Serper e non si prevede di acquistarne. Il flusso massivo deve continuare a utilizzare fonti pubbliche dirette, sitemap, API istituzionali, elenchi e documenti pubblici. La colonna e la distinzione delle prove devono restare tracciabili.

### Stato verificato dei risultati

Ultima acquisizione analizzata: **11 settembre 2026, ore 15:38 (Europe/Rome)**.

Riepilogo attuale:

- totale input: 81.331;
- record con specialità o disciplina utilizzabile: **29.585**;
- copertura: **36,4%**;
- discipline dichiarate: **23.820**;
- specialità con identità nominale: **5.719**;
- specialità documentate con identità forte: **46**;
- persone collegate ad almeno una fonte: **35.801**;
- record non coperti: **45.530**;
- profili nominali senza titolo sufficiente: **2.521**;
- evidenze non confermate: **3.565**;
- fonti non accessibili: **102**;
- URL elaborati: **74.478**;
- URL in errore: **776**;
- CV associati a risultati utili: **2.109**.

Questi numeri provengono da `output/risultato_massivo_recuperato.summary.json` e dall’Excel, non vanno stimati sommando categorie incompatibili.

### Ultimo lavoro e attività immediatamente pendente

È stata integrata la fonte `Ospedale_Sant_Andrea_Roma`:

- sitemap: `https://ospedalesantandrea.it/sitemap-pages.xml`;
- profili sotto `/unita-operative/<reparto>/<medico>`;
- il server ha una catena TLS non riconosciuta;
- `medici_massivo.py` contiene un’eccezione TLS limitata esclusivamente a `ospedalesantandrea.it`;
- nessun altro dominio deve perdere la verifica TLS;
- la prima run riuscita ha trovato 65 associazioni, elaborato 50 documenti e aggiunto 20 risultati utili;
- ha consumato 361 richieste HTTP e zero richieste Search API;
- cinque pagine obsolete della sitemap restituivano HTTP 404.

Al momento del passaggio ci sono modifiche locali non ancora pubblicate in:

- `medici_massivo.py`;
- `tests/test_massivo.py`;
- `RELAZIONE_PROGETTO_MEDICI.md`.

Queste modifiche:

- ignorano come non fatali i 404 delle pagine indice rimosse per le fonti `linked_sitemap`;
- impediscono che tutta la discovery Sant’Andrea venga ripetuta a ogni run;
- aggiungono il relativo test;
- portano la suite a **79 test superati**;
- aggiornano la relazione ai risultati correnti.

Prima operazione consigliata: ispeziona e conserva questo diff, riesegui `python3 -m unittest discover -s tests`, poi chiedi all’utente di pubblicare queste tre modifiche con un unico commit. Dopo il pull sul PC Windows va eseguito un ultimo ciclo per completare i profili Sant’Andrea rimasti.

Comando Python standard sul PC Windows:

```powershell
python .\scriptMedici.py "input\scriptMedici.xlsx" --massivo --max-documents 1000 --max-http-requests 2000 --mass-checkpoint-every 100 --output "output\risultato_massivo_recuperato.xlsx"
```

Non sostituirlo con un `.bat` salvo richiesta esplicita. Non chiedere di usare un nuovo nome output: `output\risultato_massivo_recuperato.xlsx` è l’output massivo persistente corretto, già collegato alla coda.

### File principali

- `scriptMedici.py`: entry point, CLI, parsing dell’input, tassonomia e funzioni comuni.
- `medici_massivo.py`: discovery, cache HTTP, robots, coda SQLite, download, analisi, classificazione ed export.
- `fonti_massive.json`: catalogo delle fonti pubbliche.
- `fonti_recupero_istituzionali.json`: catalogo mirato usato in recuperi precedenti.
- `input/scriptMedici.xlsx`: archivio aziendale originale.
- `output/risultato_massivo_recuperato.xlsx`: risultato massivo corrente da aggiornare sempre nello stesso file.
- `output/risultato_massivo_recuperato.summary.json`: metriche macchina dell’ultima esportazione.
- `stato_massivo/ricerca.sqlite`: coda e stato persistente; è essenziale per la ripresa sul PC che esegue la run.
- `cv_medici/`: CV verificati, denominati con codice persona, cognome e nome.
- `cv_medici_da_verificare/`: documenti non automaticamente attribuibili.
- `logs/run_medici_*.log`: log dettagliati di ogni esecuzione.
- `RELAZIONE_PROGETTO_MEDICI.md`: relazione aziendale da aggiornare dopo ogni controllo dei risultati.
- `tests/`: suite di regressione.

### Classificazione da preservare

Gli stati principali sono:

- `SPECIALITA_DOCUMENTATA`: specialità esplicita e identità forte tramite codice fiscale o data di nascita concordante;
- `SPECIALITA_CON_IDENTITA_NOMINALE`: specialità esplicita, ma identità basata soltanto sul nome completo;
- `DISCIPLINA_DICHIARATA_NEL_PROFILO`: disciplina, categoria o unità clinica pubblicata dalla fonte; non equivale automaticamente a diploma;
- `IDENTITA_CONFERMATA_SENZA_TITOLO`: identità forte ma nessuna specialità valida;
- `PROFILO_CON_IDENTITA_NOMINALE`: profilo nominale senza titolo sufficiente;
- `EVIDENZA_NON_CONFERMATA`: evidenza insufficiente o ambigua;
- `FONTE_NON_ACCESSIBILE`: fonte non scaricabile;
- `NON_COPERTO_DALLE_FONTI`: nessun candidato dalle fonti integrate.

Non sommare categorie già comprese nel totale utile. Il totale utile è la somma delle prime tre categorie di risultato: documentata, nominale e disciplina dichiarata.

Regole di qualità:

1. Nome e cognome devono concordare con il profilo o il documento.
2. In presenza di omonimia, senza CF o data di nascita non confermare automaticamente.
3. Un’unità operativa, una prestazione o un interesse clinico valgono al massimo come disciplina dichiarata.
4. Una specialità documentata richiede una formula esplicita come “Diploma di Specializzazione in…” o equivalente e identità forte.
5. “Specializzando”, scuola in corso, docenza presso una scuola o occupazione desiderata non dimostrano un titolo conseguito.
6. Conserva sempre URL, tipo fonte ed evidenza testuale.
7. Rispetta `robots.txt`, limiti HTTP e domini autorizzati.
8. Non disabilitare globalmente TLS: eventuali eccezioni devono essere motivate, testate e limitate al singolo host.

### Fonti e problemi già affrontati

Il catalogo contiene 25 fonti, 24 attive. Include portali professionali e fonti istituzionali, tra cui Humanitas, Gemelli, IEO, MultiMedica, Gruppo San Donato, San Raffaele, Niguarda, San Matteo, Campus Bio-Medico e Sant’Andrea.

`Campus_Bio_Medico` usa l’API WordPress pubblica e può leggere codice fiscale e biografie. Ha aumentato le specialità documentate da 18 a 46.

`Citta_Salute_Torino_ALPI` è presente ma disabilitata. Il PDF ufficiale è sotto `/images/`, percorso vietato dal `robots.txt`; non aggirare il divieto. Il parser PDF e il test sono conservati per un’eventuale importazione locale autorizzata.

Sant’Andrea è la fonte attiva più recente. Dopo il ciclo conclusivo, analizza quanti dei 65 candidati sono stati elaborati e quanti risultati nuovi sono stati ottenuti. Se HTTP è zero e la fonte è `done`, non chiedere un’altra run identica.

### Strategia per le fonti successive

Prima di integrare una fonte:

1. verifica `robots.txt`, accessibilità TLS e stabilità degli URL;
2. conta i profili realmente enumerabili;
3. confronta localmente i nomi con i 45.530 record non coperti;
4. stima resa, numero di richieste e qualità dell’evidenza;
5. integra soltanto fonti con un rapporto utile tra costo e risultati;
6. aggiungi test per il formato specifico;
7. usa il comando Python standard e una sola run quando possibile.

Le città con più record scoperti erano soprattutto Roma, Milano, Torino, Napoli, Palermo, Bari, Bologna, Catania, Padova, Firenze e Genova. Il Sant’Andrea è stato scelto perché pubblica profili individuali e consente il crawling. Prima erano stati valutati anche Tor Vergata e Città della Salute.

Evita di promettere grandi incrementi senza una stima locale. Con fonti istituzionali singole, gli incrementi recenti sono generalmente nell’ordine di decine o poche centinaia, perché gran parte dei profili pubblici si sovrappone ai record già coperti.

### Modalità di lavoro con l’utente

L’utente vuole risultati concreti e si irrita quando riceve passaggi frammentati o cicli inutili. Quindi:

- completa insieme codice, configurazione, test e documentazione prima di chiedere una run;
- fornisci un solo comando Python preciso;
- non creare un `.bat` per ogni fonte;
- non chiedere di rilanciare quando HTTP è zero o la coda è già completa;
- quando l’utente scrive “aggiornato”, controlla subito Git, ultimo log, summary ed Excel;
- se il Mac è indietro rispetto a `origin/main`, usa `git pull --ff-only origin main` prima dell’analisi;
- riporta l’incremento rispetto alla run precedente e non soltanto i totali;
- correggi i problemi prima di chiedere una nuova esecuzione;
- mantieni risposte brevi e operative.

### Aggiornamento dei report

Dopo ogni nuova run o controllo dei risultati aggiorna sempre:

1. `RELAZIONE_PROGETTO_MEDICI.md`, direttamente nel repository;
2. il report Excel fisso con tutti i medici utili e i CV, se l’ambiente in cui operi dispone del generatore già usato.

Nel precedente ambiente Codex il report fisso era:

`/Users/niko/.codex/.chatgpt-projects/g-p-6a8ee14119908191985d36a2e1abe4bb/outputs/report_specialita_20260909/Report_specialita_medici_20260909.xlsx`

I generatori erano `build_report_medici.py` e `build_report_medici.mjs` nella cartella del progetto ChatGPT locale, non necessariamente disponibili a Claude. Se non sono disponibili, crea nel repository un generatore Python riproducibile invece di dipendere da percorsi privati esterni. Il report deve mantenere almeno:

- foglio riepilogo;
- elenco completo dei medici con specialità o disciplina;
- elenco dei CV utili;
- data di generazione e data dell’ultima acquisizione;
- distinzione tra attendibilità alta, media e disciplina dichiarata.

Non inserire nel repository file tecnici di ispezione enormi. In precedenza un `.xlsx.inspect.ndjson` da circa 225 MB veniva eliminato dopo la verifica.

### Verifiche tecniche minime

Prima di consegnare modifiche:

```bash
python3 -m unittest discover -s tests
python3 -m py_compile medici_massivo.py scriptMedici.py
git diff --check
git status -sb
```

Non interpretare i warning prodotti dai test intenzionali di gestione errori come fallimenti: conta l’esito finale della suite.

Dopo una run leggi almeno:

```powershell
Get-Content .\output\risultato_massivo_recuperato.summary.json
Get-Content .\logs\run_medici_ULTIMO_TIMESTAMP.log
```

Verifica:

- associazioni per fonte;
- problemi dettagliati;
- documenti elaborati;
- richieste HTTP;
- Search API uguale a zero;
- variazione degli stati;
- eventuali URL ancora pending nella coda, se lo stato SQLite è disponibile.

Continua ora dal repository reale, preservando le modifiche locali descritte sopra. La prima priorità è pubblicare la correzione dei 404 Sant’Andrea, completare l’ultimo piccolo ciclo della fonte, analizzarne il risultato e aggiornare entrambi i report. Successivamente seleziona e integra la fonte istituzionale con la migliore resa stimata senza Search API.

## FINE PROMPT

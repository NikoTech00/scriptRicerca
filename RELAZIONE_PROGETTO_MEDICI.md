# Relazione tecnica — Motore di ricerca medici, specialità e CV

**Progetto:** arricchimento dell’archivio aziendale dei medici  
**Versione applicativa:** V6.0  
**Ultimo aggiornamento dei risultati:** 11 settembre 2026, ore 12:17  
**Archivio analizzato:** 81.331 persone

## 1. Scopo del progetto

Il progetto automatizza la ricerca e la raccolta di informazioni professionali pubbliche relative ai medici presenti nell’archivio Excel aziendale. L’obiettivo principale è ottenere, per il maggior numero possibile di nominativi:

- la specialità medica dichiarata o documentata;
- la disciplina professionale riportata da ospedali e portali sanitari;
- l’eventuale curriculum vitae pubblico;
- la fonte e il testo che sostengono ogni risultato;
- un livello esplicito di attendibilità dell’associazione tra fonte e persona.

Il sistema non sostituisce una certificazione amministrativa del titolo. Produce un archivio tracciabile che separa le informazioni documentate da quelle associate soltanto tramite nome e cognome, così che l’azienda possa usare i risultati in modo proporzionato alla loro affidabilità.

## 2. Risultato raggiunto

Alla data della relazione, l’elaborazione dell’intero archivio ha prodotto:

| Indicatore | Record | Quota sull’archivio |
|---|---:|---:|
| Archivio originale | 81.331 | 100,0% |
| Record con specialità o disciplina utilizzabile | 29.296 | 36,0% |
| Disciplina dichiarata in un profilo pubblico | 23.600 | 29,0% |
| Specialità con identità nominale | 5.678 | 7,0% |
| Specialità documentata con identità anagrafica forte | 18 | 0,02% |
| Persone collegate ad almeno una fonte candidata | 35.669 | 43,9% |
| Profili nominali senza titolo sufficiente | 2.692 | 3,3% |
| Evidenze non confermate | 3.556 | 4,4% |
| Fonti non accessibili | 103 | 0,1% |
| Record non coperti dalle fonti integrate | 45.662 | 56,1% |

Sono stati elaborati e memorizzati 74.184 URL; 776 URL risultano in errore. Il flusso massivo non consuma crediti Serper o di altre Search API. I risultati sopra riportati sono categorie distinte e non devono essere sommati nuovamente tra loro.

## 3. Ultimo controllo operativo

L’acquisizione più recente registrata è dell’**11 settembre 2026 alle ore 12:17**. Il controllo ha confermato **29.296 risultati utilizzabili**, senza variazioni rispetto all’acquisizione precedente. La run ha effettuato zero richieste HTTP e non ha prodotto righe di discovery delle fonti: ha quindi esportato lo stato già memorizzato senza eseguire il recupero istituzionale programmato.

Il prossimo ciclo deve essere avviato con il catalogo `fonti_recupero_istituzionali.json`, l’opzione `--refresh-sources` e senza `--export-only` o `--skip-discovery`. L’esito atteso nel log comprende le righe `FONTE Gemelli`, `FONTE IEO` e `FONTE MultiMedica`.

Per evitare errori di trascrizione su Windows è disponibile `avvia_recupero_istituzionale.bat`, che avvia il ciclo con i parametri corretti e mantiene aperta la finestra al termine. Il log registra anche catalogo, limiti e opzioni effettivamente ricevute.

## 4. Funzionamento generale

La pipeline opera in sei fasi:

1. **Importazione dell’Excel.** Legge tutte le righe e conserva le colonne originarie. Costruisce un indice locale dei nomi e individua gli omonimi presenti nell’intero archivio.
2. **Acquisizione degli indici pubblici.** Legge sitemap e pagine indice delle fonti configurate, senza interrogare motori di ricerca a pagamento.
3. **Associazione dei candidati.** Collega un profilo a un record quando l’URL o il contenuto espongono un nome completo compatibile. Gli URL numerici vengono prima aperti e poi associati dal contenuto.
4. **Analisi dei profili e dei documenti.** Estrae HTML, PDF, DOCX e alcuni DOC legacy. Cerca campi strutturati, sezioni curriculari e formule esplicite relative a specialità e discipline.
5. **Verifica dell’identità e classificazione.** Confronta nome, eventuale data di nascita e codice fiscale; gestisce separatamente omonimie, incompatibilità e fonti insufficienti.
6. **Esportazione e tracciabilità.** Aggiorna l’Excel, conserva fonte ed evidenza testuale, salva i CV con codice persona, cognome e nome e produce i fogli riepilogativi `Copertura` e `Fonti`.

```mermaid
flowchart LR
    A[Excel aziendale] --> B[Indice locale dei nominativi]
    C[Sitemap e indici pubblici] --> D[Profili e CV candidati]
    B --> E[Associazione per identità]
    D --> E
    E --> F[Estrazione di titoli e discipline]
    F --> G[Classificazione dell’attendibilità]
    G --> H[Excel completo, evidenze e CV]
    G --> I[Cache e coda SQLite]
```

## 5. Fonti integrate

Il catalogo massivo contiene attualmente 21 fonti. Comprende strutture sanitarie e portali professionali pubblici:

- Humanitas, Humanitas San Pio X, Policlinico Gemelli, IEO e MultiMedica;
- Gruppo San Donato, Ospedale San Raffaele e Policlinico San Matteo;
- Istituto Ortopedico Rizzoli, Auxologico, Maugeri e Santagostino;
- ASST Rhodense e ASST Bergamo Ovest;
- MioDottore, iDoctors, Top Doctors, Doctolib, Medicitalia, PagineMediche/Visitami e MiAgenda.

Per ogni fonte il file `fonti_massive.json` definisce domini ammessi, tipo di indice e regole di selezione. Il motore accetta solo URL pubblici appartenenti ai domini configurati, applica ritardi tra le richieste e registra gli errori senza trasformarli in risultati negativi.

Le fonti non hanno tutte lo stesso valore probatorio. Un profilo ospedaliero può dichiarare l’unità o la disciplina clinica; un CV può dichiarare un titolo conseguito; un portale professionale può pubblicare una categoria indicata dal professionista. Il report conserva questa differenza.

## 6. Criteri di qualità e attendibilità

### Specialità documentata

Il documento contiene una dichiarazione esplicita del titolo e l’identità è sostenuta da un dato anagrafico concordante, come data di nascita o codice fiscale. È il livello automatico più forte.

### Specialità con identità nominale

La fonte contiene una dichiarazione esplicita della specialità e il nome completo coincide, ma non sono disponibili dati sufficienti per una verifica anagrafica forte. Il dato è utile come proposta, con controllo umano nei casi sensibili.

### Disciplina dichiarata nel profilo

La fonte assegna pubblicamente il medico a una disciplina o categoria professionale. È un’informazione operativa attendibile sulla sua attività dichiarata, ma non dimostra da sola il conseguimento del relativo diploma di specializzazione.

### Evidenza non confermata

Comprende omonimi non disambiguabili, dati incompatibili, documenti privi di testo, contenuti insufficienti o titoli estratti fuori dalla tassonomia. Questi record restano separati dai risultati utilizzabili.

Il sistema non considera automaticamente come titolo conseguito un reparto, una pubblicazione, un insegnamento, l’iscrizione a una scuola di specializzazione o una posizione professionale desiderata. Il vocabolario normalizza varianti linguistiche verso denominazioni comuni, conservando l’evidenza originale.

## 7. Gestione dei CV

I documenti pubblici vengono scaricati una sola volta e conservati nella cache. Quando un CV è attribuibile a una persona, la copia leggibile viene salvata con il formato:

```text
CODICE_PERSONA_Cognome_Nome.estensione
```

PDF e DOCX vengono analizzati direttamente. I DOC legacy utilizzano i convertitori disponibili sul sistema. I PDF senza testo incorporato richiedono OCR o verifica manuale e non vengono promossi automaticamente a CV verificati.

La presenza di un CV e la presenza di una specialità sono due informazioni indipendenti: un CV può non dichiarare chiaramente un titolo, mentre un profilo istituzionale può dichiarare una disciplina senza pubblicare il CV.

## 8. Architettura e componenti principali

| Componente | Responsabilità |
|---|---|
| `scriptMedici.py` | Interfaccia da riga di comando, tassonomia, modalità tradizionale e recuperi locali |
| `medici_massivo.py` | Discovery delle fonti, coda persistente, download, analisi, classificazione ed export |
| `fonti_massive.json` | Catalogo completo delle fonti pubbliche |
| `fonti_recupero_istituzionali.json` | Catalogo mirato per recuperi e riesami istituzionali |
| `stato_massivo/` | Database SQLite, cache dei documenti e stato di ripresa |
| `cv_medici/` | CV attribuiti e salvati con denominazione leggibile |
| `cv_medici_da_verificare/` | Documenti candidati che richiedono controllo |
| `output/` | Excel completi, riepiloghi JSON e risultati intermedi |
| `logs/` | Log cronologici delle esecuzioni |
| `tests/` | Test automatici del motore e della gestione degli errori |

Le dipendenze principali sono Python 3.10 o successivo, Requests, Beautiful Soup, OpenPyXL, PyPDF e Olefile.

## 9. Ripresa, prestazioni e affidabilità operativa

Lo stato di lavorazione risiede in SQLite. La pipeline può essere interrotta e rilanciata senza ricominciare l’intero archivio. Ogni documento completato viene registrato; l’Excel viene aggiornato periodicamente tramite checkpoint e nuovamente alla chiusura.

Le copie locali sono identificate con impronta SHA-256. Se una pagina non è cambiata, può essere rianalizzata senza un nuovo download. Le modifiche ai parser possono attivare il riesame mirato dei soli profili interessati. Un blocco di esecuzione impedisce a due processi di modificare contemporaneamente lo stesso stato.

L’output viene scritto in modo atomico: prima viene creato un file completo temporaneo e poi viene sostituito il report precedente. Se Excel tiene aperto il file o il sistema impedisce la sostituzione, la copia completa di recupero viene conservata e segnalata.

## 10. Esecuzione ordinaria

Installazione iniziale:

```powershell
python -m pip install -r requirements.txt
```

Esecuzione massiva o prosecuzione della coda:

```powershell
python .\scriptMedici.py "input\scriptMedici.xlsx" --massivo --max-documents 5000 --max-http-requests 6000 --mass-checkpoint-every 500 --output "output\risultato_massivo_recuperato.xlsx"
```

Esportazione dello stato senza rete:

```powershell
python .\scriptMedici.py "input\scriptMedici.xlsx" --massivo --export-only --output "output\risultato_massivo_recuperato.xlsx"
```

Verifica automatica del software:

```powershell
python -B -m unittest discover -s tests -q
```

La suite corrente comprende 71 test e copre, tra gli altri aspetti, gestione della coda, ripresa, errori di rete, salvataggio atomico, omonimie, normalizzazione delle discipline, fonti istituzionali e denominazione dei CV.

## 11. Protezione e trattamento dei dati

L’Excel, il database di stato, i CV e i risultati rimangono sui computer aziendali. La modalità massiva non invia l’archivio o i codici fiscali a Search API. Le richieste di rete sono rivolte alle pagine pubbliche configurate nel catalogo; i relativi URL possono contenere il nome pubblico del professionista.

Il progetto raccoglie esclusivamente informazioni professionali rese pubbliche dalle fonti consultate. Conservazione, accesso, periodo di mantenimento e uso aziendale dei dati devono seguire le regole interne e le indicazioni del responsabile del trattamento. La tracciabilità della fonte permette la revisione e l’eventuale aggiornamento del singolo dato.

## 12. Limiti attuali

Il principale limite non è la capacità di elaborazione, ma la disponibilità di fonti pubbliche indicizzabili e sufficientemente strutturate. I 45.662 record non coperti non rappresentano medici inesistenti o privi di specialità: indicano che le fonti attualmente integrate non hanno prodotto un candidato associabile.

Altri limiti materiali sono:

- omonimi senza data di nascita o codice fiscale pubblico;
- profili che riportano soltanto la qualifica generica di medico;
- pagine protette da login, CAPTCHA o restrizioni tecniche;
- PDF acquisiti come immagini e privi di testo estraibile;
- differenza tra disciplina clinica esercitata e titolo accademico conseguito;
- aggiornamento e rimozione delle pagine da parte delle fonti esterne.

Un incremento consistente oltre la copertura attuale richiede nuove fonti nominative pubbliche, elenchi istituzionali utilizzabili o documenti aziendali autorizzati. Aumentare soltanto il numero di thread o cambiare linguaggio di programmazione non crea evidenze che non sono pubblicate.

## 13. Sviluppi consigliati

Le attività con il miglior rapporto tra tempo e risultato sono:

1. completare il recupero mirato dei profili istituzionali già associati ma privi di disciplina;
2. integrare nuovi ospedali e gruppi sanitari dotati di elenco nominativo pubblico;
3. aggiungere OCR locale per i PDF scansionati più promettenti;
4. creare una coda di revisione umana prioritaria per omonimi e specialità nominali ad alto valore;
5. programmare aggiornamenti periodici delle sitemap e il confronto tra acquisizioni;
6. definire con l’azienda quali livelli di attendibilità possono essere usati direttamente nei processi interni.

## 14. Conclusione

Il progetto ha trasformato una ricerca manuale non sostenibile su 81.331 nominativi in una pipeline ripetibile, tracciabile e senza costi di Search API. Il risultato corrente rende disponibili specialità o discipline per 29.296 persone, conserva l’origine di ogni informazione e mantiene separati i casi che richiedono verifica.

Il software è già utilizzabile per produrre report aziendali e proseguire l’arricchimento incrementale. Il prossimo aumento significativo della copertura dipenderà soprattutto dall’integrazione di nuove fonti pubbliche e dalla risoluzione controllata dei casi nominali, mantenendo gli attuali criteri di qualità.

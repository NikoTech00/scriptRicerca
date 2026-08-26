from __future__ import annotations

import argparse
import json
import logging
import os
import random
import shutil
import tempfile
import time
from dataclasses import dataclass
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from google import genai
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


# ============================================================
# CONFIGURAZIONE
# ============================================================

INPUT_COLUMNS = (
    "Cognome",
    "Nome",
    "Data di Nascita",
)

RESULT_COLUMNS = (
    "Stato ricerca",
    "Struttura",
    "Reparto/UO",
    "Citta",
    "Email professionale",
    "Telefono professionale",
    "Ruolo trovato",
    "Confidenza",
    "Note",
    "Fonti",
    "Tentativi",
    "Ultimo aggiornamento UTC",
    "Errore",
)

TERMINAL_STATUSES = {
    "COMPLETATO",
    "NESSUN_RISULTATO",
}

DEFAULT_ROLE = "farmacista ospedaliero/a"

DEFAULT_MODEL = "gemini-3.7-flash"


# ============================================================
# MODELLI DATI
# ============================================================

@dataclass(frozen=True)
class Person:
    row: int
    surname: str
    name: str
    birth_date: str


# ============================================================
# SCHEMA JSON RISPOSTA GEMINI
# ============================================================

RESULT_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "found": {
            "type": "boolean"
        },
        "facility": {
            "type": "string"
        },
        "department": {
            "type": "string"
        },
        "city": {
            "type": "string"
        },
        "professional_email": {
            "type": "string"
        },
        "professional_phone": {
            "type": "string"
        },
        "role_found": {
            "type": "string"
        },
        "confidence": {
            "type": "string",
            "enum": [
                "alta",
                "media",
                "bassa",
                "nessuna",
            ],
        },
        "notes": {
            "type": "string"
        },
    },
    "required": [
        "found",
        "facility",
        "department",
        "city",
        "professional_email",
        "professional_phone",
        "role_found",
        "confidence",
        "notes",
    ],
}


# ============================================================
# ARGOMENTI CLI
# ============================================================

def parse_args() -> argparse.Namespace:

    parser = argparse.ArgumentParser(
        description=(
            "Ricerca automatizzata di struttura e contatti "
            "professionali pubblici di farmacisti ospedalieri."
        )
    )

    parser.add_argument(
        "input",
        type=Path,
        help="File Excel .xlsx da elaborare",
    )

    parser.add_argument(
        "--output",
        type=Path,
        help=(
            "File Excel di output. "
            "Default: <input>_risultati.xlsx"
        ),
    )

    parser.add_argument(
        "--sheet",
        default="Dati",
        help="Nome del foglio Excel da elaborare (default: Dati)",
    )

    parser.add_argument(
        "--model",
        default=os.getenv(
            "GEMINI_MODEL",
            DEFAULT_MODEL,
        ),
        help=f"Modello Gemini (default: {DEFAULT_MODEL})",
    )

    # --limit è il comando che userai normalmente.
    # --max-rows rimane come alias per compatibilità.
    parser.add_argument(
        "--limit",
        "--max-rows",
        dest="max_rows",
        type=int,
        default=None,
        help=(
            "Numero massimo di persone da elaborare "
            "in questa esecuzione. Es: --limit 5"
        ),
    )

    parser.add_argument(
        "--start-row",
        type=int,
        default=2,
        help=(
            "Prima riga Excel da considerare. "
            "La riga 1 contiene le intestazioni. "
            "Default: 2"
        ),
    )

    parser.add_argument(
        "--max-retries",
        type=int,
        default=4,
        help=(
            "Numero massimo di tentativi per persona "
            "(default: 4)"
        ),
    )

    parser.add_argument(
        "--delay",
        type=float,
        default=2.0,
        help=(
            "Secondi di pausa tra una persona e la successiva "
            "(default: 2)"
        ),
    )

    parser.add_argument(
        "--retry-errors",
        action="store_true",
        help=(
            "Riprova anche le righe che nel file di output "
            "hanno Stato ricerca=ERRORE"
        ),
    )

    args = parser.parse_args()

    # --------------------------------------------------------
    # Validazione parametri
    # --------------------------------------------------------

    if args.max_rows is not None and args.max_rows <= 0:
        parser.error("--limit deve essere maggiore di 0")

    if args.start_row < 2:
        parser.error("--start-row deve essere almeno 2")

    if args.max_retries <= 0:
        parser.error("--max-retries deve essere maggiore di 0")

    if args.delay < 0:
        parser.error("--delay non può essere negativo")

    return args


# ============================================================
# FUNZIONI GENERALI
# ============================================================

def clean(value: Any) -> str:
    if value is None:
        return ""

    return str(value).strip()


def format_birth_date(value: Any) -> str:

    if isinstance(value, (datetime, date)):
        return value.strftime("%d/%m/%Y")

    return clean(value)


def utc_now() -> str:
    return datetime.now(
        timezone.utc
    ).isoformat(
        timespec="seconds"
    )


# ============================================================
# EXCEL - HEADERS
# ============================================================

def normalized_headers(ws) -> dict[str, int]:

    headers: dict[str, int] = {}

    for cell in ws[1]:

        label = clean(cell.value)

        if label:
            headers[label.casefold()] = cell.column

    return headers


def require_input_columns(
    headers: dict[str, int]
) -> None:

    missing = [
        name
        for name in INPUT_COLUMNS
        if name.casefold() not in headers
    ]

    if missing:
        raise ValueError(
            "Colonne obbligatorie mancanti: "
            + ", ".join(missing)
        )


def ensure_result_columns(
    ws
) -> dict[str, int]:

    headers = normalized_headers(ws)

    for label in RESULT_COLUMNS:

        key = label.casefold()

        if key in headers:
            continue

        col = ws.max_column + 1

        cell = ws.cell(
            row=1,
            column=col,
            value=label,
        )

        cell.font = Font(
            bold=True,
            color="FFFFFF",
        )

        cell.fill = PatternFill(
            "solid",
            fgColor="1F4E78",
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

        headers[key] = col

    ws.freeze_panes = "A2"

    return headers


# ============================================================
# FILE OUTPUT
# ============================================================

def output_path_for(
    input_path: Path,
    requested: Path | None,
) -> Path:

    if requested is not None:
        return requested

    return input_path.with_name(
        f"{input_path.stem}_risultati.xlsx"
    )


def prepare_output(
    input_path: Path,
    output_path: Path,
) -> None:

    if not input_path.exists():
        raise FileNotFoundError(
            f"File di input non trovato: {input_path}"
        )

    if input_path.suffix.lower() != ".xlsx":
        raise ValueError(
            "Il file di input deve essere .xlsx"
        )

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    # Se esiste già NON viene sovrascritto.
    # Questo permette il resume.
    if output_path.exists():

        logging.info(
            "Riprendo file di output esistente: %s",
            output_path,
        )

        return

    if input_path.resolve() == output_path.resolve():
        return

    shutil.copy2(
        input_path,
        output_path,
    )

    logging.info(
        "Creata copia di lavoro: %s",
        output_path,
    )


# ============================================================
# SALVATAGGIO ATOMICO
# ============================================================

def atomic_save(
    workbook,
    output_path: Path,
) -> None:
    """
    Salva prima in un file temporaneo e poi sostituisce
    l'output.

    In questo modo riduciamo il rischio di corrompere
    l'Excel se lo script viene interrotto durante il save.
    """

    fd, temp_name = tempfile.mkstemp(
        prefix=f".{output_path.stem}_",
        suffix=".xlsx",
        dir=output_path.parent,
    )

    os.close(fd)

    temp_path = Path(temp_name)

    try:

        workbook.save(
            temp_path
        )

        os.replace(
            temp_path,
            output_path,
        )

    finally:

        temp_path.unlink(
            missing_ok=True
        )


# ============================================================
# LETTURA PERSONA DA EXCEL
# ============================================================

def person_from_row(
    ws,
    row: int,
    headers: dict[str, int],
) -> Person:

    return Person(

        row=row,

        surname=clean(
            ws.cell(
                row,
                headers["cognome"],
            ).value
        ),

        name=clean(
            ws.cell(
                row,
                headers["nome"],
            ).value
        ),

        birth_date=format_birth_date(
            ws.cell(
                row,
                headers["data di nascita"],
            ).value
        ),
    )


# ============================================================
# PROMPT
# ============================================================

def build_prompt(
    person: Person,
) -> str:

    return f"""
Devi effettuare una ricerca web accurata su una persona che
lavora o ha lavorato come farmacista ospedaliero.

DATI DISPONIBILI

Cognome: {person.surname}
Nome: {person.name}
Data di nascita: {person.birth_date or "non disponibile"}
Professione attesa: {DEFAULT_ROLE}

OBIETTIVO

Identifica, se possibile, la struttura sanitaria presso cui
questa persona lavora attualmente.

Se non è possibile determinare quella attuale, individua la
struttura professionale più recente verificabile.

Cerca in particolare:

- ospedali;
- aziende ospedaliere;
- ASST;
- ATS;
- ASL;
- AUSL;
- aziende sanitarie;
- IRCCS;
- policlinici;
- farmacie ospedaliere;
- università;
- servizi sanitari regionali;
- enti sanitari pubblici o privati.

INFORMAZIONI RICHIESTE

- struttura sanitaria;
- reparto, UO o farmacia ospedaliera;
- città;
- ruolo professionale trovato;
- email professionale pubblicamente disponibile;
- telefono professionale pubblicamente disponibile.

FONTI

Dai priorità assoluta a:

1. siti ufficiali di ospedali e aziende sanitarie;
2. ASL / AUSL / ATS / ASST;
3. Regioni e Servizi Sanitari Regionali;
4. amministrazione trasparente;
5. documenti e PDF istituzionali;
6. Ordine dei Farmacisti;
7. università;
8. società scientifiche;
9. documenti ufficiali relativi a concorsi, incarichi,
   nomine o delibere.

REGOLE IMPORTANTI

1. Non inventare informazioni.

2. Non dedurre indirizzi email utilizzando pattern aziendali.

Ad esempio, NON generare:

nome.cognome@ospedale.it

a meno che quell'indirizzo non sia realmente pubblicato
in una fonte.

3. Non cercare o restituire:

- email personali;
- numeri telefonici personali;
- indirizzi di abitazione;
- altri recapiti privati.

4. Sono ammessi esclusivamente contatti professionali
pubblicamente disponibili.

5. Se non esiste un contatto nominativo ma è pubblicato
il contatto della farmacia ospedaliera, del reparto o della
struttura, puoi restituire quello.

6. Utilizza la data di nascita esclusivamente per
disambiguare eventuali omonimi.

7. Non riportare la data di nascita nelle note.

8. found=true soltanto quando esistono evidenze
ragionevoli che colleghino quella specifica persona alla
struttura o alla professione.

9. Se esistono omonimi o dubbi sull'identità, abbassa
la confidenza e spiegalo brevemente nelle note.

CONFIDENZA

alta:
identità e struttura supportate chiaramente da fonti
affidabili.

media:
evidenze forti ma non completamente definitive.

bassa:
possibile corrispondenza, ma esistono dubbi.

nessuna:
nessun risultato sufficientemente verificabile.

Se un'informazione non è verificabile, restituisci il relativo
campo come stringa vuota.
""".strip()


# ============================================================
# ESTRAZIONE FONTI GEMINI
# ============================================================

def extract_source_urls(
    interaction: Any
) -> list[str]:

    urls: list[str] = []

    for step in (
        getattr(
            interaction,
            "steps",
            None,
        )
        or []
    ):

        if getattr(
            step,
            "type",
            None,
        ) != "model_output":
            continue

        for block in (
            getattr(
                step,
                "content",
                None,
            )
            or []
        ):

            for annotation in (
                getattr(
                    block,
                    "annotations",
                    None,
                )
                or []
            ):

                if getattr(
                    annotation,
                    "type",
                    None,
                ) != "url_citation":
                    continue

                url = clean(
                    getattr(
                        annotation,
                        "url",
                        "",
                    )
                )

                if url and url not in urls:
                    urls.append(url)

    return urls


# ============================================================
# RICERCA GEMINI
# ============================================================

def research_person(
    client: genai.Client,
    model: str,
    person: Person,
) -> dict[str, Any]:
    """
    Non accede direttamente all'Excel.

    Questo è intenzionale: in futuro questa funzione
    potrà essere eseguita tranquillamente da worker paralleli.
    """

    interaction = client.interactions.create(

        model=model,

        input=build_prompt(
            person
        ),

        tools=[
            {
                "type": "google_search"
            }
        ],

        response_format={
            "type": "text",
            "mime_type": "application/json",
            "schema": RESULT_SCHEMA,
        },
    )

    output_text = clean(
        getattr(
            interaction,
            "output_text",
            "",
        )
    )

    if not output_text:
        raise ValueError(
            "Gemini ha restituito una risposta vuota."
        )

    try:

        result = json.loads(
            output_text
        )

    except json.JSONDecodeError as exc:

        raise ValueError(
            "Gemini non ha restituito JSON valido."
        ) from exc

    if not isinstance(
        result,
        dict,
    ):
        raise ValueError(
            "La risposta Gemini non è un oggetto JSON."
        )

    # --------------------------------------------------------
    # Fonti reali citate dal grounding Google
    # --------------------------------------------------------

    result["sources"] = extract_source_urls(
        interaction
    )

    return result


# ============================================================
# RETRY
# ============================================================

def research_with_retry(
    client: genai.Client,
    model: str,
    person: Person,
    max_retries: int,
) -> tuple[dict[str, Any], int]:

    last_error: Exception | None = None

    for attempt in range(
        1,
        max_retries + 1,
    ):

        try:

            result = research_person(
                client,
                model,
                person,
            )

            return (
                result,
                attempt,
            )

        except (
            KeyboardInterrupt,
            SystemExit,
        ):
            raise

        except Exception as exc:

            last_error = exc

            logging.warning(
                "Riga %s | tentativo %s/%s fallito: %s",
                person.row,
                attempt,
                max_retries,
                exc,
            )

            if attempt == max_retries:
                break

            # exponential backoff + jitter
            wait = min(
                60.0,
                (2 ** (attempt - 1))
                + random.uniform(
                    0.0,
                    1.0,
                ),
            )

            logging.info(
                "Nuovo tentativo tra %.1f secondi...",
                wait,
            )

            time.sleep(
                wait
            )

    raise RuntimeError(
        str(last_error)
        if last_error
        else "Errore sconosciuto"
    )


# ============================================================
# SCRITTURA EXCEL
# ============================================================

def set_cell(
    ws,
    row: int,
    headers: dict[str, int],
    label: str,
    value: Any,
) -> None:

    ws.cell(
        row=row,
        column=headers[
            label.casefold()
        ],
        value=value,
    )


def write_result(
    ws,
    row: int,
    headers: dict[str, int],
    result: dict[str, Any],
    attempts: int,
) -> None:

    found = bool(
        result.get(
            "found",
            False,
        )
    )

    status = (
        "COMPLETATO"
        if found
        else "NESSUN_RISULTATO"
    )

    sources = result.get(
        "sources"
    ) or []

    values = {

        "Stato ricerca":
            status,

        "Struttura":
            clean(
                result.get(
                    "facility"
                )
            ),

        "Reparto/UO":
            clean(
                result.get(
                    "department"
                )
            ),

        "Citta":
            clean(
                result.get(
                    "city"
                )
            ),

        "Email professionale":
            clean(
                result.get(
                    "professional_email"
                )
            ),

        "Telefono professionale":
            clean(
                result.get(
                    "professional_phone"
                )
            ),

        "Ruolo trovato":
            clean(
                result.get(
                    "role_found"
                )
            ),

        "Confidenza":
            clean(
                result.get(
                    "confidence"
                )
            ),

        "Note":
            clean(
                result.get(
                    "notes"
                )
            ),

        "Fonti":
            "\n".join(
                sources
            ),

        "Tentativi":
            attempts,

        "Ultimo aggiornamento UTC":
            utc_now(),

        "Errore":
            "",
    }

    for label, value in values.items():

        set_cell(
            ws,
            row,
            headers,
            label,
            value,
        )


def write_error(
    ws,
    row: int,
    headers: dict[str, int],
    error: Exception,
    attempts: int,
) -> None:

    set_cell(
        ws,
        row,
        headers,
        "Stato ricerca",
        "ERRORE",
    )

    set_cell(
        ws,
        row,
        headers,
        "Tentativi",
        attempts,
    )

    set_cell(
        ws,
        row,
        headers,
        "Ultimo aggiornamento UTC",
        utc_now(),
    )

    set_cell(
        ws,
        row,
        headers,
        "Errore",
        clean(error)[:2000],
    )


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    # Prima dotenv, poi parse_args:
    # in questo modo GEMINI_MODEL può essere letto dal .env.
    load_dotenv()

    args = parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format=(
            "%(asctime)s | "
            "%(levelname)s | "
            "%(message)s"
        ),
        datefmt="%H:%M:%S",
    )

    # --------------------------------------------------------
    # API KEY
    # --------------------------------------------------------

    api_key = os.getenv(
        "GEMINI_API_KEY"
    )

    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY non trovata.\n"
            "Inseriscila nel file .env:\n"
            "GEMINI_API_KEY=xxxxxxxx"
        )

    # --------------------------------------------------------
    # PATH
    # --------------------------------------------------------

    input_path = (
        args.input
        .expanduser()
        .resolve()
    )

    output_path = output_path_for(
        input_path,
        args.output,
    )

    output_path = (
        output_path
        .expanduser()
        .resolve()
    )

    prepare_output(
        input_path,
        output_path,
    )

    # --------------------------------------------------------
    # APERTURA EXCEL
    # --------------------------------------------------------

    workbook = load_workbook(
        output_path
    )

    if args.sheet not in workbook.sheetnames:

        raise ValueError(
            f"Foglio '{args.sheet}' non trovato.\n"
            f"Fogli disponibili: {workbook.sheetnames}"
        )

    ws = workbook[
        args.sheet
    ]

    # --------------------------------------------------------
    # COLONNE
    # --------------------------------------------------------

    headers = normalized_headers(
        ws
    )

    require_input_columns(
        headers
    )

    headers = ensure_result_columns(
        ws
    )

    # Salviamo immediatamente le nuove colonne.
    atomic_save(
        workbook,
        output_path,
    )

    # --------------------------------------------------------
    # GEMINI
    # --------------------------------------------------------

    client = genai.Client(
        api_key=api_key
    )

    processed = 0
    skipped = 0

    logging.info(
        "============================================"
    )

    logging.info(
        "Avvio ricerca farmacisti"
    )

    logging.info(
        "Input : %s",
        input_path,
    )

    logging.info(
        "Output: %s",
        output_path,
    )

    logging.info(
        "Foglio: %s",
        args.sheet,
    )

    logging.info(
        "Modello: %s",
        args.model,
    )

    if args.max_rows is not None:

        logging.info(
            "Limite esecuzione: %s persone",
            args.max_rows,
        )

    else:

        logging.info(
            "Limite esecuzione: nessuno"
        )

    logging.info(
        "============================================"
    )

    # --------------------------------------------------------
    # CICLO SEQUENZIALE
    # --------------------------------------------------------

    start_row = max(
        2,
        args.start_row,
    )

    for row in range(
        start_row,
        ws.max_row + 1,
    ):

        # ----------------------------------------------------
        # Limite richiesto da CLI
        # ----------------------------------------------------

        if (
            args.max_rows is not None
            and processed >= args.max_rows
        ):
            break

        # ----------------------------------------------------
        # Stato precedente
        # ----------------------------------------------------

        status = clean(
            ws.cell(
                row,
                headers[
                    "stato ricerca"
                ],
            ).value
        ).upper()

        # Già completato -> salta
        if status in TERMINAL_STATUSES:

            skipped += 1

            continue

        # Riga in errore precedente:
        # viene riprovata solo con --retry-errors
        if (
            status == "ERRORE"
            and not args.retry_errors
        ):

            skipped += 1

            continue

        # ----------------------------------------------------
        # DATI PERSONA
        # ----------------------------------------------------

        person = person_from_row(
            ws,
            row,
            headers,
        )

        # Riga completamente vuota
        if (
            not person.surname
            and not person.name
        ):
            continue

        logging.info(
            "--------------------------------------------"
        )

        logging.info(
            "[%s] Riga Excel %s | %s %s | nascita: %s",
            processed + 1,
            row,
            person.name,
            person.surname,
            person.birth_date or "N/D",
        )

        # ----------------------------------------------------
        # RICERCA
        # ----------------------------------------------------

        try:

            result, attempts = research_with_retry(

                client=client,

                model=args.model,

                person=person,

                max_retries=max(
                    1,
                    args.max_retries,
                ),
            )

            write_result(
                ws,
                row,
                headers,
                result,
                attempts,
            )

            logging.info(
                "RISULTATO | %s",
                (
                    "TROVATO"
                    if result.get("found")
                    else "NESSUN RISULTATO"
                ),
            )

            logging.info(
                "Struttura: %s",
                clean(
                    result.get(
                        "facility"
                    )
                )
                or "N/D",
            )

            logging.info(
                "Reparto/UO: %s",
                clean(
                    result.get(
                        "department"
                    )
                )
                or "N/D",
            )

            logging.info(
                "Email: %s",
                clean(
                    result.get(
                        "professional_email"
                    )
                )
                or "N/D",
            )

            logging.info(
                "Telefono: %s",
                clean(
                    result.get(
                        "professional_phone"
                    )
                )
                or "N/D",
            )

            logging.info(
                "Confidenza: %s",
                clean(
                    result.get(
                        "confidence"
                    )
                )
                or "N/D",
            )

            logging.info(
                "Fonti trovate: %s",
                len(
                    result.get(
                        "sources"
                    )
                    or []
                ),
            )

        except (
            KeyboardInterrupt,
            SystemExit,
        ):

            logging.warning(
                "Interruzione richiesta. "
                "Salvataggio checkpoint..."
            )

            atomic_save(
                workbook,
                output_path,
            )

            raise

        except Exception as exc:

            write_error(
                ws,
                row,
                headers,
                exc,
                max(
                    1,
                    args.max_retries,
                ),
            )

            logging.error(
                "Riga %s non completata: %s",
                row,
                exc,
            )

        # ----------------------------------------------------
        # CHECKPOINT
        # ----------------------------------------------------

        atomic_save(
            workbook,
            output_path,
        )

        processed += 1

        logging.info(
            "Checkpoint salvato."
        )

        # ----------------------------------------------------
        # DELAY
        # ----------------------------------------------------

        if (
            args.delay > 0
            and (
                args.max_rows is None
                or processed < args.max_rows
            )
        ):
            time.sleep(
                args.delay
            )

    # --------------------------------------------------------
    # FINE
    # --------------------------------------------------------

    logging.info(
        "============================================"
    )

    logging.info(
        "ESECUZIONE COMPLETATA"
    )

    logging.info(
        "Persone elaborate: %s",
        processed,
    )

    logging.info(
        "Righe già elaborate saltate: %s",
        skipped,
    )

    logging.info(
        "File risultati: %s",
        output_path,
    )

    logging.info(
        "============================================"
    )

    return 0


if __name__ == "__main__":
    raise SystemExit(
        main()
    )
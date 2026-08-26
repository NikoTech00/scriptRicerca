from __future__ import annotations

import argparse
import io
import ipaddress
import json
import logging
import os
import random
import re
import shutil
import socket
import tempfile
import time
from dataclasses import dataclass
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup
from ddgs import DDGS
from dotenv import load_dotenv
from google import genai
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from pypdf import PdfReader


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

DEFAULT_MODEL = "gemini-3.6-flash"

DEFAULT_SEARCH_RESULTS = 8
DEFAULT_MAX_PAGES = 6
DEFAULT_MAX_CHARS_PER_PAGE = 7000
DEFAULT_MAX_CONTEXT_CHARS = 35000

HTTP_TIMEOUT = 12

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/151.0 Safari/537.36"
)

LOG_DIR = Path("logs")


# ============================================================
# MODELLI DATI
# ============================================================

@dataclass(frozen=True)
class Person:
    row: int
    surname: str
    name: str
    birth_date: str


@dataclass
class SearchResult:
    title: str
    url: str
    snippet: str
    content: str = ""


# ============================================================
# SCHEMA RISPOSTA GEMINI
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
        "source_indexes": {
            "type": "array",
            "items": {
                "type": "integer"
            },
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
        "source_indexes",
    ],
}


# ============================================================
# CLI
# ============================================================

def parse_args() -> argparse.Namespace:

    parser = argparse.ArgumentParser(
        description=(
            "Ricerca gratuita sul web e analisi tramite Gemini "
            "di farmacisti ospedalieri."
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
        help="File output. Default: <input>_risultati.xlsx",
    )

    parser.add_argument(
        "--sheet",
        default="Dati",
        help="Foglio da elaborare (default: Dati)",
    )

    parser.add_argument(
        "--model",
        default=os.getenv(
            "GEMINI_MODEL",
            DEFAULT_MODEL,
        ),
        help=f"Modello Gemini (default: {DEFAULT_MODEL})",
    )

    parser.add_argument(
        "--limit",
        "--max-rows",
        dest="max_rows",
        type=int,
        default=None,
        help="Numero massimo di persone da elaborare",
    )

    parser.add_argument(
        "--start-row",
        type=int,
        default=2,
        help="Prima riga Excel da considerare",
    )

    parser.add_argument(
        "--max-retries",
        type=int,
        default=3,
        help="Retry Gemini per persona",
    )

    parser.add_argument(
        "--delay",
        type=float,
        default=2.0,
        help="Pausa tra persone",
    )

    parser.add_argument(
        "--retry-errors",
        action="store_true",
        help="Riprova righe con Stato ricerca=ERRORE",
    )

    parser.add_argument(
        "--search-results",
        type=int,
        default=DEFAULT_SEARCH_RESULTS,
        help="Risultati massimi raccolti dal motore di ricerca",
    )

    parser.add_argument(
        "--max-pages",
        type=int,
        default=DEFAULT_MAX_PAGES,
        help="Numero massimo di pagine da scaricare per persona",
    )

    parser.add_argument(
        "--log-dir",
        type=Path,
        default=LOG_DIR,
        help="Cartella log",
    )

    args = parser.parse_args()

    if args.max_rows is not None and args.max_rows <= 0:
        parser.error("--limit deve essere maggiore di 0")

    if args.start_row < 2:
        parser.error("--start-row deve essere almeno 2")

    if args.max_retries <= 0:
        parser.error("--max-retries deve essere maggiore di 0")

    if args.delay < 0:
        parser.error("--delay non può essere negativo")

    if args.search_results <= 0:
        parser.error("--search-results deve essere maggiore di 0")

    if args.max_pages <= 0:
        parser.error("--max-pages deve essere maggiore di 0")

    return args


# ============================================================
# LOG
# ============================================================

def configure_logging(log_dir: Path) -> Path:

    log_dir = log_dir.expanduser().resolve()
    log_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    log_file = log_dir / f"run_{timestamp}.log"

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = logging.FileHandler(
        log_file,
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    logging.info("Log: %s", log_file)

    return log_file


# ============================================================
# UTILITY
# ============================================================

def clean(value: Any) -> str:

    if value is None:
        return ""

    return str(value).strip()


def clean_text(text: str) -> str:

    text = re.sub(
        r"\s+",
        " ",
        text or "",
    )

    return text.strip()


def format_birth_date(value: Any) -> str:

    if isinstance(
        value,
        (datetime, date),
    ):
        return value.strftime("%d/%m/%Y")

    return clean(value)


def utc_now() -> str:

    return datetime.now(
        timezone.utc
    ).isoformat(
        timespec="seconds"
    )


# ============================================================
# EXCEL
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
        column
        for column in INPUT_COLUMNS
        if column.casefold() not in headers
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


def output_path_for(
    input_path: Path,
    requested: Path | None,
) -> Path:

    if requested:
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
            f"File input non trovato: {input_path}"
        )

    if input_path.suffix.lower() != ".xlsx":

        raise ValueError(
            "Il file deve essere .xlsx"
        )

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    if output_path.exists():

        logging.info(
            "Riprendo output esistente: %s",
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
        "Creata copia: %s",
        output_path,
    )


def atomic_save(
    workbook,
    output_path: Path,
) -> None:

    fd, temp_name = tempfile.mkstemp(
        prefix=f".{output_path.stem}_",
        suffix=".xlsx",
        dir=output_path.parent,
    )

    os.close(fd)

    temp_path = Path(temp_name)

    try:

        workbook.save(temp_path)

        os.replace(
            temp_path,
            output_path,
        )

    finally:

        temp_path.unlink(
            missing_ok=True
        )


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
# QUERY DI RICERCA
# ============================================================

def build_search_queries(
    person: Person
) -> list[str]:

    full_name = (
        f'"{person.name} {person.surname}"'
    )

    queries = [
        f'{full_name} "farmacista ospedaliero"',
        f'{full_name} "farmacia ospedaliera"',
        f'{full_name} farmacista ospedale',
        f'{full_name} farmacista ASL',
        f'{full_name} farmacista ASST',
        f'{full_name} dirigente farmacista',
    ]

    return queries


# ============================================================
# RICERCA WEB GRATUITA
# ============================================================

def search_web(
    person: Person,
    max_results: int,
) -> list[SearchResult]:

    queries = build_search_queries(
        person
    )

    unique: dict[str, SearchResult] = {}

    # Distribuiamo il numero massimo di risultati
    # fra diverse query.
    per_query = max(
        2,
        min(
            5,
            max_results,
        ),
    )

    for query in queries:

        if len(unique) >= max_results:
            break

        logging.info(
            "Web search: %s",
            query,
        )

        try:

            results = DDGS().text(
                query,
                region="it-it",
                safesearch="moderate",
                max_results=per_query,
            )

            for item in results or []:

                url = clean(
                    item.get("href")
                    or item.get("url")
                )

                if not url:
                    continue

                if url in unique:
                    continue

                unique[url] = SearchResult(
                    title=clean(
                        item.get("title")
                    ),
                    url=url,
                    snippet=clean(
                        item.get("body")
                        or item.get("snippet")
                    ),
                )

                if len(unique) >= max_results:
                    break

        except Exception as exc:

            logging.warning(
                "Ricerca fallita per query '%s': %s",
                query,
                exc,
            )

        # Evita troppe richieste ravvicinate
        time.sleep(
            random.uniform(
                0.4,
                0.9,
            )
        )

    results = list(
        unique.values()
    )

    logging.info(
        "Risultati web unici trovati: %s",
        len(results),
    )

    return results


# ============================================================
# SICUREZZA URL
# ============================================================

def is_public_http_url(
    url: str
) -> bool:

    try:

        parsed = urlparse(url)

        if parsed.scheme not in {
            "http",
            "https",
        }:
            return False

        hostname = parsed.hostname

        if not hostname:
            return False

        # Protezione basilare contro URL locali.
        try:

            addresses = socket.getaddrinfo(
                hostname,
                None,
            )

            for address in addresses:

                ip = ipaddress.ip_address(
                    address[4][0]
                )

                if (
                    ip.is_private
                    or ip.is_loopback
                    or ip.is_link_local
                    or ip.is_reserved
                ):
                    return False

        except socket.gaierror:
            return False

        return True

    except Exception:
        return False


# ============================================================
# DOWNLOAD PAGINE
# ============================================================

def extract_html_text(
    content: bytes
) -> str:

    soup = BeautifulSoup(
        content,
        "html.parser",
    )

    # Elimina parti inutili.
    for tag in soup(
        [
            "script",
            "style",
            "noscript",
            "svg",
            "nav",
            "footer",
        ]
    ):
        tag.decompose()

    text = soup.get_text(
        " ",
        strip=True,
    )

    return clean_text(text)


def extract_pdf_text(
    content: bytes
) -> str:

    try:

        reader = PdfReader(
            io.BytesIO(content)
        )

        parts: list[str] = []

        # Per evitare PDF enormi prendiamo le prime 15 pagine.
        for page in reader.pages[:15]:

            text = page.extract_text()

            if text:
                parts.append(text)

        return clean_text(
            "\n".join(parts)
        )

    except Exception as exc:

        logging.warning(
            "PDF non leggibile: %s",
            exc,
        )

        return ""


def fetch_page_text(
    session: requests.Session,
    url: str,
) -> str:

    if not is_public_http_url(url):

        logging.warning(
            "URL scartato: %s",
            url,
        )

        return ""

    try:

        response = session.get(
            url,
            timeout=HTTP_TIMEOUT,
            allow_redirects=True,
        )

        response.raise_for_status()

        # Limite dimensione ~8 MB
        content = response.content[
            :8 * 1024 * 1024
        ]

        content_type = (
            response.headers
            .get(
                "Content-Type",
                "",
            )
            .lower()
        )

        if (
            "application/pdf"
            in content_type
            or url.lower().endswith(".pdf")
        ):

            return extract_pdf_text(
                content
            )

        if (
            "text/html"
            in content_type
            or not content_type
        ):

            return extract_html_text(
                content
            )

        return ""

    except requests.RequestException as exc:

        logging.warning(
            "Download fallito %s | %s",
            url,
            exc,
        )

        return ""

    except Exception as exc:

        logging.warning(
            "Errore lettura %s | %s",
            url,
            exc,
        )

        return ""


# ============================================================
# SCARICA RISULTATI
# ============================================================

def enrich_search_results(
    results: list[SearchResult],
    max_pages: int,
) -> list[SearchResult]:

    session = requests.Session()

    session.headers.update(
        {
            "User-Agent": USER_AGENT,
            "Accept-Language": "it-IT,it;q=0.9,en;q=0.6",
        }
    )

    pages_read = 0

    for result in results:

        if pages_read >= max_pages:
            break

        logging.info(
            "Leggo: %s",
            result.url,
        )

        text = fetch_page_text(
            session,
            result.url,
        )

        if text:

            result.content = text[
                :DEFAULT_MAX_CHARS_PER_PAGE
            ]

            pages_read += 1

            logging.info(
                "Pagina acquisita: %s caratteri",
                len(result.content),
            )

        else:

            logging.info(
                "Uso solo snippet per: %s",
                result.url,
            )

        time.sleep(
            random.uniform(
                0.3,
                0.7,
            )
        )

    logging.info(
        "Pagine lette: %s",
        pages_read,
    )

    return results


# ============================================================
# COSTRUZIONE CONTESTO
# ============================================================

def build_sources_context(
    results: list[SearchResult],
) -> tuple[str, list[str]]:

    blocks: list[str] = []
    urls: list[str] = []

    total_chars = 0

    for index, result in enumerate(
        results,
        start=1,
    ):

        text = (
            result.content
            if result.content
            else result.snippet
        )

        text = clean_text(text)

        if not text:
            continue

        block = f"""
[FONTE {index}]
Titolo: {result.title}
URL: {result.url}

Contenuto:
{text}
""".strip()

        if (
            total_chars
            + len(block)
            > DEFAULT_MAX_CONTEXT_CHARS
        ):
            break

        blocks.append(block)
        urls.append(result.url)

        total_chars += len(block)

    return (
        "\n\n".join(blocks),
        urls,
    )


# ============================================================
# PROMPT GEMINI
# ============================================================

def build_analysis_prompt(
    person: Person,
    sources_context: str,
) -> str:

    return f"""
Sei un analista di informazioni professionali.

NON hai accesso al web.

Devi utilizzare ESCLUSIVAMENTE le fonti che Python ha già
raccolto e che trovi in fondo a questo messaggio.

PERSONA

Nome: {person.name}
Cognome: {person.surname}
Data di nascita: {person.birth_date or "non disponibile"}
Professione attesa: {DEFAULT_ROLE}

OBIETTIVO

Determina se le fonti identificano ragionevolmente questa
persona come farmacista ospedaliero/a e trova:

- struttura sanitaria;
- reparto / UO / farmacia ospedaliera;
- città;
- ruolo professionale;
- email professionale pubblicamente pubblicata;
- telefono professionale pubblicamente pubblicato.

REGOLE

1. Non utilizzare conoscenze esterne alle fonti fornite.

2. Non inventare dati.

3. Non dedurre email da nome, cognome o dominio.

4. Un'email può essere restituita SOLO se compare
letteralmente nelle fonti.

5. Un telefono può essere restituito SOLO se compare
letteralmente nelle fonti.

6. Sono consentiti:
   - recapiti professionali nominativi;
   - farmacia ospedaliera;
   - reparto/UO;
   - centralino o struttura sanitaria.

7. Non restituire:
   - email personali;
   - telefoni personali;
   - indirizzi di abitazione.

8. La data di nascita serve solo per disambiguare omonimi.

9. found=true solo se le fonti collegano ragionevolmente
la persona alla professione o alla struttura sanitaria.

10. source_indexes deve contenere gli indici delle sole fonti
che supportano concretamente il risultato.

CONFIDENZA

alta:
identificazione molto solida e supportata da fonte
istituzionale o più fonti coerenti.

media:
corrispondenza probabile ma non completamente definitiva.

bassa:
possibile corrispondenza con dubbi significativi.

nessuna:
nessun collegamento sufficientemente affidabile.

========================
FONTI RACCOLTE DA PYTHON
========================

{sources_context}
""".strip()


# ============================================================
# GEMINI - SOLO ANALISI
# ============================================================

def analyze_with_gemini(
    client: genai.Client,
    model: str,
    person: Person,
    search_results: list[SearchResult],
) -> dict[str, Any]:

    context, _ = build_sources_context(
        search_results
    )

    if not context:

        return {
            "found": False,
            "facility": "",
            "department": "",
            "city": "",
            "professional_email": "",
            "professional_phone": "",
            "role_found": "",
            "confidence": "nessuna",
            "notes": (
                "Nessuna fonte web utilizzabile trovata."
            ),
            "source_indexes": [],
            "sources": [],
        }

    interaction = client.interactions.create(
        model=model,
        input=build_analysis_prompt(
            person,
            context,
        ),
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
            "Gemini non ha restituito JSON valido. "
            f"Output: {output_text[:1000]}"
        ) from exc

    if not isinstance(
        result,
        dict,
    ):
        raise ValueError(
            "Risposta Gemini non valida."
        )

    # --------------------------------------------------------
    # TRADUCE GLI INDICI DELLE FONTI IN URL
    # --------------------------------------------------------

    source_indexes = (
        result.get(
            "source_indexes"
        )
        or []
    )

    selected_urls: list[str] = []

    for index in source_indexes:

        try:

            source_index = int(index) - 1

            if (
                0 <= source_index
                < len(search_results)
            ):

                url = search_results[
                    source_index
                ].url

                if url not in selected_urls:
                    selected_urls.append(url)

        except (TypeError, ValueError):
            continue

    result["sources"] = selected_urls

    return result


# ============================================================
# PIPELINE COMPLETA PERSONA
# ============================================================

def research_person(
    client: genai.Client,
    model: str,
    person: Person,
    search_results_limit: int,
    max_pages: int,
) -> dict[str, Any]:

    logging.info(
        "FASE 1 | ricerca web gratuita"
    )

    results = search_web(
        person,
        search_results_limit,
    )

    if not results:

        return {
            "found": False,
            "facility": "",
            "department": "",
            "city": "",
            "professional_email": "",
            "professional_phone": "",
            "role_found": "",
            "confidence": "nessuna",
            "notes": "Nessun risultato trovato sul web.",
            "source_indexes": [],
            "sources": [],
        }

    logging.info(
        "FASE 2 | acquisizione pagine"
    )

    results = enrich_search_results(
        results,
        max_pages,
    )

    logging.info(
        "FASE 3 | analisi Gemini"
    )

    result = analyze_with_gemini(
        client,
        model,
        person,
        results,
    )

    return result


# ============================================================
# RETRY GEMINI / PIPELINE
# ============================================================

def research_with_retry(
    client: genai.Client,
    model: str,
    person: Person,
    max_retries: int,
    search_results_limit: int,
    max_pages: int,
) -> tuple[dict[str, Any], int]:

    last_error: Exception | None = None

    for attempt in range(
        1,
        max_retries + 1,
    ):

        try:

            logging.info(
                "Tentativo %s/%s",
                attempt,
                max_retries,
            )

            result = research_person(
                client=client,
                model=model,
                person=person,
                search_results_limit=search_results_limit,
                max_pages=max_pages,
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

            logging.exception(
                "Tentativo %s/%s fallito per riga %s",
                attempt,
                max_retries,
                person.row,
            )

            if attempt >= max_retries:
                break

            wait = min(
                30.0,
                (
                    2 ** (attempt - 1)
                )
                + random.uniform(
                    0.0,
                    1.0,
                ),
            )

            logging.info(
                "Retry tra %.1f secondi",
                wait,
            )

            time.sleep(wait)

    raise RuntimeError(
        str(last_error)
        if last_error
        else "Errore sconosciuto"
    )


# ============================================================
# SCRITTURA RISULTATI
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

    values = {
        "Stato ricerca":
            status,

        "Struttura":
            clean(
                result.get("facility")
            ),

        "Reparto/UO":
            clean(
                result.get("department")
            ),

        "Citta":
            clean(
                result.get("city")
            ),

        "Email professionale":
            clean(
                result.get("professional_email")
            ),

        "Telefono professionale":
            clean(
                result.get("professional_phone")
            ),

        "Ruolo trovato":
            clean(
                result.get("role_found")
            ),

        "Confidenza":
            clean(
                result.get("confidence")
            ),

        "Note":
            clean(
                result.get("notes")
            ),

        "Fonti":
            "\n".join(
                result.get("sources")
                or []
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
        clean(error)[:5000],
    )


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    load_dotenv()

    args = parse_args()

    log_file = configure_logging(
        args.log_dir
    )

    api_key = os.getenv(
        "GEMINI_API_KEY"
    )

    if not api_key:

        raise RuntimeError(
            "GEMINI_API_KEY non trovata nel file .env"
        )

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

    workbook = load_workbook(
        output_path
    )

    if args.sheet not in workbook.sheetnames:

        raise ValueError(
            f"Foglio '{args.sheet}' non trovato. "
            f"Disponibili: {workbook.sheetnames}"
        )

    ws = workbook[
        args.sheet
    ]

    headers = normalized_headers(
        ws
    )

    require_input_columns(
        headers
    )

    headers = ensure_result_columns(
        ws
    )

    atomic_save(
        workbook,
        output_path,
    )

    client = genai.Client(
        api_key=api_key
    )

    processed = 0
    successful = 0
    errors = 0
    skipped = 0

    logging.info(
        "============================================"
    )

    logging.info(
        "RICERCA FARMACISTI - MODALITA GRATUITA"
    )

    logging.info(
        "Input: %s",
        input_path,
    )

    logging.info(
        "Output: %s",
        output_path,
    )

    logging.info(
        "Log: %s",
        log_file,
    )

    logging.info(
        "Modello Gemini: %s",
        args.model,
    )

    logging.info(
        "Ricerca web: DDGS"
    )

    logging.info(
        "Risultati ricerca/persona: %s",
        args.search_results,
    )

    logging.info(
        "Pagine lette/persona: %s",
        args.max_pages,
    )

    logging.info(
        "============================================"
    )

    start_row = max(
        2,
        args.start_row,
    )

    for row in range(
        start_row,
        ws.max_row + 1,
    ):

        if (
            args.max_rows is not None
            and processed >= args.max_rows
        ):
            break

        status = clean(
            ws.cell(
                row,
                headers[
                    "stato ricerca"
                ],
            ).value
        ).upper()

        if status in TERMINAL_STATUSES:

            skipped += 1
            continue

        if (
            status == "ERRORE"
            and not args.retry_errors
        ):

            skipped += 1
            continue

        person = person_from_row(
            ws,
            row,
            headers,
        )

        if (
            not person.name
            and not person.surname
        ):
            continue

        logging.info(
            "--------------------------------------------"
        )

        logging.info(
            "[%s] Riga %s | %s %s | nascita: %s",
            processed + 1,
            row,
            person.name,
            person.surname,
            person.birth_date or "N/D",
        )

        try:

            result, attempts = research_with_retry(
                client=client,
                model=args.model,
                person=person,
                max_retries=args.max_retries,
                search_results_limit=args.search_results,
                max_pages=args.max_pages,
            )

            write_result(
                ws,
                row,
                headers,
                result,
                attempts,
            )

            successful += 1

            logging.info(
                "RISULTATO: %s",
                (
                    "TROVATO"
                    if result.get("found")
                    else "NESSUN RISULTATO"
                ),
            )

            logging.info(
                "Struttura: %s",
                clean(
                    result.get("facility")
                )
                or "N/D",
            )

            logging.info(
                "Reparto/UO: %s",
                clean(
                    result.get("department")
                )
                or "N/D",
            )

            logging.info(
                "Email: %s",
                clean(
                    result.get("professional_email")
                )
                or "N/D",
            )

            logging.info(
                "Telefono: %s",
                clean(
                    result.get("professional_phone")
                )
                or "N/D",
            )

            logging.info(
                "Confidenza: %s",
                clean(
                    result.get("confidence")
                )
                or "N/D",
            )

            logging.info(
                "Fonti usate: %s",
                len(
                    result.get("sources")
                    or []
                ),
            )

        except (
            KeyboardInterrupt,
            SystemExit,
        ):

            logging.warning(
                "Interruzione manuale. "
                "Salvataggio checkpoint."
            )

            atomic_save(
                workbook,
                output_path,
            )

            raise

        except Exception as exc:

            errors += 1

            write_error(
                ws,
                row,
                headers,
                exc,
                args.max_retries,
            )

            logging.exception(
                "Riga %s non completata",
                row,
            )

        atomic_save(
            workbook,
            output_path,
        )

        processed += 1

        logging.info(
            "Checkpoint salvato."
        )

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

    logging.info(
        "============================================"
    )

    logging.info(
        "ESECUZIONE TERMINATA"
    )

    logging.info(
        "Persone elaborate: %s",
        processed,
    )

    logging.info(
        "Ricerche concluse: %s",
        successful,
    )

    logging.info(
        "Errori: %s",
        errors,
    )

    logging.info(
        "Righe saltate: %s",
        skipped,
    )

    logging.info(
        "Output: %s",
        output_path,
    )

    logging.info(
        "Log: %s",
        log_file,
    )

    logging.info(
        "============================================"
    )

    return 0


if __name__ == "__main__":

    try:

        raise SystemExit(
            main()
        )

    except KeyboardInterrupt:

        print(
            "\nEsecuzione interrotta."
        )

        raise SystemExit(130)
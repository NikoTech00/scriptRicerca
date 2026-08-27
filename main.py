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
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from ddgs import DDGS
from dotenv import load_dotenv
from google import genai
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from pypdf import PdfReader


# ============================================================
# VERSIONE / CONFIGURAZIONE
# ============================================================

VERSION = "V6.4"

INPUT_COLUMNS = (
    "Pers_ID",
    "Cognome",
    "Nome",
    "Data di Nascita",
)

RESULT_COLUMNS = (
    "Stato ricerca",
    "Indice ricerca",
    "Termine ricerca usato",
    "Struttura",
    "Reparto/UO",
    "Citta",
    "Email professionale",
    "Telefono professionale",
    "Ruolo trovato",
    "Confidenza",
    "Note",
    "Fonti",
    "CV salvato",
    "Tentativi",
    "Ultimo aggiornamento UTC",
    "Errore",
)

DEFAULT_ROLE = "farmacista ospedaliero/a"
DEFAULT_MODEL = "gemini-3.6-flash"

DEFAULT_SEARCH_RESULTS = 8
DEFAULT_RESULTS_TO_GEMINI = 12

HTTP_TIMEOUT = 15
MAX_CV_BYTES = 15 * 1024 * 1024
MAX_PAGE_BYTES = 10 * 1024 * 1024
MAX_PAGE_CHARS = 18000
MAX_CONTEXT_CHARS = 32000

SEARCH_BACKEND = os.getenv("SEARCH_BACKEND", "auto")

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/151.0 Safari/537.36"
)

LOG_DIR = Path("logs")
CV_DIR = Path("cv")

ITALIAN_MONTHS = {
    1: "gennaio",
    2: "febbraio",
    3: "marzo",
    4: "aprile",
    5: "maggio",
    6: "giugno",
    7: "luglio",
    8: "agosto",
    9: "settembre",
    10: "ottobre",
    11: "novembre",
    12: "dicembre",
}

CV_POSITIVE_TERMS = (
    "curriculum vitae",
    "curriculum professionale",
    "curriculum formativo",
    "curriculum scientifico",
    "curriculum personale",
    "cv europass",
    "europass",
)

CV_STRUCTURE_TERMS = (
    "curriculum vitae",
    "europass",
    "esperienza professionale",
    "esperienze professionali",
    "esperienza lavorativa",
    "esperienze lavorative",
    "istruzione e formazione",
    "formazione professionale",
    "titoli di studio",
    "incarichi professionali",
    "capacita e competenze",
    "capacità e competenze",
    "posizione ricoperta",
    "principali attività",
    "principali attivita",
    "principali mansioni",
)

CV_NEGATIVE_TERMS = (
    "graduatoria",
    "graduatorie",
    "concorso",
    "concorsi",
    "prova scritta",
    "prova pratica",
    "prova orale",
    "esito prova",
    "esiti",
    "candidati ammessi",
    "elenco candidati",
    "commissione esaminatrice",
    "verbale della commissione",
    "determinazione",
    "deliberazione",
    "delibera",
    "avviso pubblico",
    "bando",
    "manifestazione di interesse",
    "tracce",
    "quiz",
    "curriculum scolastico",
    "curriculum dello studente",
    "curricolo",
)

PROFESSIONAL_TERMS = (
    "farmacista ospedaliero",
    "dirigente farmacista",
    "farmacia ospedaliera",
    "farmacista",
    "farmaceutico",
    "farmaceutica",
)

HEALTHCARE_TERMS = (
    "ospedale",
    "ospedaliero",
    "azienda ospedaliera",
    "azienda sanitaria",
    "asl",
    "ausl",
    "asst",
    "aou",
    "irccs",
    "policlinico",
    "farmacia ospedaliera",
    "servizio farmaceutico",
)

EMAIL_RE = re.compile(
    r"\b[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}\b",
    flags=re.IGNORECASE,
)

PHONE_RE = re.compile(
    r"(?<!\d)(?:\+39[\s.\-]?)?(?:0\d{1,3}[\s.\-]?)?\d{6,10}(?!\d)"
)


# ============================================================
# MODELLI DATI
# ============================================================

@dataclass(frozen=True)
class Person:
    row: int
    pers_id: str
    surname: str
    name: str
    birth_date: str


@dataclass(frozen=True)
class SearchResult:
    title: str
    url: str
    snippet: str
    query_index: int
    query_text: str
    backend: str
    category: str = "ordinary"


@dataclass
class CollectedEvidence:
    results: list[SearchResult]
    primary: SearchResult | None
    primary_text: str
    cv_source: SearchResult | None
    cv_path: Path | None
    cv_text: str


class GeminiDeferredError(RuntimeError):
    """Errore Gemini temporaneo: i dati web raccolti non vanno persi."""


# ============================================================
# SCHEMA GEMINI
# ============================================================

RESULT_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "found": {"type": "boolean"},
        "facility": {"type": "string"},
        "department": {"type": "string"},
        "city": {"type": "string"},
        "professional_email": {"type": "string"},
        "professional_phone": {"type": "string"},
        "role_found": {"type": "string"},
        "confidence": {
            "type": "string",
            "enum": ["alta", "media", "bassa", "nessuna"],
        },
        "notes": {"type": "string"},
        "supporting_indexes": {
            "type": "array",
            "items": {"type": "integer"},
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
        "supporting_indexes",
    ],
}


# ============================================================
# CLI
# ============================================================

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=f"Ricerca farmacisti ospedalieri - {VERSION}"
    )

    parser.add_argument("input", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--sheet", default="Dati")
    parser.add_argument(
        "--model",
        default=os.getenv("GEMINI_MODEL", DEFAULT_MODEL),
    )
    parser.add_argument(
        "--limit",
        "--max-rows",
        dest="max_rows",
        type=int,
        default=None,
    )
    parser.add_argument("--start-row", type=int, default=2)
    parser.add_argument(
        "--max-retries",
        type=int,
        default=3,
        help="Retry solo per la fase Gemini, non per le ricerche web.",
    )
    parser.add_argument("--delay", type=float, default=2.0)
    parser.add_argument("--retry-errors", action="store_true")
    parser.add_argument("--retry-no-results", action="store_true")
    parser.add_argument(
        "--search-results",
        type=int,
        default=DEFAULT_SEARCH_RESULTS,
    )
    parser.add_argument(
        "--results-to-gemini",
        type=int,
        default=DEFAULT_RESULTS_TO_GEMINI,
    )
    parser.add_argument("--log-dir", type=Path, default=LOG_DIR)
    parser.add_argument("--cv-dir", type=Path, default=CV_DIR)
    parser.add_argument(
        "--search-backend",
        default=SEARCH_BACKEND,
        help="Backend DDGS. Default: env SEARCH_BACKEND oppure auto.",
    )

    args = parser.parse_args()

    if args.max_rows is not None and args.max_rows <= 0:
        parser.error("--limit deve essere > 0")
    if args.start_row < 2:
        parser.error("--start-row deve essere >= 2")
    if args.max_retries <= 0:
        parser.error("--max-retries deve essere > 0")
    if args.delay < 0:
        parser.error("--delay non può essere negativo")
    if args.search_results <= 0:
        parser.error("--search-results deve essere > 0")
    if args.results_to_gemini <= 0:
        parser.error("--results-to-gemini deve essere > 0")

    return args


# ============================================================
# LOG
# ============================================================

def configure_logging(log_dir: Path) -> Path:
    log_dir = log_dir.expanduser().resolve()
    log_dir.mkdir(parents=True, exist_ok=True)

    log_file = (
        log_dir
        / f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    )

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    console_handler = logging.StreamHandler()

    file_handler.setFormatter(formatter)
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

    text = str(value).strip()

    if text.casefold() in {
        "null",
        "none",
        "nan",
        "n/a",
        "nd",
        "n.d.",
        "non disponibile",
    }:
        return ""

    return text


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def normalize_for_match(value: str) -> str:
    value = (
        value.casefold()
        .replace("-", " ")
        .replace("_", " ")
        .replace("'", " ")
        .replace("’", " ")
    )
    return re.sub(r"\s+", " ", value).strip()


def format_birth_date(value: Any) -> str:
    if isinstance(value, (date, datetime)):
        return value.strftime("%d/%m/%Y")
    return clean(value)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def safe_filename_part(value: str) -> str:
    value = clean(value)
    value = re.sub(r'[<>:"/\\|?*]', "_", value)
    value = re.sub(r"\s+", "_", value)
    return value.strip("._ ") or "ND"


def canonical_url(url: str) -> str:
    return url.split("#")[0].rstrip("/").casefold()


def normalize_date_string(value: str) -> str:
    value = clean(value).replace("-", "/")
    try:
        return datetime.strptime(value, "%d/%m/%Y").strftime("%d/%m/%Y")
    except ValueError:
        return value


def extract_first_email(text: str) -> str:
    match = EMAIL_RE.search(text or "")
    return clean(match.group(0)) if match else ""


def extract_first_phone(text: str) -> str:
    for match in PHONE_RE.finditer(text or ""):
        value = clean(match.group(0))
        digits = re.sub(r"\D", "", value)
        if 7 <= len(digits) <= 13:
            return value
    return ""


# ============================================================
# EXCEL
# ============================================================

def normalized_headers(ws) -> dict[str, int]:
    result: dict[str, int] = {}

    for cell in ws[1]:
        label = clean(cell.value)
        if label:
            result[label.casefold()] = cell.column

    return result


def require_input_columns(headers: dict[str, int]) -> None:
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


def validate_pers_ids(ws, headers: dict[str, int]) -> None:
    missing_rows: list[int] = []
    seen: dict[str, int] = {}
    duplicates: list[tuple[str, int, int]] = []

    pers_col = headers["pers_id"]
    surname_col = headers["cognome"]
    name_col = headers["nome"]

    for row in range(2, ws.max_row + 1):
        surname = clean(ws.cell(row, surname_col).value)
        name = clean(ws.cell(row, name_col).value)

        if not surname and not name:
            continue

        pers_id = clean(ws.cell(row, pers_col).value)

        if not pers_id:
            missing_rows.append(row)
            continue

        if pers_id in seen:
            duplicates.append((pers_id, seen[pers_id], row))
        else:
            seen[pers_id] = row

    if missing_rows:
        raise ValueError(
            "Pers_ID mancante nelle righe: "
            + ", ".join(str(row) for row in missing_rows[:30])
        )

    for pers_id, first_row, second_row in duplicates:
        logging.warning(
            "Pers_ID duplicato '%s' alle righe %s e %s",
            pers_id,
            first_row,
            second_row,
        )


def ensure_result_columns(ws) -> dict[str, int]:
    headers = normalized_headers(ws)

    for label in RESULT_COLUMNS:
        key = label.casefold()

        if key in headers:
            continue

        col = ws.max_column + 1
        cell = ws.cell(1, col, label)

        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
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
        raise ValueError("Il file input deve essere .xlsx")

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

    if input_path.resolve() != output_path.resolve():
        shutil.copy2(input_path, output_path)

    logging.info(
        "Creata copia di lavoro: %s",
        output_path,
    )


def atomic_save(workbook, output_path: Path) -> None:
    fd, temp_name = tempfile.mkstemp(
        prefix=f".{output_path.stem}_",
        suffix=".xlsx",
        dir=output_path.parent,
    )
    os.close(fd)
    temp_path = Path(temp_name)

    try:
        workbook.save(temp_path)
        os.replace(temp_path, output_path)
    finally:
        temp_path.unlink(missing_ok=True)


def person_from_row(
    ws,
    row: int,
    headers: dict[str, int],
) -> Person:
    return Person(
        row=row,
        pers_id=clean(
            ws.cell(row, headers["pers_id"]).value
        ),
        surname=clean(
            ws.cell(row, headers["cognome"]).value
        ),
        name=clean(
            ws.cell(row, headers["nome"]).value
        ),
        birth_date=format_birth_date(
            ws.cell(
                row,
                headers["data di nascita"],
            ).value
        ),
    )


# ============================================================
# QUERY
# ============================================================

def birth_date_variants(person: Person) -> list[str]:
    value = clean(person.birth_date)

    if not value:
        return []

    try:
        parsed = datetime.strptime(value, "%d/%m/%Y")
    except ValueError:
        return [value]

    variants = [
        parsed.strftime("%d/%m/%Y"),
        parsed.strftime("%Y-%m-%d"),
        f"{parsed.day} {ITALIAN_MONTHS[parsed.month]} {parsed.year}",
    ]

    return list(dict.fromkeys(variants))


def build_search_stages(
    person: Person,
) -> list[tuple[int, str, list[str]]]:
    full_name = f'"{person.name} {person.surname}"'
    dates = birth_date_variants(person)

    stage_1 = [
        f'"{birth}" {full_name} "farmacista ospedaliero"'
        for birth in dates
    ] or [
        f'{full_name} "farmacista ospedaliero"'
    ]

    stage_2 = [
        f'"{birth}" {full_name} farmacista'
        for birth in dates
    ] or [
        f"{full_name} farmacista"
    ]

    stage_3 = [
        f'{full_name} "farmacista ospedaliero"',
        f'{full_name} "dirigente farmacista"',
    ]

    stage_4 = [
        f"{full_name} farmacista",
        f'{full_name} "farmacia ospedaliera"',
    ]

    return [
        (1, "DATA + NOME COGNOME + FARMACISTA OSPEDALIERO", stage_1),
        (2, "DATA + NOME COGNOME + FARMACISTA", stage_2),
        (3, "NOME COGNOME + FARMACISTA OSPEDALIERO", stage_3),
        (4, "NOME COGNOME + FARMACISTA", stage_4),
    ]


def build_linkedin_queries(person: Person) -> list[str]:
    full_name = f'"{person.name} {person.surname}"'

    return [
        f"site:linkedin.com/in {full_name} farmacista",
        f"site:linkedin.com/in {full_name} ospedale",
    ]


def build_cv_queries(person: Person) -> list[str]:
    full_name = f'"{person.name} {person.surname}"'

    return [
        f'{full_name} "curriculum vitae" farmacista',
        f'{full_name} "curriculum vitae" filetype:pdf',
        f'{full_name} "curriculum" "dirigente farmacista"',
        f'{full_name} "curriculum vitae" ospedale',
        f'{full_name} "europass" farmacista',
    ]


# ============================================================
# CLASSIFICAZIONE FONTI
# ============================================================

def get_domain(url: str) -> str:
    try:
        return (urlparse(url).hostname or "").casefold()
    except Exception:
        return ""


def is_linkedin_source(url: str) -> bool:
    return "linkedin.com" in get_domain(url)


def is_weak_source(url: str) -> bool:
    domain = get_domain(url)
    weak_domains = (
        "facebook.com",
        "instagram.com",
        "indeed.com",
        "quora.com",
        "pinterest.",
        "tiktok.com",
    )
    return any(value in domain for value in weak_domains)


def is_institutional_source(url: str) -> bool:
    domain = get_domain(url)
    tokens = (
        "asl",
        "ausl",
        "asst",
        "ats",
        "osped",
        "policlin",
        "irccs",
        "regione",
        "salute",
        "sanita",
        "gov.it",
        "aou",
        "asur",
        "asp",
    )
    return any(token in domain for token in tokens)


def is_professional_source(url: str) -> bool:
    domain = get_domain(url)
    tokens = (
        "sifoweb",
        "sifo.it",
        "fofi",
        "ordinefarmacisti",
        "ordinefarmacist",
        "farmacista33",
        "universita",
        "unimi",
        "unibo",
        "unipd",
        "unito",
        "unina",
        "uniroma",
    )
    return any(token in domain for token in tokens)


# ============================================================
# SCORE / CV
# ============================================================

def result_identity_score(
    result: SearchResult,
    person: Person,
) -> int:
    text = normalize_for_match(
        f"{result.title} {result.snippet} {result.url}"
    )

    name = normalize_for_match(person.name)
    surname = normalize_for_match(person.surname)
    full_name = normalize_for_match(
        f"{person.name} {person.surname}"
    )
    reverse_name = normalize_for_match(
        f"{person.surname} {person.name}"
    )

    score = 0

    if full_name and full_name in text:
        score += 150
    if reverse_name and reverse_name in text:
        score += 150
    if surname and surname in text:
        score += 45
    if name and name in text:
        score += 30

    return score


def is_pdf_result(result: SearchResult) -> bool:
    text = normalize_for_match(
        f"{result.title} {result.snippet} {result.url}"
    )
    return (
        ".pdf" in result.url.casefold()
        or " pdf " in f" {text} "
    )


def looks_like_cv(result: SearchResult) -> bool:
    text = normalize_for_match(
        f"{result.title} {result.snippet} {result.url}"
    )

    if any(term in text for term in CV_NEGATIVE_TERMS):
        return False

    return any(term in text for term in CV_POSITIVE_TERMS)


def cv_candidate_score(
    result: SearchResult,
    person: Person,
) -> int:
    text = normalize_for_match(
        f"{result.title} {result.snippet} {result.url}"
    )

    score = result_identity_score(result, person)

    for term in CV_POSITIVE_TERMS:
        if term in text:
            score += 40

    if ".pdf" in result.url.casefold():
        score += 25

    if is_institutional_source(result.url):
        score += 20

    for term in CV_NEGATIVE_TERMS:
        if term in text:
            score -= 150

    return score


def should_try_as_cv(
    result: SearchResult,
    person: Person,
) -> bool:
    text = normalize_for_match(
        f"{result.title} {result.snippet} {result.url}"
    )

    if any(term in text for term in CV_NEGATIVE_TERMS):
        return False

    has_cv_term = any(
        term in text
        for term in CV_POSITIVE_TERMS
    )

    return (
        has_cv_term
        and cv_candidate_score(result, person) >= 70
    )


def result_score(
    result: SearchResult,
    person: Person,
) -> int:
    text = normalize_for_match(
        f"{result.title} {result.snippet} {result.url}"
    )

    score = result_identity_score(result, person)

    professional_scores = {
        "farmacista ospedaliero": 90,
        "dirigente farmacista": 80,
        "farmacia ospedaliera": 65,
        "farmacista": 45,
        "farmaceut": 30,
        "ospedal": 30,
        "asl": 25,
        "ausl": 25,
        "asst": 25,
        "irccs": 25,
        "azienda ospedaliera": 30,
        "azienda sanitaria": 25,
        "delibera": 20,
        "trasparenza": 20,
    }

    for keyword, points in professional_scores.items():
        if keyword in text:
            score += points

    if looks_like_cv(result):
        score += 220
    elif is_pdf_result(result):
        score += 50

    if is_institutional_source(result.url):
        score += 110

    if is_professional_source(result.url):
        score += 70

    if is_linkedin_source(result.url):
        score += 55

    if is_weak_source(result.url):
        score -= 60

    bad_words = (
        "come diventare",
        "stipendio",
        "master",
        "corso di laurea",
        "offerta di lavoro",
        "offerte di lavoro",
    )

    for word in bad_words:
        if word in text:
            score -= 90

    return score


# ============================================================
# SEARCH
# ============================================================

def search_query(
    query: str,
    query_index: int,
    category: str,
    max_results: int,
    backend: str,
) -> list[SearchResult]:
    logging.info(
        "SEARCH | categoria=%s | livello=%s | backend=%s | %s",
        category,
        query_index,
        backend,
        query,
    )

    try:
        raw_results = DDGS(
            timeout=HTTP_TIMEOUT
        ).text(
            query,
            region="it-it",
            safesearch="moderate",
            max_results=max_results,
            backend=backend,
        ) or []

    except Exception as exc:
        logging.warning(
            "SEARCH FALLITA | categoria=%s | livello=%s | %s: %s",
            category,
            query_index,
            type(exc).__name__,
            exc,
        )
        return []

    collected: list[SearchResult] = []
    seen: set[str] = set()

    for item in raw_results:
        url = clean(
            item.get("href")
            or item.get("url")
        )

        if not url.startswith(("http://", "https://")):
            continue

        key = canonical_url(url)

        if key in seen:
            continue

        seen.add(key)

        collected.append(
            SearchResult(
                title=clean(item.get("title")),
                url=url,
                snippet=clean(
                    item.get("body")
                    or item.get("snippet")
                ),
                query_index=query_index,
                query_text=query,
                backend=backend,
                category=category,
            )
        )

    logging.info(
        "SEARCH COMPLETATA | categoria=%s | livello=%s | risultati=%s",
        category,
        query_index,
        len(collected),
    )

    return collected


def deduplicate_results(
    items: list[SearchResult],
) -> list[SearchResult]:
    unique: dict[str, SearchResult] = {}

    for result in items:
        key = canonical_url(result.url)
        previous = unique.get(key)

        if previous is None:
            unique[key] = result
        else:
            # Conserva la variante più informativa.
            previous_len = len(previous.title) + len(previous.snippet)
            current_len = len(result.title) + len(result.snippet)

            if current_len > previous_len:
                unique[key] = result

    return list(unique.values())


def search_person(
    person: Person,
    max_results: int,
    backend: str,
) -> list[SearchResult]:
    collected: list[SearchResult] = []

    for stage_index, stage_name, queries in build_search_stages(
        person
    ):
        logging.info("============================================")
        logging.info(
            "LIVELLO %s/4 | %s",
            stage_index,
            stage_name,
        )

        stage_results: list[SearchResult] = []

        for query in queries:
            stage_results.extend(
                search_query(
                    query=query,
                    query_index=stage_index,
                    category="ordinary",
                    max_results=max_results,
                    backend=backend,
                )
            )
            time.sleep(random.uniform(0.2, 0.5))

        stage_results = deduplicate_results(stage_results)

        logging.info(
            "L%s | grezzi conservati=%s",
            stage_index,
            len(stage_results),
        )

        collected.extend(stage_results)

    logging.info("============================================")
    logging.info("RICERCA LINKEDIN DEDICATA")

    for query in build_linkedin_queries(person):
        collected.extend(
            search_query(
                query=query,
                query_index=5,
                category="linkedin",
                max_results=max_results,
                backend=backend,
            )
        )
        time.sleep(random.uniform(0.2, 0.5))

    results = deduplicate_results(collected)

    results.sort(
        key=lambda result: result_score(
            result,
            person,
        ),
        reverse=True,
    )

    logging.info(
        "RISULTATI PERSONA UNICI | %s",
        len(results),
    )

    for index, result in enumerate(
        results[:20],
        start=1,
    ):
        logging.info(
            "PERSONA #%s | score=%s | identity=%s | %s",
            index,
            result_score(result, person),
            result_identity_score(result, person),
            result.url,
        )

    return results


def search_cv(
    person: Person,
    max_results: int,
    backend: str,
) -> list[SearchResult]:
    logging.info("============================================")
    logging.info(
        "RICERCA CV DEDICATA | Pers_ID=%s",
        person.pers_id,
    )

    collected: list[SearchResult] = []

    for query in build_cv_queries(person):
        collected.extend(
            search_query(
                query=query,
                query_index=6,
                category="cv",
                max_results=max_results,
                backend=backend,
            )
        )
        time.sleep(random.uniform(0.2, 0.5))

    results = deduplicate_results(collected)

    results.sort(
        key=lambda result: cv_candidate_score(
            result,
            person,
        ),
        reverse=True,
    )

    logging.info(
        "RISULTATI CV UNICI | %s",
        len(results),
    )

    return results


# ============================================================
# URL SAFETY / HTTP
# ============================================================

def is_public_http_url(url: str) -> bool:
    try:
        parsed = urlparse(url)

        if parsed.scheme not in {"http", "https"}:
            return False

        hostname = parsed.hostname

        if not hostname:
            return False

        addresses = socket.getaddrinfo(
            hostname,
            parsed.port
            or (
                443
                if parsed.scheme == "https"
                else 80
            ),
            type=socket.SOCK_STREAM,
        )

        for address in addresses:
            ip = ipaddress.ip_address(address[4][0])

            if (
                ip.is_private
                or ip.is_loopback
                or ip.is_link_local
                or ip.is_reserved
                or ip.is_multicast
            ):
                return False

        return True

    except (OSError, ValueError):
        return False


def http_get(url: str) -> requests.Response | None:
    if not is_public_http_url(url):
        logging.warning(
            "URL NON PUBBLICO/INVALIDO | %s",
            url,
        )
        return None

    try:
        response = requests.get(
            url,
            headers={
                "User-Agent": USER_AGENT,
                "Accept-Language": "it-IT,it;q=0.9,en;q=0.6",
            },
            timeout=HTTP_TIMEOUT,
            allow_redirects=True,
            stream=True,
        )

        response.raise_for_status()

        if not is_public_http_url(response.url):
            response.close()
            return None

        return response

    except requests.RequestException as exc:
        logging.warning(
            "DOWNLOAD FALLITO | %s | %s",
            url,
            exc,
        )
        return None


# ============================================================
# PDF / HTML
# ============================================================

def read_stream_limited(
    response: requests.Response,
    max_bytes: int,
) -> bytes | None:
    data = bytearray()

    for chunk in response.iter_content(
        chunk_size=65536
    ):
        if not chunk:
            continue

        data.extend(chunk)

        if len(data) > max_bytes:
            return None

    return bytes(data)


def extract_pdf_text(content: bytes) -> str:
    try:
        reader = PdfReader(io.BytesIO(content))
        parts: list[str] = []

        for page in reader.pages[:50]:
            text = page.extract_text()
            if text:
                parts.append(text)

        return clean_text("\n".join(parts))

    except Exception as exc:
        logging.warning(
            "PDF NON LEGGIBILE | %s",
            exc,
        )
        return ""


def extract_html_text(content: bytes) -> str:
    soup = BeautifulSoup(
        content,
        "html.parser",
    )

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

    return clean_text(
        soup.get_text(
            " ",
            strip=True,
        )
    )


def fetch_page_text(url: str) -> str:
    response = http_get(url)

    if response is None:
        return ""

    try:
        raw = read_stream_limited(
            response,
            MAX_PAGE_BYTES,
        )

        if raw is None:
            logging.warning(
                "PAGINA TROPPO GRANDE | %s",
                url,
            )
            return ""

        content_type = (
            response.headers
            .get(
                "Content-Type",
                "",
            )
            .casefold()
        )

        if (
            raw.startswith(b"%PDF")
            or "application/pdf" in content_type
        ):
            text = extract_pdf_text(raw)
        else:
            text = extract_html_text(raw)

        return text[:MAX_PAGE_CHARS]

    finally:
        response.close()


# ============================================================
# CV
# ============================================================

def cv_file_path(
    person: Person,
    cv_dir: Path,
) -> Path:
    filename = (
        f"{safe_filename_part(person.pers_id)}_"
        f"{safe_filename_part(person.surname)}_"
        f"{safe_filename_part(person.name)}.pdf"
    )
    return cv_dir / filename


def pdf_text_matches_person(
    text: str,
    person: Person,
) -> bool:
    normalized = normalize_for_match(text)

    name = normalize_for_match(person.name)
    surname = normalize_for_match(person.surname)
    full_name = normalize_for_match(
        f"{person.name} {person.surname}"
    )
    reverse_name = normalize_for_match(
        f"{person.surname} {person.name}"
    )

    if not name or not surname:
        return False

    identity_ok = (
        full_name in normalized
        or reverse_name in normalized
        or (
            name in normalized
            and surname in normalized
        )
    )

    if not identity_ok:
        return False

    explicit_cv = (
        "curriculum vitae" in normalized
        or "europass" in normalized
    )

    structure_hits = sum(
        1
        for term in CV_STRUCTURE_TERMS
        if normalize_for_match(term) in normalized
    )

    if not explicit_cv and structure_hits < 2:
        return False

    negative_hits = sum(
        1
        for term in CV_NEGATIVE_TERMS
        if normalize_for_match(term) in normalized
    )

    if negative_hits >= 2 and not explicit_cv:
        return False

    if not any(
        normalize_for_match(term) in normalized
        for term in PROFESSIONAL_TERMS
    ):
        return False

    if person.birth_date:
        birth_match = re.search(
            r"(?:data\s+di\s+nascita|"
            r"nato\s+(?:a|il)|"
            r"nata\s+(?:a|il))"
            r".{0,60}?"
            r"(\d{1,2}[/-]\d{1,2}[/-]\d{4})",
            text,
            flags=re.IGNORECASE | re.DOTALL,
        )

        if birth_match:
            found_birth = normalize_date_string(
                birth_match.group(1)
            )
            expected_birth = normalize_date_string(
                person.birth_date
            )

            if found_birth != expected_birth:
                return False

    return True


def try_pdf_url(
    url: str,
    person: Person,
    cv_dir: Path,
) -> tuple[Path, str] | None:
    response = http_get(url)

    if response is None:
        return None

    try:
        raw = read_stream_limited(
            response,
            MAX_CV_BYTES,
        )

        if raw is None:
            logging.warning(
                "PDF TROPPO GRANDE | %s",
                url,
            )
            return None

        content_type = (
            response.headers
            .get(
                "Content-Type",
                "",
            )
            .casefold()
        )

        if not (
            raw.startswith(b"%PDF")
            or "application/pdf" in content_type
        ):
            return None

        text = extract_pdf_text(raw)

        if not text:
            return None

        if not pdf_text_matches_person(
            text,
            person,
        ):
            logging.info(
                "PDF SCARTATO: non è un CV verificato della persona | %s",
                url,
            )
            return None

        cv_dir.mkdir(
            parents=True,
            exist_ok=True,
        )

        destination = cv_file_path(
            person,
            cv_dir,
        )

        destination.write_bytes(raw)

        logging.info(
            "CV SALVATO E VERIFICATO | %s",
            destination,
        )

        return destination, text

    finally:
        response.close()


def download_cv(
    result: SearchResult,
    person: Person,
    cv_dir: Path,
) -> tuple[Path, str] | None:
    logging.info(
        "DOWNLOAD CANDIDATO CV | score=%s | %s",
        cv_candidate_score(result, person),
        result.url,
    )

    direct = try_pdf_url(
        result.url,
        person,
        cv_dir,
    )

    if direct:
        return direct

    response = http_get(result.url)

    if response is None:
        return None

    try:
        raw = read_stream_limited(
            response,
            2 * 1024 * 1024,
        )

        if raw is None:
            return None

        content_type = (
            response.headers
            .get(
                "Content-Type",
                "",
            )
            .casefold()
        )

        if (
            "html" not in content_type
            and not raw.lstrip().startswith((b"<", b"<!"))
        ):
            return None

        soup = BeautifulSoup(
            raw,
            "html.parser",
        )

    finally:
        response.close()

    links: list[tuple[int, str]] = []

    for anchor in soup.select("a[href]"):
        href = clean(anchor.get("href"))

        if not href:
            continue

        candidate_url = urljoin(
            result.url,
            href,
        )

        label = normalize_for_match(
            f"{candidate_url} "
            f"{anchor.get_text(' ', strip=True)}"
        )

        positive = sum(
            1
            for term in CV_POSITIVE_TERMS
            if term in label
        )

        if ".pdf" in candidate_url.casefold():
            positive += 2

        negative = sum(
            1
            for term in CV_NEGATIVE_TERMS
            if term in label
        )

        link_score = (
            positive * 40
            - negative * 100
        )

        if link_score > 0:
            links.append(
                (link_score, candidate_url)
            )

    links.sort(
        key=lambda item: item[0],
        reverse=True,
    )

    seen: set[str] = set()

    for _, candidate_url in links[:15]:
        key = canonical_url(candidate_url)

        if key in seen:
            continue

        seen.add(key)

        found = try_pdf_url(
            candidate_url,
            person,
            cv_dir,
        )

        if found:
            return found

    return None


# ============================================================
# EVIDENZA PYTHON FORTE
# ============================================================

def detect_role(text: str) -> str:
    normalized = normalize_for_match(text)

    role_candidates = (
        "dirigente farmacista",
        "farmacista ospedaliero",
        "farmacista",
    )

    for role in role_candidates:
        if role in normalized:
            return role

    return ""


def strong_result_for_python(
    result: SearchResult,
    person: Person,
) -> bool:
    text = normalize_for_match(
        f"{result.title} {result.snippet}"
    )

    identity = result_identity_score(
        result,
        person,
    )

    professional = any(
        normalize_for_match(term) in text
        for term in PROFESSIONAL_TERMS
    )

    healthcare = any(
        normalize_for_match(term) in text
        for term in HEALTHCARE_TERMS
    )

    trusted_source = (
        is_institutional_source(result.url)
        or is_professional_source(result.url)
        or is_linkedin_source(result.url)
    )

    return (
        identity >= 180
        and professional
        and healthcare
        and trusted_source
    )


def python_auto_result(
    person: Person,
    evidence: CollectedEvidence,
) -> dict[str, Any] | None:
    # 1) CV realmente verificato: è il caso più forte.
    if evidence.cv_source and evidence.cv_text:
        text = evidence.cv_text

        role = detect_role(text)
        email = extract_first_email(text)
        phone = extract_first_phone(text)

        notes = (
            "Completato senza Gemini: CV PDF verificato sul contenuto "
            "per identità e struttura tipica di curriculum."
        )

        return {
            "found": True,
            "facility": "",
            "department": "",
            "city": "",
            "professional_email": email,
            "professional_phone": phone,
            "role_found": role,
            "confidence": "alta",
            "notes": notes,
            "supporting_indexes": [],
            "sources": [evidence.cv_source.url],
            "query_index": evidence.cv_source.query_index,
            "query_text": evidence.cv_source.query_text,
            "cv_path": (
                str(evidence.cv_path)
                if evidence.cv_path
                else ""
            ),
            "python_auto": True,
        }

    # 2) Fonte istituzionale/professionale/LinkedIn molto forte.
    strong = [
        result
        for result in evidence.results
        if strong_result_for_python(
            result,
            person,
        )
    ]

    if not strong:
        return None

    strong.sort(
        key=lambda result: result_score(
            result,
            person,
        ),
        reverse=True,
    )

    best = strong[0]
    combined = clean_text(
        f"{best.title} {best.snippet} {evidence.primary_text}"
    )

    role = detect_role(combined)
    email = extract_first_email(combined)
    phone = extract_first_phone(combined)

    confidence = (
        "alta"
        if is_institutional_source(best.url)
        else "media"
    )

    source_label = (
        "fonte istituzionale"
        if is_institutional_source(best.url)
        else (
            "fonte professionale"
            if is_professional_source(best.url)
            else "LinkedIn professionale"
        )
    )

    return {
        "found": True,
        "facility": "",
        "department": "",
        "city": "",
        "professional_email": email,
        "professional_phone": phone,
        "role_found": role,
        "confidence": confidence,
        "notes": (
            "Completato senza Gemini: evidenza Python forte da "
            f"{source_label}, con identità e ruolo sanitario coerenti. "
            "I campi non esplicitamente presenti nella fonte restano vuoti."
        ),
        "supporting_indexes": [],
        "sources": [best.url],
        "query_index": best.query_index,
        "query_text": best.query_text,
        "cv_path": "",
        "python_auto": True,
    }


# ============================================================
# RACCOLTA DATI WEB - UNA SOLA VOLTA
# ============================================================

def collect_evidence(
    person: Person,
    search_results_limit: int,
    cv_dir: Path,
    search_backend: str,
) -> CollectedEvidence:
    logging.info(
        "RACCOLTA WEB | Pers_ID=%s | inizio",
        person.pers_id,
    )

    ordinary_results = search_person(
        person,
        search_results_limit,
        search_backend,
    )

    cv_results = search_cv(
        person,
        search_results_limit,
        search_backend,
    )

    all_results = deduplicate_results(
        ordinary_results
        + cv_results
    )

    all_results.sort(
        key=lambda result: result_score(
            result,
            person,
        ),
        reverse=True,
    )

    logging.info(
        "RACCOLTA WEB | risultati totali unici=%s",
        len(all_results),
    )

    cv_path: Path | None = None
    cv_text = ""
    cv_source: SearchResult | None = None

    cv_candidates = [
        result
        for result in all_results
        if should_try_as_cv(
            result,
            person,
        )
    ]

    cv_candidates.sort(
        key=lambda result: cv_candidate_score(
            result,
            person,
        ),
        reverse=True,
    )

    logging.info(
        "CANDIDATI CV SEVERI | %s",
        len(cv_candidates),
    )

    for candidate in cv_candidates[:12]:
        found = download_cv(
            candidate,
            person,
            cv_dir,
        )

        if found:
            cv_path, cv_text = found
            cv_source = candidate
            break

    primary: SearchResult | None = None
    primary_text = ""

    if cv_source is not None:
        primary = cv_source
        logging.info(
            "FONTE PRIMARIA = CV VERIFICATO | %s",
            cv_source.url,
        )
    elif all_results:
        primary = all_results[0]
        logging.info(
            "FONTE PRIMARIA WEB | score=%s | %s",
            result_score(primary, person),
            primary.url,
        )
        primary_text = fetch_page_text(
            primary.url
        )

    logging.info(
        "RACCOLTA WEB | Pers_ID=%s | fine | CV=%s",
        person.pers_id,
        "SI" if cv_source else "NO",
    )

    return CollectedEvidence(
        results=all_results,
        primary=primary,
        primary_text=primary_text,
        cv_source=cv_source,
        cv_path=cv_path,
        cv_text=cv_text,
    )


# ============================================================
# CONTEXT GEMINI
# ============================================================

def source_type_for(
    result: SearchResult,
) -> str:
    if looks_like_cv(result):
        return "CV/PDF CANDIDATO"
    if is_institutional_source(result.url):
        return "ISTITUZIONALE"
    if is_professional_source(result.url):
        return "PROFESSIONALE"
    if is_linkedin_source(result.url):
        return "LINKEDIN PROFESSIONALE"
    if is_weak_source(result.url):
        return "SOCIAL/DEBOLE"
    return "ALTRO"


def build_gemini_context(
    person: Person,
    evidence: CollectedEvidence,
    max_results: int,
) -> tuple[str, list[SearchResult]]:
    ranked = sorted(
        evidence.results,
        key=lambda result: result_score(
            result,
            person,
        ),
        reverse=True,
    )[:max_results]

    parts: list[str] = []
    chars = 0

    for index, result in enumerate(
        ranked,
        start=1,
    ):
        block = f"""
[RISULTATO {index}]
Categoria: {result.category}
Livello ricerca: {result.query_index}
Backend: {result.backend}
Tipo fonte: {source_type_for(result)}
Score Python: {result_score(result, person)}
Titolo: {result.title}
URL: {result.url}
Snippet: {result.snippet}
""".strip()

        if (
            chars
            + len(block)
            > MAX_CONTEXT_CHARS
        ):
            break

        parts.append(block)
        chars += len(block)

    if evidence.cv_text:
        remaining = (
            MAX_CONTEXT_CHARS
            - chars
        )

        if remaining > 1000:
            cv_block = (
                "[CV PDF VERIFICATO DELLA PERSONA]\n"
                + evidence.cv_text[:remaining]
            )
            parts.append(cv_block)
            chars += len(cv_block)

    if evidence.primary_text:
        remaining = (
            MAX_CONTEXT_CHARS
            - chars
        )

        if remaining > 500:
            parts.append(
                "[CONTENUTO FONTE PRIMARIA]\n"
                + evidence.primary_text[:remaining]
            )

    return "\n\n".join(parts), ranked


# ============================================================
# GEMINI - SOLO ANALISI, NESSUNA RICERCA WEB
# ============================================================

def analyze_with_gemini_once(
    client: genai.Client,
    model: str,
    person: Person,
    evidence: CollectedEvidence,
    results_to_gemini: int,
) -> dict[str, Any]:
    context, ranked = build_gemini_context(
        person=person,
        evidence=evidence,
        max_results=results_to_gemini,
    )

    prompt = f"""
Analizza esclusivamente le informazioni già raccolte da Python.

NON hai accesso al web.
NON utilizzare conoscenze esterne.
NON inventare informazioni.

IDENTIFICATIVO RECORD
Pers_ID: {person.pers_id}

PERSONA
Nome: {person.name}
Cognome: {person.surname}
Data di nascita: {person.birth_date or "non disponibile"}
Professione attesa: {DEFAULT_ROLE}

OBIETTIVO

Verificare se questa specifica persona è o è stata:
- farmacista ospedaliero/a;
- dirigente farmacista;
- farmacista operante presso una struttura sanitaria.

Se verificabile, estrai:
- struttura sanitaria;
- reparto / UO / farmacia ospedaliera;
- città;
- ruolo;
- email professionale;
- telefono professionale.

GERARCHIA DELLE FONTI

1. CV PDF verificato sul contenuto;
2. documenti ufficiali / PDF istituzionali;
3. siti ASL / AUSL / ASST / ospedali / IRCCS;
4. SIFO, Ordini, università e fonti professionali;
5. LinkedIn professionale;
6. altre fonti.

REGOLE

1. La data di nascita serve per disambiguare omonimi.
2. found=true solo se la persona è ragionevolmente collegata
   alla professione o a una struttura sanitaria.
3. Non inventare struttura, reparto, email o telefono.
4. Non dedurre indirizzi email da pattern.
5. Email e telefono possono essere restituiti solo se presenti
   letteralmente nel materiale fornito.
6. Non restituire recapiti personali/privati.
7. LinkedIn è una fonte professionale valida se il profilo è
   chiaramente riferibile alla persona.
8. Se trovi solo "farmacista" senza collegamento sanitario,
   riduci la confidenza.
9. Se esistono omonimi o contraddizioni importanti,
   usa confidenza bassa oppure found=false.
10. supporting_indexes contiene gli indici 1-based dei soli
    RISULTATI WEB usati come supporto.

MATERIALE RACCOLTO

{context}
""".strip()

    interaction = client.interactions.create(
        model=model,
        input=prompt,
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
            "Gemini ha restituito risposta vuota."
        )

    result = json.loads(output_text)

    if not isinstance(result, dict):
        raise ValueError(
            "Gemini non ha restituito un oggetto JSON."
        )

    selected_sources: list[str] = []

    for raw_index in (
        result.get("supporting_indexes")
        or []
    ):
        try:
            index = int(raw_index) - 1
        except (TypeError, ValueError):
            continue

        if 0 <= index < len(ranked):
            url = ranked[index].url

            if url not in selected_sources:
                selected_sources.append(url)

    if (
        result.get("found")
        and not selected_sources
        and ranked
    ):
        selected_sources.append(
            ranked[0].url
        )

    if (
        evidence.cv_source
        and evidence.cv_source.url not in selected_sources
        and result.get("found")
    ):
        selected_sources.insert(
            0,
            evidence.cv_source.url,
        )

    result["sources"] = selected_sources
    result["cv_path"] = (
        str(evidence.cv_path)
        if evidence.cv_path
        else ""
    )

    if evidence.primary:
        result["query_index"] = evidence.primary.query_index
        result["query_text"] = evidence.primary.query_text
    else:
        result["query_index"] = ""
        result["query_text"] = ""

    return result


def extract_retry_seconds(
    exc: Exception,
) -> float | None:
    text = str(exc)

    patterns = (
        r"retry in\s+([0-9.]+)s",
        r"retry after\s+([0-9.]+)",
        r"retry_after['\"]?\s*[:=]\s*([0-9.]+)",
    )

    for pattern in patterns:
        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE,
        )

        if match:
            try:
                return float(
                    match.group(1)
                )
            except ValueError:
                pass

    return None


def is_rate_limit_error(
    exc: Exception,
) -> bool:
    text = str(exc).casefold()

    return (
        "429" in text
        or "too_many_requests" in text
        or "rate limit" in text
        or "quota exceeded" in text
        or "resource_exhausted" in text
    )


def analyze_with_gemini_retry(
    client: genai.Client,
    model: str,
    person: Person,
    evidence: CollectedEvidence,
    results_to_gemini: int,
    max_retries: int,
) -> tuple[dict[str, Any], int]:
    """
    IMPORTANTE V6.4:
    i retry riguardano SOLO Gemini.
    Le ricerche web e il download CV NON vengono ripetuti.
    """
    last_error: Exception | None = None

    for attempt in range(
        1,
        max_retries + 1,
    ):
        try:
            logging.info(
                "GEMINI | tentativo analisi %s/%s | Pers_ID=%s | modello=%s",
                attempt,
                max_retries,
                person.pers_id,
                model,
            )

            result = analyze_with_gemini_once(
                client=client,
                model=model,
                person=person,
                evidence=evidence,
                results_to_gemini=results_to_gemini,
            )

            return result, attempt

        except (KeyboardInterrupt, SystemExit):
            raise

        except Exception as exc:
            last_error = exc

            logging.warning(
                "GEMINI FALLITO | tentativo=%s/%s | Pers_ID=%s | %s: %s",
                attempt,
                max_retries,
                person.pers_id,
                type(exc).__name__,
                exc,
            )

            if attempt >= max_retries:
                break

            if is_rate_limit_error(exc):
                retry_seconds = extract_retry_seconds(exc)

                if retry_seconds is None:
                    retry_seconds = 20.0

                wait = min(
                    90.0,
                    retry_seconds + 2.0,
                )
            else:
                wait = min(
                    20.0,
                    (2 ** (attempt - 1))
                    + random.uniform(0.0, 1.0),
                )

            logging.info(
                "GEMINI | attesa retry %.1f secondi; "
                "la ricerca web NON verrà ripetuta.",
                wait,
            )
            time.sleep(wait)

    if last_error and is_rate_limit_error(last_error):
        raise GeminiDeferredError(
            f"Gemini temporaneamente non disponibile/quota: {last_error}"
        )

    raise RuntimeError(
        str(last_error)
        if last_error
        else "Errore Gemini sconosciuto"
    )


# ============================================================
# PIPELINE PERSONA
# ============================================================

def empty_result(
    person: Person,
    note: str,
    evidence: CollectedEvidence | None = None,
) -> dict[str, Any]:
    primary = (
        evidence.primary
        if evidence
        else None
    )

    return {
        "found": False,
        "facility": "",
        "department": "",
        "city": "",
        "professional_email": "",
        "professional_phone": "",
        "role_found": "",
        "confidence": "nessuna",
        "notes": note,
        "supporting_indexes": [],
        "sources": (
            [primary.url]
            if primary
            else []
        ),
        "query_index": (
            primary.query_index
            if primary
            else 6
        ),
        "query_text": (
            primary.query_text
            if primary
            else build_cv_queries(person)[-1]
        ),
        "cv_path": (
            str(evidence.cv_path)
            if evidence and evidence.cv_path
            else ""
        ),
    }


def deferred_result(
    person: Person,
    evidence: CollectedEvidence,
    error: Exception,
) -> dict[str, Any]:
    primary = evidence.primary

    source_urls = [
        result.url
        for result in evidence.results[:5]
    ]

    if evidence.cv_source:
        source_urls.insert(
            0,
            evidence.cv_source.url,
        )

    source_urls = list(
        dict.fromkeys(source_urls)
    )

    return {
        "found": False,
        "deferred": True,
        "facility": "",
        "department": "",
        "city": "",
        "professional_email": "",
        "professional_phone": "",
        "role_found": "",
        "confidence": "nessuna",
        "notes": (
            "Ricerca web completata e conservata. "
            "Analisi Gemini rinviata per limite/quota temporanea; "
            "non è un NESSUN_RISULTATO definitivo."
        ),
        "supporting_indexes": [],
        "sources": source_urls,
        "query_index": (
            primary.query_index
            if primary
            else 6
        ),
        "query_text": (
            primary.query_text
            if primary
            else build_cv_queries(person)[-1]
        ),
        "cv_path": (
            str(evidence.cv_path)
            if evidence.cv_path
            else ""
        ),
        "deferred_error": clean(error)[:1500],
    }


def research_person(
    client: genai.Client,
    model: str,
    person: Person,
    search_results_limit: int,
    results_to_gemini: int,
    max_retries: int,
    cv_dir: Path,
    search_backend: str,
) -> tuple[dict[str, Any], int]:
    """
    V6.4:
    1) raccoglie web/CV UNA sola volta;
    2) prova decisione Python;
    3) Gemini solo se il caso rimane ambiguo;
    4) retry Gemini senza rifare il web.
    """
    evidence = collect_evidence(
        person=person,
        search_results_limit=search_results_limit,
        cv_dir=cv_dir,
        search_backend=search_backend,
    )

    if not evidence.results:
        logging.info(
            "DECISIONE | nessun risultato web utilizzabile | "
            "Gemini non necessario."
        )

        return (
            empty_result(
                person,
                "Nessun risultato utilizzabile restituito "
                "dai provider di ricerca.",
                evidence,
            ),
            0,
        )

    auto = python_auto_result(
        person,
        evidence,
    )

    if auto is not None:
        logging.info(
            "DECISIONE | PYTHON AUTO | Gemini saltato | "
            "Pers_ID=%s | confidenza=%s",
            person.pers_id,
            auto.get("confidence"),
        )

        return auto, 0

    logging.info(
        "DECISIONE | caso ambiguo | Gemini necessario | "
        "Pers_ID=%s | modello=%s",
        person.pers_id,
        model,
    )

    try:
        return analyze_with_gemini_retry(
            client=client,
            model=model,
            person=person,
            evidence=evidence,
            results_to_gemini=results_to_gemini,
            max_retries=max_retries,
        )

    except GeminiDeferredError as exc:
        logging.warning(
            "GEMINI RINVIATO | Pers_ID=%s | batch continua | %s",
            person.pers_id,
            exc,
        )

        return deferred_result(
            person,
            evidence,
            exc,
        ), max_retries


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
        row,
        headers[label.casefold()],
        value,
    )


def write_result(
    ws,
    row: int,
    headers: dict[str, int],
    result: dict[str, Any],
    attempts: int,
) -> None:
    if result.get("deferred"):
        status = "DA_RIPROVARE_GEMINI"
    else:
        status = (
            "COMPLETATO"
            if result.get("found")
            else "NESSUN_RISULTATO"
        )

    error_value = ""

    if result.get("deferred"):
        error_value = clean(
            result.get("deferred_error")
        )

    values = {
        "Stato ricerca": status,
        "Indice ricerca": result.get("query_index", ""),
        "Termine ricerca usato": clean(result.get("query_text")),
        "Struttura": clean(result.get("facility")),
        "Reparto/UO": clean(result.get("department")),
        "Citta": clean(result.get("city")),
        "Email professionale": clean(result.get("professional_email")),
        "Telefono professionale": clean(result.get("professional_phone")),
        "Ruolo trovato": clean(result.get("role_found")),
        "Confidenza": clean(result.get("confidence")),
        "Note": clean(result.get("notes")),
        "Fonti": "\n".join(result.get("sources") or []),
        "CV salvato": clean(result.get("cv_path")),
        "Tentativi": attempts,
        "Ultimo aggiornamento UTC": utc_now(),
        "Errore": error_value,
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

    cv_dir = (
        args.cv_dir
        .expanduser()
        .resolve()
    )
    cv_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    api_key = os.getenv(
        "GEMINI_API_KEY"
    )

    if not api_key:
        raise RuntimeError(
            "GEMINI_API_KEY non trovata."
        )

    input_path = (
        args.input
        .expanduser()
        .resolve()
    )

    output_path = output_path_for(
        input_path,
        args.output,
    ).expanduser().resolve()

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

    validate_pers_ids(
        ws,
        headers,
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
    completed = 0
    no_results = 0
    deferred = 0
    errors = 0
    skipped = 0

    logging.info("============================================")
    logging.info("RICERCA FARMACISTI - %s", VERSION)
    logging.info("Input: %s", input_path)
    logging.info("Output: %s", output_path)
    logging.info("Cartella CV: %s", cv_dir)
    logging.info("Pers_ID: OBBLIGATORIO")
    logging.info("Backend ricerca configurato: %s", args.search_backend)
    logging.info("Gemini modello configurato: %s", args.model)
    logging.info("Gemini modello effettivo richiesto all'API: %s", args.model)
    logging.info("Web e Gemini separati: ATTIVO")
    logging.info("Retry Gemini senza rifare web: ATTIVO")
    logging.info("Decisione Python per casi forti: ATTIVA")
    logging.info("CV falso-positivo: filtro severo ATTIVO")
    logging.info(
        "Priorità: CV verificato > istituzionale > "
        "professionale > LinkedIn > Gemini per ambigui"
    )
    logging.info("Log: %s", log_file)
    logging.info("============================================")

    for row in range(
        max(2, args.start_row),
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
                headers["stato ricerca"],
            ).value
        ).upper()

        if status == "COMPLETATO":
            skipped += 1
            continue

        if (
            status == "NESSUN_RISULTATO"
            and not args.retry_no_results
        ):
            skipped += 1
            continue

        if (
            status == "ERRORE"
            and not args.retry_errors
        ):
            skipped += 1
            continue

        # DA_RIPROVARE_GEMINI non viene saltato:
        # al run successivo viene riprocessato.
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

        if not person.pers_id:
            raise RuntimeError(
                f"Pers_ID mancante alla riga {row}"
            )

        logging.info("--------------------------------------------")
        logging.info(
            "RIGA %s | Pers_ID=%s | %s %s | nascita=%s",
            row,
            person.pers_id,
            person.name,
            person.surname,
            person.birth_date or "N/D",
        )

        try:
            result, attempts = research_person(
                client=client,
                model=args.model,
                person=person,
                search_results_limit=args.search_results,
                results_to_gemini=args.results_to_gemini,
                max_retries=args.max_retries,
                cv_dir=cv_dir,
                search_backend=args.search_backend,
            )

            write_result(
                ws,
                row,
                headers,
                result,
                attempts,
            )

            if result.get("deferred"):
                deferred += 1
            elif result.get("found"):
                completed += 1
            else:
                no_results += 1

            logging.info(
                "RISULTATO | Pers_ID=%s | stato=%s | "
                "python_auto=%s | tentativi_gemini=%s",
                person.pers_id,
                (
                    "DA_RIPROVARE_GEMINI"
                    if result.get("deferred")
                    else (
                        "COMPLETATO"
                        if result.get("found")
                        else "NESSUN RISULTATO"
                    )
                ),
                bool(result.get("python_auto")),
                attempts,
            )

            logging.info(
                "Struttura: %s",
                clean(result.get("facility")) or "N/D",
            )
            logging.info(
                "Ruolo: %s",
                clean(result.get("role_found")) or "N/D",
            )
            logging.info(
                "Confidenza: %s",
                clean(result.get("confidence")) or "N/D",
            )
            logging.info(
                "CV: %s",
                clean(result.get("cv_path")) or "NON TROVATO",
            )

        except (KeyboardInterrupt, SystemExit):
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
                "ERRORE | Pers_ID=%s | Riga=%s",
                person.pers_id,
                row,
            )

        atomic_save(
            workbook,
            output_path,
        )

        processed += 1

        logging.info(
            "Checkpoint Excel salvato."
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

    logging.info("============================================")
    logging.info("FINE ESECUZIONE %s", VERSION)
    logging.info("Elaborate: %s", processed)
    logging.info("Completate: %s", completed)
    logging.info("Nessun risultato: %s", no_results)
    logging.info("Da riprovare Gemini: %s", deferred)
    logging.info("Errori: %s", errors)
    logging.info("Saltate: %s", skipped)
    logging.info("Output: %s", output_path)
    logging.info("CV: %s", cv_dir)
    logging.info("Log: %s", log_file)
    logging.info("============================================")

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(
            main()
        )
    except KeyboardInterrupt:
        print("\nInterrotto.")
        raise SystemExit(130)

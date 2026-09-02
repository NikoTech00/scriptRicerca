from __future__ import annotations

import argparse
import hashlib
import ipaddress
import json
import logging
import os
import re
import shutil
import socket
import tempfile
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, asdict
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urljoin, urlparse

# Third-party imports are loaded lazily/guarded.
# In questo modo, se manca una dipendenza, lo script crea comunque il log
# e una copia dell'Excel di output, poi termina con un messaggio leggibile.
requests = None
BeautifulSoup = None
DDGS = None
load_dotenv = None
load_workbook = None
Alignment = Font = PatternFill = None
PdfReader = None
genai = None


# ============================================================
# VERSIONE / OBIETTIVO
# ============================================================

VERSION = "V1.0 MEDICI - SPECIALITA + CV"

# Input reale verificato su scriptMedici.xlsx:
# Pers_Id, Pers_Cognome, Pers_Nome, Pers_DataNascita, Pers_CodFis,
# Indirizzi_Citta, Medico_Id, Email, ecc.
#
# Strategia:
# 1) Cerca prima il CV: se verificato, lo salva e prova a ricavare la specialità.
# 2) Se il CV non basta, cerca fonti affidabili sulla specialità.
# 3) L'AI è opzionale e lavora SOLO sulle evidenze già raccolte.
# 4) Nessun dato viene inventato.
# 5) Elaborazione concorrente + cache + resume.

OUTPUT_COLUMNS = (
    "Ricerca_Stato",
    "Specialita",
    "Specialita_Confidenza",
    "CV_Salvato",
    "CV_URL",
    "Fonti_Ricerca",
    "Ricerca_Note",
    "Ultimo_Aggiornamento_UTC",
    "Ricerca_Errore",
)

DEFAULT_SHEET = "Foglio1"
DEFAULT_WORKERS = 8
DEFAULT_SEARCH_CONCURRENCY = 4
DEFAULT_SEARCH_RESULTS = 6
DEFAULT_BACKEND = "duckduckgo"
DEFAULT_SAVE_EVERY = 20

HTTP_TIMEOUT = 8
MAX_PDF_BYTES = 20 * 1024 * 1024
MAX_HTML_BYTES = 4 * 1024 * 1024
MAX_PDF_PAGES = 80
MAX_TEXT_CHARS = 60000

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/151.0 Safari/537.36"
)

SEARCH_SEMAPHORE: threading.BoundedSemaphore | None = None


# ============================================================
# BOOTSTRAP / PREFLIGHT DIPENDENZE
# ============================================================

def bootstrap_log_path(log_dir: Path) -> Path:
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir / f"startup_medici_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"


def write_bootstrap_log(path: Path, message: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as fh:
        fh.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {message}\n")


def import_dependencies(log_path: Path) -> tuple[bool, list[str]]:
    global requests, BeautifulSoup, DDGS, load_dotenv
    global load_workbook, Alignment, Font, PatternFill, PdfReader, genai

    missing: list[str] = []

    try:
        import requests as _requests
        requests = _requests
    except Exception as exc:
        missing.append(f"requests ({exc})")

    try:
        from bs4 import BeautifulSoup as _BeautifulSoup
        BeautifulSoup = _BeautifulSoup
    except Exception as exc:
        missing.append(f"beautifulsoup4 ({exc})")

    try:
        from ddgs import DDGS as _DDGS
        DDGS = _DDGS
    except Exception as exc:
        missing.append(f"ddgs ({exc})")

    try:
        from dotenv import load_dotenv as _load_dotenv
        load_dotenv = _load_dotenv
    except Exception as exc:
        missing.append(f"python-dotenv ({exc})")

    try:
        from openpyxl import load_workbook as _load_workbook
        from openpyxl.styles import Alignment as _Alignment
        from openpyxl.styles import Font as _Font
        from openpyxl.styles import PatternFill as _PatternFill
        load_workbook = _load_workbook
        Alignment = _Alignment
        Font = _Font
        PatternFill = _PatternFill
    except Exception as exc:
        missing.append(f"openpyxl ({exc})")

    try:
        from pypdf import PdfReader as _PdfReader
        PdfReader = _PdfReader
    except Exception as exc:
        missing.append(f"pypdf ({exc})")

    # Gemini è opzionale: non blocca l'esecuzione.
    try:
        from google import genai as _genai
        genai = _genai
    except Exception as exc:
        genai = None
        write_bootstrap_log(
            log_path,
            f"INFO | google-genai non disponibile: {type(exc).__name__}: {exc}. "
            "Lo script potrà funzionare senza AI."
        )

    if missing:
        write_bootstrap_log(
            log_path,
            "ERRORE | Dipendenze obbligatorie mancanti: " + "; ".join(missing)
        )
        return False, missing

    write_bootstrap_log(log_path, "OK | Dipendenze obbligatorie caricate correttamente.")
    return True, []


def parse_early_paths(argv: list[str]) -> tuple[Path | None, Path | None, Path]:
    """
    Ricava input/output/log-dir prima di argparse e prima delle dipendenze.
    Serve per creare sempre un log e, quando possibile, la copia Excel di output.
    """
    input_path: Path | None = None
    output_path: Path | None = None
    log_dir = Path("logs")

    # primo argomento non-opzione = input
    for i, arg in enumerate(argv[1:], start=1):
        if not arg.startswith("-"):
            input_path = Path(arg)
            break

    if "--output" in argv:
        try:
            output_path = Path(argv[argv.index("--output") + 1])
        except Exception:
            pass

    if "--log-dir" in argv:
        try:
            log_dir = Path(argv[argv.index("--log-dir") + 1])
        except Exception:
            pass

    return input_path, output_path, log_dir


def create_initial_output_copy(input_path: Path | None, output_path: Path | None) -> Path | None:
    if input_path is None:
        return None

    input_path = input_path.expanduser().resolve()
    if output_path is None:
        output_path = input_path.with_name(f"{input_path.stem}_specialita_cv.xlsx")
    else:
        output_path = output_path.expanduser().resolve()

    if input_path.exists() and input_path.suffix.casefold() == ".xlsx":
        output_path.parent.mkdir(parents=True, exist_ok=True)
        if not output_path.exists() and input_path != output_path:
            shutil.copy2(input_path, output_path)
    return output_path

# ============================================================
# MODELLI
# ============================================================

@dataclass(frozen=True)
class Person:
    row: int
    pers_id: str
    medico_id: str
    surname: str
    name: str
    birth_date: str
    fiscal_code: str
    city: str
    email: str

    @property
    def full_name(self) -> str:
        return f"{self.name} {self.surname}".strip()

    @property
    def output_code(self) -> str:
        # "codice del medico": usa Medico_Id se realmente valorizzato.
        # Nel file molti Medico_Id sono 0, quindi fallback obbligatorio a Pers_Id.
        mid = clean(self.medico_id)
        if mid and mid not in {"0", "0.0"}:
            return numericish(mid)
        return numericish(self.pers_id)


@dataclass(frozen=True)
class SearchResult:
    title: str
    url: str
    snippet: str
    query: str
    provider: str
    category: str


@dataclass
class Evidence:
    results: list[SearchResult]
    cv_path: str = ""
    cv_url: str = ""
    cv_text: str = ""
    specialty: str = ""
    specialty_confidence: str = ""
    specialty_source: str = ""
    notes: list[str] | None = None

    def __post_init__(self):
        if self.notes is None:
            self.notes = []


# ============================================================
# SPECIALITA - TASSONOMIA / NORMALIZZAZIONE
# ============================================================

# Include le principali scuole/discipline italiane e sinonimi frequenti.
SPECIALTY_ALIASES: dict[str, tuple[str, ...]] = {
    "Allergologia e Immunologia Clinica": (
        "allergologia e immunologia clinica", "allergologia", "immunologia clinica"
    ),
    "Anatomia Patologica": ("anatomia patologica", "anatomopatologia"),
    "Anestesia, Rianimazione, Terapia Intensiva e del Dolore": (
        "anestesia rianimazione terapia intensiva e del dolore",
        "anestesia e rianimazione",
        "anestesiologia",
        "rianimazione",
        "terapia intensiva",
    ),
    "Audiologia e Foniatria": ("audiologia e foniatria", "foniatria", "audiologia"),
    "Cardiochirurgia": ("cardiochirurgia",),
    "Cardiologia": ("cardiologia", "malattie dell'apparato cardiovascolare"),
    "Chirurgia Generale": ("chirurgia generale",),
    "Chirurgia Maxillo-Facciale": ("chirurgia maxillo facciale", "chirurgia maxillo-facciale"),
    "Chirurgia Pediatrica": ("chirurgia pediatrica",),
    "Chirurgia Plastica, Ricostruttiva ed Estetica": (
        "chirurgia plastica ricostruttiva ed estetica",
        "chirurgia plastica",
    ),
    "Chirurgia Toracica": ("chirurgia toracica",),
    "Chirurgia Vascolare": ("chirurgia vascolare",),
    "Dermatologia e Venereologia": ("dermatologia e venereologia", "dermatologia"),
    "Ematologia": ("ematologia",),
    "Endocrinologia e Malattie del Metabolismo": (
        "endocrinologia e malattie del metabolismo",
        "endocrinologia",
        "malattie del metabolismo",
        "diabetologia",
    ),
    "Farmacologia e Tossicologia Clinica": (
        "farmacologia e tossicologia clinica", "farmacologia clinica", "tossicologia clinica"
    ),
    "Genetica Medica": ("genetica medica",),
    "Geriatria": ("geriatria",),
    "Ginecologia e Ostetricia": ("ginecologia e ostetricia", "ostetricia e ginecologia", "ginecologia"),
    "Igiene e Medicina Preventiva": (
        "igiene e medicina preventiva", "igiene medicina preventiva", "igiene"
    ),
    "Malattie dell'Apparato Digerente": (
        "malattie dell'apparato digerente", "gastroenterologia", "gastroenterologia ed endoscopia digestiva"
    ),
    "Malattie dell'Apparato Respiratorio": (
        "malattie dell'apparato respiratorio", "pneumologia", "pneumologia e tisiologia"
    ),
    "Malattie Infettive e Tropicali": (
        "malattie infettive e tropicali", "malattie infettive", "infettivologia"
    ),
    "Medicina d'Emergenza-Urgenza": (
        "medicina d'emergenza urgenza", "medicina d'emergenza-urgenza",
        "medicina di emergenza urgenza", "medicina di emergenza-urgenza",
        "medicina e chirurgia d'accettazione e d'urgenza",
    ),
    "Medicina del Lavoro": ("medicina del lavoro",),
    "Medicina dello Sport e dell'Esercizio Fisico": (
        "medicina dello sport e dell'esercizio fisico", "medicina dello sport"
    ),
    "Medicina Fisica e Riabilitativa": (
        "medicina fisica e riabilitativa", "fisiatria", "medicina riabilitativa"
    ),
    "Medicina Interna": ("medicina interna",),
    "Medicina Legale": ("medicina legale",),
    "Medicina Nucleare": ("medicina nucleare",),
    "Microbiologia e Virologia": ("microbiologia e virologia", "microbiologia", "virologia"),
    "Nefrologia": ("nefrologia",),
    "Neurochirurgia": ("neurochirurgia",),
    "Neurologia": ("neurologia",),
    "Neuropsichiatria Infantile": ("neuropsichiatria infantile",),
    "Oftalmologia": ("oftalmologia", "oculistica"),
    "Oncologia Medica": ("oncologia medica", "oncologia"),
    "Ortopedia e Traumatologia": ("ortopedia e traumatologia", "ortopedia", "traumatologia"),
    "Otorinolaringoiatria": ("otorinolaringoiatria", "otorino"),
    "Patologia Clinica e Biochimica Clinica": (
        "patologia clinica e biochimica clinica", "patologia clinica", "biochimica clinica"
    ),
    "Pediatria": ("pediatria",),
    "Psichiatria": ("psichiatria",),
    "Radiodiagnostica": ("radiodiagnostica", "radiologia", "diagnostica per immagini"),
    "Radioterapia": ("radioterapia", "radioterapia oncologica"),
    "Reumatologia": ("reumatologia",),
    "Scienza dell'Alimentazione": ("scienza dell'alimentazione", "scienze dell'alimentazione"),
    "Statistica Sanitaria e Biometria": ("statistica sanitaria e biometria",),
    "Urologia": ("urologia",),
}

# Contesti che NON vanno scambiati per specialità.
NON_SPECIALTY_TERMS = (
    "laurea in medicina e chirurgia",
    "medico chirurgo",
    "iscritto all'ordine",
    "abilitazione alla professione",
    "master",
    "dottorato",
)

CV_POSITIVE_TERMS = (
    "curriculum vitae",
    "curriculum professionale",
    "curriculum formativo",
    "curriculum scientifico",
    "europass",
    "esperienza professionale",
    "esperienze professionali",
    "istruzione e formazione",
    "titoli di studio",
)

CV_NEGATIVE_TERMS = (
    "graduatoria",
    "elenco candidati",
    "candidati ammessi",
    "concorso",
    "commissione esaminatrice",
    "verbale",
    "bando",
    "avviso pubblico",
    "prova scritta",
    "prova orale",
    "delibera",
    "deliberazione",
    "determinazione",
)

MEDICAL_CONTEXT_TERMS = (
    "medico",
    "medicina",
    "chirurgo",
    "specialista",
    "specializzazione",
    "ospedale",
    "azienda sanitaria",
    "asl",
    "ausl",
    "asst",
    "aou",
    "irccs",
    "policlinico",
    "università",
    "universita",
    "ordine dei medici",
)

# ============================================================
# CLI
# ============================================================

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description=f"{VERSION}: recupero Specialità + CV da Excel medici."
    )
    p.add_argument("input", type=Path)
    p.add_argument("--output", type=Path)
    p.add_argument("--sheet", default=DEFAULT_SHEET)
    p.add_argument("--limit", type=int)
    p.add_argument("--start-row", type=int, default=2)
    p.add_argument("--workers", type=int, default=DEFAULT_WORKERS)
    p.add_argument("--search-concurrency", type=int, default=DEFAULT_SEARCH_CONCURRENCY)
    p.add_argument("--search-results", type=int, default=DEFAULT_SEARCH_RESULTS)
    p.add_argument("--search-backend", default=DEFAULT_BACKEND)
    p.add_argument("--save-every", type=int, default=DEFAULT_SAVE_EVERY)
    p.add_argument("--retry-completed", action="store_true")
    p.add_argument("--retry-errors", action="store_true")
    p.add_argument("--no-ai", action="store_true")
    p.add_argument("--ai-model", default=os.getenv("GEMINI_MODEL", "gemini-3.7-flash"))
    p.add_argument("--cv-dir", type=Path, default=Path("cv_medici"))
    p.add_argument("--cache-dir", type=Path, default=Path("cache_medici"))
    p.add_argument("--log-dir", type=Path, default=Path("logs"))
    args = p.parse_args()

    if args.start_row < 2:
        p.error("--start-row deve essere >= 2")
    if args.limit is not None and args.limit <= 0:
        p.error("--limit deve essere > 0")
    if args.workers <= 0:
        p.error("--workers deve essere > 0")
    if args.search_concurrency <= 0:
        p.error("--search-concurrency deve essere > 0")
    if args.search_results <= 0:
        p.error("--search-results deve essere > 0")
    if args.save_every <= 0:
        p.error("--save-every deve essere > 0")
    return args


# ============================================================
# LOG / UTILITY
# ============================================================

def configure_logging(log_dir: Path) -> Path:
    log_dir.mkdir(parents=True, exist_ok=True)
    path = log_dir / f"run_medici_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(threadName)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    fh = logging.FileHandler(path, encoding="utf-8")
    sh = logging.StreamHandler()
    fh.setFormatter(fmt)
    sh.setFormatter(fmt)
    logger.addHandler(fh)
    logger.addHandler(sh)
    return path


def clean(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.casefold() in {"nan", "none", "null", "n/a", "nd", "n.d."}:
        return ""
    return text


def numericish(value: str) -> str:
    value = clean(value)
    if re.fullmatch(r"\d+\.0", value):
        return value[:-2]
    return value


def normalize(text: str) -> str:
    text = clean(text).casefold()
    text = (
        text.replace("’", "'")
        .replace("–", "-")
        .replace("—", "-")
    )
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def compact(text: str, max_chars: int = 10000) -> str:
    text = re.sub(r"\s+", " ", clean(text)).strip()
    return text[:max_chars]


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def safe_filename_part(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*]', "_", clean(value))
    value = re.sub(r"\s+", "_", value)
    return value.strip(" ._") or "ND"


def canonical_url(url: str) -> str:
    return clean(url).split("#")[0].rstrip("/").casefold()


def domain(url: str) -> str:
    try:
        return (urlparse(url).hostname or "").casefold()
    except Exception:
        return ""


def format_birth_date(value: Any) -> str:
    if isinstance(value, (date, datetime)):
        return value.strftime("%d/%m/%Y")
    text = clean(value)
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}.*", text):
        try:
            return datetime.fromisoformat(text[:10]).strftime("%d/%m/%Y")
        except Exception:
            pass
    return text


def atomic_json_write(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(
        prefix=f".{path.stem}_", suffix=".json", dir=path.parent
    )
    os.close(fd)
    tmp = Path(temp_name)
    try:
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, path)
    finally:
        tmp.unlink(missing_ok=True)


# ============================================================
# EXCEL
# ============================================================

def normalized_headers(ws) -> dict[str, int]:
    out: dict[str, int] = {}
    for c in ws[1]:
        label = clean(c.value)
        if label:
            out[label.casefold()] = c.column
    return out


def require_columns(headers: dict[str, int]) -> None:
    required = ("Pers_Id", "Pers_Cognome", "Pers_Nome")
    missing = [x for x in required if x.casefold() not in headers]
    if missing:
        raise ValueError("Colonne obbligatorie mancanti: " + ", ".join(missing))


def ensure_output_columns(ws) -> dict[str, int]:
    headers = normalized_headers(ws)
    for label in OUTPUT_COLUMNS:
        if label.casefold() in headers:
            continue
        col = ws.max_column + 1
        cell = ws.cell(1, col, label)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        headers[label.casefold()] = col
    ws.freeze_panes = "A2"
    return headers


def getv(ws, row: int, headers: dict[str, int], name: str) -> Any:
    col = headers.get(name.casefold())
    return ws.cell(row, col).value if col else ""


def person_from_row(ws, row: int, headers: dict[str, int]) -> Person:
    return Person(
        row=row,
        pers_id=numericish(clean(getv(ws, row, headers, "Pers_Id"))),
        medico_id=numericish(clean(getv(ws, row, headers, "Medico_Id"))),
        surname=clean(getv(ws, row, headers, "Pers_Cognome")),
        name=clean(getv(ws, row, headers, "Pers_Nome")),
        birth_date=format_birth_date(getv(ws, row, headers, "Pers_DataNascita")),
        fiscal_code=clean(getv(ws, row, headers, "Pers_CodFis")).upper(),
        city=clean(getv(ws, row, headers, "Indirizzi_Citta")),
        email=clean(getv(ws, row, headers, "emailPredefinita")) or clean(
            getv(ws, row, headers, "Email")
        ),
    )


def output_path_for(input_path: Path, requested: Path | None) -> Path:
    if requested:
        return requested
    return input_path.with_name(f"{input_path.stem}_specialita_cv.xlsx")


def prepare_output(input_path: Path, output_path: Path) -> None:
    if not input_path.exists():
        raise FileNotFoundError(f"Input non trovato: {input_path}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if output_path.exists():
        return
    shutil.copy2(input_path, output_path)


def atomic_save(wb, output_path: Path) -> None:
    fd, tmp_name = tempfile.mkstemp(
        prefix=f".{output_path.stem}_", suffix=".xlsx", dir=output_path.parent
    )
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        wb.save(tmp)
        os.replace(tmp, output_path)
    finally:
        tmp.unlink(missing_ok=True)


def set_cell(ws, row: int, headers: dict[str, int], label: str, value: Any) -> None:
    ws.cell(row, headers[label.casefold()], value)


# ============================================================
# SEARCH PROVIDERS
# ============================================================

def search_serper(query: str, max_results: int) -> list[SearchResult]:
    key = os.getenv("SERPER_API_KEY")
    if not key:
        return []
    r = requests.post(
        "https://google.serper.dev/search",
        headers={"X-API-KEY": key, "Content-Type": "application/json"},
        json={"q": query, "gl": "it", "hl": "it", "num": max_results},
        timeout=HTTP_TIMEOUT,
    )
    r.raise_for_status()
    data = r.json()
    results = []
    for item in data.get("organic", [])[:max_results]:
        url = clean(item.get("link"))
        if url:
            results.append(
                SearchResult(
                    title=clean(item.get("title")),
                    url=url,
                    snippet=clean(item.get("snippet")),
                    query=query,
                    provider="serper",
                    category="",
                )
            )
    return results


def search_brave(query: str, max_results: int) -> list[SearchResult]:
    key = os.getenv("BRAVE_SEARCH_API_KEY")
    if not key:
        return []
    r = requests.get(
        "https://api.search.brave.com/res/v1/web/search",
        headers={"X-Subscription-Token": key, "Accept": "application/json"},
        params={"q": query, "count": max_results, "country": "IT", "search_lang": "it"},
        timeout=HTTP_TIMEOUT,
    )
    r.raise_for_status()
    data = r.json()
    results = []
    for item in data.get("web", {}).get("results", [])[:max_results]:
        url = clean(item.get("url"))
        if url:
            results.append(
                SearchResult(
                    title=clean(item.get("title")),
                    url=url,
                    snippet=clean(item.get("description")),
                    query=query,
                    provider="brave",
                    category="",
                )
            )
    return results


def search_ddgs(query: str, max_results: int, backend: str) -> list[SearchResult]:
    raw = DDGS(timeout=HTTP_TIMEOUT).text(
        query,
        region="it-it",
        safesearch="moderate",
        max_results=max_results,
        backend=backend,
    ) or []
    results = []
    for item in raw:
        url = clean(item.get("href") or item.get("url"))
        if url:
            results.append(
                SearchResult(
                    title=clean(item.get("title")),
                    url=url,
                    snippet=clean(item.get("body") or item.get("snippet")),
                    query=query,
                    provider=f"ddgs:{backend}",
                    category="",
                )
            )
    return results


def run_search(query: str, max_results: int, backend: str) -> list[SearchResult]:
    """
    Priorità:
    1. SERPER_API_KEY (Google) se configurata
    2. BRAVE_SEARCH_API_KEY se configurata
    3. DDGS gratuito

    Non chiama più provider del necessario: il primo che dà risultati utili vince.
    """
    global SEARCH_SEMAPHORE
    if SEARCH_SEMAPHORE is None:
        raise RuntimeError("Search semaphore non inizializzato.")

    with SEARCH_SEMAPHORE:
        providers = []
        if os.getenv("SERPER_API_KEY"):
            providers.append(("serper", lambda: search_serper(query, max_results)))
        if os.getenv("BRAVE_SEARCH_API_KEY"):
            providers.append(("brave", lambda: search_brave(query, max_results)))
        providers.append(("ddgs", lambda: search_ddgs(query, max_results, backend)))

        for provider_name, fn in providers:
            try:
                results = fn()
                if results:
                    logging.info(
                        "SEARCH OK | provider=%s | risultati=%s | %s",
                        provider_name, len(results), query
                    )
                    return dedupe_results(results)
            except Exception as exc:
                logging.warning(
                    "SEARCH FAIL | provider=%s | %s | %s",
                    provider_name, type(exc).__name__, exc
                )
        return []


def dedupe_results(results: Iterable[SearchResult]) -> list[SearchResult]:
    seen: set[str] = set()
    out: list[SearchResult] = []
    for result in results:
        key = canonical_url(result.url)
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(result)
    return out


# ============================================================
# QUERY STRATEGY
# ============================================================

def email_domain(person: Person) -> str:
    if "@" not in person.email:
        return ""
    return person.email.rsplit("@", 1)[-1].strip().casefold()


def build_cv_queries(person: Person) -> list[str]:
    full = f'"{person.full_name}"'
    queries = [
        f'{full} "curriculum vitae" medico',
        f'{full} curriculum filetype:pdf medico',
    ]
    if person.city:
        queries.append(f'{full} "{person.city}" "curriculum vitae"')
    d = email_domain(person)
    if d and not any(x in d for x in ("gmail.", "outlook.", "hotmail.", "libero.", "virgilio.", "yahoo.")):
        queries.insert(0, f'site:{d} {full} curriculum')
    return list(dict.fromkeys(queries))


def build_specialty_queries(person: Person) -> list[str]:
    full = f'"{person.full_name}"'
    queries = [
        f'{full} medico specialista',
        f'{full} "specialista in"',
    ]
    if person.city:
        queries.append(f'{full} medico "{person.city}"')
    d = email_domain(person)
    if d and not any(x in d for x in ("gmail.", "outlook.", "hotmail.", "libero.", "virgilio.", "yahoo.")):
        queries.insert(0, f'site:{d} {full} medico')
    return list(dict.fromkeys(queries))


# ============================================================
# RELEVANCE / TRUST
# ============================================================

def identity_score(text: str, person: Person) -> int:
    n = normalize(text)
    name = normalize(person.name)
    surname = normalize(person.surname)
    full = normalize(person.full_name)
    reverse = normalize(f"{person.surname} {person.name}")

    score = 0
    if full and full in n:
        score += 160
    if reverse and reverse in n:
        score += 160
    if surname and surname in n:
        score += 45
    if name and name in n:
        score += 30
    if person.fiscal_code and normalize(person.fiscal_code) in n:
        score += 500
    return score


def trusted_source_score(url: str) -> int:
    d = domain(url)
    score = 0

    official_tokens = (
        ".gov.it", "regione.", "salute.gov", "sanita",
        "asl", "ausl", "asst", "ats", "aou", "ao.", "ulss", "asp.",
        "irccs", "osped", "policlin", "aziendaospedal", "aziendasanitaria",
        "univ", "unimi", "unibo", "unipd", "unito", "unina", "uniroma",
        "fnomceo", "ordinemedici", "omceo",
    )
    if any(t in d for t in official_tokens):
        score += 120

    professional_tokens = (
        "humanitas", "grupposandonato", "materdomini",
        "multimedica", "gemelli", "auxologico",
    )
    if any(t in d for t in professional_tokens):
        score += 70

    weak = ("facebook.", "instagram.", "tiktok.", "pinterest.", "linkedin.")
    if any(t in d for t in weak):
        score -= 40

    return score


def relevant_result(result: SearchResult, person: Person) -> bool:
    text = f"{result.title} {result.snippet} {result.url}"
    score = identity_score(text, person)
    if score < 160:
        return False

    n = normalize(text)
    return any(t in n for t in MEDICAL_CONTEXT_TERMS) or "curriculum" in n or ".pdf" in n


def rank_result(result: SearchResult, person: Person) -> int:
    text = f"{result.title} {result.snippet} {result.url}"
    n = normalize(text)

    score = identity_score(text, person)
    score += trusted_source_score(result.url)

    if "curriculum vitae" in n:
        score += 120
    elif "curriculum" in n:
        score += 70

    if ".pdf" in result.url.casefold():
        score += 40

    if "specialista in" in n or "specializzazione in" in n:
        score += 90

    for term in CV_NEGATIVE_TERMS:
        if term in n:
            score -= 120

    return score


# ============================================================
# SAFE HTTP / PDF / HTML
# ============================================================

def is_public_http_url(url: str) -> bool:
    try:
        p = urlparse(url)
        if p.scheme not in {"http", "https"} or not p.hostname:
            return False

        # Evita localhost/private network.
        try:
            infos = socket.getaddrinfo(
                p.hostname,
                p.port or (443 if p.scheme == "https" else 80),
                type=socket.SOCK_STREAM,
            )
            for info in infos:
                ip = ipaddress.ip_address(info[4][0])
                if (
                    ip.is_private or ip.is_loopback or ip.is_link_local
                    or ip.is_reserved or ip.is_multicast
                ):
                    return False
        except OSError:
            # DNS può fallire temporaneamente; requests lo gestirà.
            pass
        return True
    except Exception:
        return False


def get_response(url: str, max_bytes: int) -> tuple[requests.Response, bytes] | None:
    if not is_public_http_url(url):
        return None

    try:
        r = requests.get(
            url,
            headers={
                "User-Agent": USER_AGENT,
                "Accept-Language": "it-IT,it;q=0.9,en;q=0.5",
            },
            timeout=HTTP_TIMEOUT,
            allow_redirects=True,
            stream=True,
        )
        r.raise_for_status()

        data = bytearray()
        for chunk in r.iter_content(65536):
            if not chunk:
                continue
            data.extend(chunk)
            if len(data) > max_bytes:
                r.close()
                return None
        return r, bytes(data)
    except Exception:
        return None


def extract_pdf_text(raw: bytes) -> str:
    import io

    try:
        reader = PdfReader(io.BytesIO(raw))
        parts: list[str] = []
        for page in reader.pages[:MAX_PDF_PAGES]:
            text = page.extract_text() or ""
            if text:
                parts.append(text)
        return compact("\n".join(parts), MAX_TEXT_CHARS)
    except Exception:
        return ""


def extract_html_text(raw: bytes) -> str:
    try:
        soup = BeautifulSoup(raw, "html.parser")
        for tag in soup(["script", "style", "nav", "footer", "noscript", "svg"]):
            tag.decompose()
        return compact(soup.get_text(" ", strip=True), 20000)
    except Exception:
        return ""


# ============================================================
# CV VERIFICATION
# ============================================================

def date_variants(date_str: str) -> set[str]:
    if not date_str:
        return set()
    try:
        d = datetime.strptime(date_str, "%d/%m/%Y")
        return {
            d.strftime("%d/%m/%Y"),
            d.strftime("%d-%m-%Y"),
            d.strftime("%d.%m.%Y"),
            d.strftime("%Y-%m-%d"),
        }
    except Exception:
        return {date_str}


def verified_cv_text(text: str, person: Person) -> tuple[bool, str]:
    n = normalize(text)

    # Identità: nome+cognome oppure codice fiscale.
    identity = identity_score(text, person)
    if identity < 160:
        return False, "Identità non verificata nel PDF."

    # Se il CF è presente deve combaciare; è la prova migliore.
    if person.fiscal_code:
        cf_matches = re.findall(r"\b[A-Z]{6}\d{2}[A-Z]\d{2}[A-Z]\d{3}[A-Z]\b", text.upper())
        if cf_matches and person.fiscal_code not in cf_matches:
            return False, "Codice fiscale nel PDF diverso dal medico."

    # Se troviamo una DOB esplicita e differisce, scarta.
    if person.birth_date:
        m = re.search(
            r"(?:data\s+di\s+nascita|nato\s+il|nata\s+il).{0,50}"
            r"(\d{1,2}[./-]\d{1,2}[./-]\d{4})",
            text,
            flags=re.I | re.S,
        )
        if m:
            found = m.group(1).replace(".", "/").replace("-", "/")
            if found not in date_variants(person.birth_date):
                return False, "Data di nascita nel PDF incompatibile."

    cv_hits = sum(1 for term in CV_POSITIVE_TERMS if term in n)
    negative_hits = sum(1 for term in CV_NEGATIVE_TERMS if term in n)
    medical_hits = sum(1 for term in MEDICAL_CONTEXT_TERMS if term in n)

    if cv_hits < 2 and "curriculum vitae" not in n and "europass" not in n:
        return False, "Il PDF non ha una struttura da CV."

    if medical_hits == 0:
        return False, "Il PDF non mostra un contesto medico."

    if negative_hits >= 3 and "curriculum vitae" not in n:
        return False, "Documento amministrativo/concorsuale, non CV."

    return True, "CV verificato sul contenuto."


def cv_destination(person: Person, cv_dir: Path) -> Path:
    code = safe_filename_part(person.output_code)
    surname = safe_filename_part(person.surname)
    name = safe_filename_part(person.name)
    return cv_dir / f"{code}-{surname}-{name}.pdf"


def try_pdf_as_cv(
    url: str,
    person: Person,
    cv_dir: Path,
) -> tuple[str, str] | None:
    response_data = get_response(url, MAX_PDF_BYTES)
    if not response_data:
        return None
    response, raw = response_data
    try:
        ctype = response.headers.get("Content-Type", "").casefold()
        if not (raw.startswith(b"%PDF") or "application/pdf" in ctype):
            return None

        text = extract_pdf_text(raw)
        if not text:
            return None

        ok, reason = verified_cv_text(text, person)
        if not ok:
            logging.info("CV SCARTATO | %s | %s", url, reason)
            return None

        cv_dir.mkdir(parents=True, exist_ok=True)
        destination = cv_destination(person, cv_dir)
        destination.write_bytes(raw)

        logging.info("CV VERIFICATO | Pers_Id=%s | %s", person.pers_id, destination)
        return str(destination), text
    finally:
        response.close()


def try_landing_page_for_cv(
    result: SearchResult,
    person: Person,
    cv_dir: Path,
) -> tuple[str, str, str] | None:
    # Prima tenta come PDF diretto.
    direct = try_pdf_as_cv(result.url, person, cv_dir)
    if direct:
        path, text = direct
        return path, result.url, text

    response_data = get_response(result.url, MAX_HTML_BYTES)
    if not response_data:
        return None

    response, raw = response_data
    try:
        ctype = response.headers.get("Content-Type", "").casefold()
        if "html" not in ctype and not raw.lstrip().startswith((b"<", b"<!")):
            return None

        soup = BeautifulSoup(raw, "html.parser")
        candidates: list[tuple[int, str]] = []

        for a in soup.select("a[href]"):
            href = clean(a.get("href"))
            if not href:
                continue
            url = urljoin(result.url, href)
            label = normalize(f"{a.get_text(' ', strip=True)} {url}")

            score = 0
            if ".pdf" in url.casefold():
                score += 50
            if "curriculum" in label or "cv " in f"{label} " or "europass" in label:
                score += 80
            if normalize(person.surname) in label:
                score += 35
            if normalize(person.name) in label:
                score += 25
            if any(term in label for term in CV_NEGATIVE_TERMS):
                score -= 100

            if score >= 70:
                candidates.append((score, url))

        seen: set[str] = set()
        for _, url in sorted(candidates, reverse=True)[:10]:
            key = canonical_url(url)
            if key in seen:
                continue
            seen.add(key)
            found = try_pdf_as_cv(url, person, cv_dir)
            if found:
                path, text = found
                return path, url, text
    finally:
        response.close()

    return None


# ============================================================
# SPECIALTY EXTRACTION
# ============================================================

def specialty_matches(text: str) -> list[tuple[int, str]]:
    n = normalize(text)
    scores: dict[str, int] = {}

    # Espressioni esplicite pesano molto.
    explicit_patterns = (
        r"specializzat[oa]\s+in\s+([a-zà-ÿ0-9 '’/().,&+-]{3,100})",
        r"specializzazione\s+in\s+([a-zà-ÿ0-9 '’/().,&+-]{3,100})",
        r"specialista\s+in\s+([a-zà-ÿ0-9 '’/().,&+-]{3,100})",
        r"disciplina\s*[:\-]\s*([a-zà-ÿ0-9 '’/().,&+-]{3,100})",
    )

    explicit_chunks: list[str] = []
    for pattern in explicit_patterns:
        for m in re.finditer(pattern, n, flags=re.I):
            explicit_chunks.append(m.group(1)[:100])

    for canonical, aliases in SPECIALTY_ALIASES.items():
        best = 0

        for alias in aliases:
            a = normalize(alias)
            if not a:
                continue

            if any(a in chunk for chunk in explicit_chunks):
                best = max(best, 180)

            # Match generale.
            if a in n:
                best = max(best, 70)

            # Titoli/ruoli frequenti.
            if f"u.o. {a}" in n or f"unità operativa {a}" in n or f"unita operativa {a}" in n:
                best = max(best, 100)

        if best:
            scores[canonical] = best

    return sorted(((score, spec) for spec, score in scores.items()), reverse=True)


def choose_specialty_from_text(text: str) -> tuple[str, str]:
    matches = specialty_matches(text)
    if not matches:
        return "", ""

    top_score, top_spec = matches[0]

    # Se esiste un match esplicito forte, prendilo.
    if top_score >= 150:
        return top_spec, "alta"

    # Se il primo è molto sopra gli altri, medio-alta.
    if len(matches) == 1 or top_score >= matches[1][0] + 30:
        return top_spec, "media"

    # Più discipline presenti (tipico in pagine generiche): ambiguo.
    return "", ""


def specialty_from_search_results(
    results: list[SearchResult],
    person: Person,
) -> tuple[str, str, str]:
    candidates: list[tuple[int, str, str]] = []

    for result in results:
        if not relevant_result(result, person):
            continue

        text = f"{result.title} {result.snippet}"
        specialty, confidence = choose_specialty_from_text(text)
        if not specialty:
            continue

        score = rank_result(result, person)
        if confidence == "alta":
            score += 100
        else:
            score += 40

        candidates.append((score, specialty, result.url))

    if not candidates:
        return "", "", ""

    candidates.sort(reverse=True)
    best = candidates[0]

    # Conflitto forte fra fonti vicine.
    if len(candidates) > 1 and candidates[1][1] != best[1] and candidates[1][0] >= best[0] - 30:
        return "", "", ""

    confidence = "alta" if best[0] >= 450 else "media"
    return best[1], confidence, best[2]


# ============================================================
# AI OPTIONAL - SOLO EVIDENZE
# ============================================================

def ai_available(no_ai: bool) -> bool:
    return (
        not no_ai
        and genai is not None
        and bool(os.getenv("GEMINI_API_KEY"))
    )


def ai_extract_specialty(
    person: Person,
    evidence_text: str,
    model: str,
) -> tuple[str, str]:
    if not evidence_text.strip():
        return "", ""

    key = os.getenv("GEMINI_API_KEY")
    if not key or genai is None:
        return "", ""

    prompt = f"""
Devi estrarre la specialità medica di UNA specifica persona esclusivamente
dalle evidenze fornite sotto.

PERSONA:
Nome: {person.name}
Cognome: {person.surname}
Data di nascita: {person.birth_date or "N/D"}
Città database: {person.city or "N/D"}

REGOLE:
- Non usare conoscenza esterna.
- Non inventare.
- "Medico chirurgo" e "laurea in medicina e chirurgia" NON sono specialità.
- Se una specializzazione è esplicitamente dichiarata, restituiscila.
- Se il testo descrive solo un reparto/ruolo senza dimostrare la specialità,
  restituisci found=false.
- Se ci sono omonimi o discipline in conflitto, found=false.
- Restituisci JSON puro:
  {{"found": true/false, "specialty": "...", "confidence": "alta|media|bassa|nessuna"}}

EVIDENZE:
{evidence_text[:24000]}
""".strip()

    try:
        client = genai.Client(api_key=key)

        # API moderna preferita.
        if hasattr(client, "interactions"):
            interaction = client.interactions.create(
                model=model,
                input=prompt,
                response_format={
                    "type": "text",
                    "mime_type": "application/json",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "found": {"type": "boolean"},
                            "specialty": {"type": "string"},
                            "confidence": {
                                "type": "string",
                                "enum": ["alta", "media", "bassa", "nessuna"],
                            },
                        },
                        "required": ["found", "specialty", "confidence"],
                    },
                },
            )
            text = clean(getattr(interaction, "output_text", ""))
        else:
            # Fallback SDK.
            response = client.models.generate_content(
                model=model,
                contents=prompt,
                config={"response_mime_type": "application/json"},
            )
            text = clean(getattr(response, "text", ""))

        data = json.loads(text)
        if data.get("found") and clean(data.get("specialty")):
            return clean(data["specialty"]), clean(data.get("confidence")) or "media"
    except Exception as exc:
        logging.warning(
            "AI FAIL | Pers_Id=%s | %s: %s",
            person.pers_id, type(exc).__name__, exc
        )

    return "", ""


# ============================================================
# CACHE
# ============================================================

def person_cache_path(person: Person, cache_dir: Path) -> Path:
    identity = f"{person.pers_id}|{person.full_name}|{person.birth_date}|{VERSION}"
    digest = hashlib.sha1(identity.encode("utf-8")).hexdigest()[:12]
    return cache_dir / f"{safe_filename_part(person.pers_id)}_{digest}.json"


def result_from_cache(person: Person, cache_dir: Path) -> dict[str, Any] | None:
    path = person_cache_path(person, cache_dir)
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if data.get("version") != VERSION:
            return None
        return data.get("result")
    except Exception:
        return None


def save_result_cache(person: Person, cache_dir: Path, result: dict[str, Any]) -> None:
    atomic_json_write(
        person_cache_path(person, cache_dir),
        {
            "version": VERSION,
            "person": asdict(person),
            "result": result,
            "saved_at": utc_now(),
        },
    )


# ============================================================
# PERSON PIPELINE
# ============================================================

def research_person(
    person: Person,
    cv_dir: Path,
    cache_dir: Path,
    max_results: int,
    backend: str,
    use_ai: bool,
    ai_model: str,
) -> dict[str, Any]:
    cached = result_from_cache(person, cache_dir)
    if cached:
        cached["from_cache"] = True
        return cached

    started = time.perf_counter()
    notes: list[str] = []
    sources: list[str] = []
    all_results: list[SearchResult] = []

    logging.info(
        "START | row=%s | Pers_Id=%s | Medico_Id=%s | %s",
        person.row, person.pers_id, person.medico_id, person.full_name
    )

    # --------------------------------------------------------
    # FASE 1: CV FIRST
    # --------------------------------------------------------
    cv_path = ""
    cv_url = ""
    cv_text = ""

    for query in build_cv_queries(person):
        batch = run_search(query, max_results, backend)
        batch = [
            SearchResult(
                title=r.title, url=r.url, snippet=r.snippet, query=r.query,
                provider=r.provider, category="cv"
            )
            for r in batch
        ]

        relevant = [r for r in batch if relevant_result(r, person)]
        relevant.sort(key=lambda r: rank_result(r, person), reverse=True)

        all_results.extend(relevant)

        for result in relevant[:4]:
            found = try_landing_page_for_cv(result, person, cv_dir)
            if found:
                cv_path, cv_url, cv_text = found
                sources.append(cv_url)
                notes.append("CV ufficiale/professionale verificato sul contenuto.")
                break

        if cv_path:
            break

    # --------------------------------------------------------
    # FASE 2: SPECIALITA DAL CV
    # --------------------------------------------------------
    specialty = ""
    specialty_confidence = ""
    specialty_source = ""

    if cv_text:
        specialty, specialty_confidence = choose_specialty_from_text(cv_text)
        if specialty:
            specialty_source = cv_url
            notes.append("Specialità ricavata dal CV verificato.")

    # --------------------------------------------------------
    # FASE 3: RICERCA SPECIALITA SE SERVE
    # --------------------------------------------------------
    if not specialty:
        for query in build_specialty_queries(person):
            batch = run_search(query, max_results, backend)
            batch = [
                SearchResult(
                    title=r.title, url=r.url, snippet=r.snippet, query=r.query,
                    provider=r.provider, category="specialty"
                )
                for r in batch
            ]

            relevant = [r for r in batch if relevant_result(r, person)]
            relevant.sort(key=lambda r: rank_result(r, person), reverse=True)
            all_results.extend(relevant)

            # Early stop appena troviamo una specialità non ambigua.
            found_spec, found_conf, found_source = specialty_from_search_results(
                dedupe_results(all_results), person
            )
            if found_spec:
                specialty = found_spec
                specialty_confidence = found_conf
                specialty_source = found_source
                sources.append(found_source)
                notes.append("Specialità ricavata da fonte web pertinente.")
                break

    # --------------------------------------------------------
    # FASE 4: AI SOLO SE EVIDENZA ESISTE MA È AMBIGUA
    # --------------------------------------------------------
    if not specialty and use_ai:
        ranked = sorted(
            dedupe_results(all_results),
            key=lambda r: rank_result(r, person),
            reverse=True,
        )[:8]

        evidence_blocks = []
        if cv_text:
            evidence_blocks.append("[CV VERIFICATO]\n" + cv_text[:16000])

        for i, r in enumerate(ranked, 1):
            evidence_blocks.append(
                f"[FONTE {i}]\nTitolo: {r.title}\nURL: {r.url}\nSnippet: {r.snippet}"
            )

        ai_spec, ai_conf = ai_extract_specialty(
            person,
            "\n\n".join(evidence_blocks),
            ai_model,
        )
        if ai_spec:
            specialty = ai_spec
            specialty_confidence = ai_conf or "media"
            specialty_source = "AI_SU_EVIDENZE"
            notes.append("Specialità estratta dall'AI esclusivamente dalle evidenze raccolte.")

    # --------------------------------------------------------
    # FONTI / STATO
    # --------------------------------------------------------
    ranked_sources = sorted(
        dedupe_results(all_results),
        key=lambda r: rank_result(r, person),
        reverse=True,
    )

    if specialty_source and specialty_source.startswith("http"):
        sources.append(specialty_source)

    for r in ranked_sources[:5]:
        if r.url not in sources:
            sources.append(r.url)

    sources = list(dict.fromkeys(sources))

    if specialty and cv_path:
        status = "COMPLETATO"
    elif specialty:
        status = "SPECIALITA_TROVATA"
    elif cv_path:
        status = "CV_TROVATO_SPECIALITA_DA_VERIFICARE"
    elif all_results:
        status = "DA_VERIFICARE"
    else:
        status = "NESSUN_RISULTATO"

    elapsed = time.perf_counter() - started

    result = {
        "row": person.row,
        "status": status,
        "specialty": specialty,
        "specialty_confidence": specialty_confidence or ("nessuna" if not specialty else "media"),
        "cv_path": cv_path,
        "cv_url": cv_url,
        "sources": "\n".join(sources),
        "notes": " ".join(notes),
        "updated": utc_now(),
        "error": "",
        "elapsed": elapsed,
        "from_cache": False,
    }

    save_result_cache(person, cache_dir, result)

    logging.info(
        "DONE | Pers_Id=%s | %.1fs | stato=%s | specialita=%s | cv=%s",
        person.pers_id,
        elapsed,
        status,
        specialty or "N/D",
        "SI" if cv_path else "NO",
    )
    return result


# ============================================================
# WRITE RESULT
# ============================================================

def write_result(ws, headers: dict[str, int], result: dict[str, Any]) -> None:
    row = int(result["row"])
    values = {
        "Ricerca_Stato": result.get("status", ""),
        "Specialita": result.get("specialty", ""),
        "Specialita_Confidenza": result.get("specialty_confidence", ""),
        "CV_Salvato": result.get("cv_path", ""),
        "CV_URL": result.get("cv_url", ""),
        "Fonti_Ricerca": result.get("sources", ""),
        "Ricerca_Note": result.get("notes", ""),
        "Ultimo_Aggiornamento_UTC": result.get("updated", utc_now()),
        "Ricerca_Errore": result.get("error", ""),
    }
    for label, value in values.items():
        set_cell(ws, row, headers, label, value)


# ============================================================
# MAIN
# ============================================================

def main() -> int:
    global SEARCH_SEMAPHORE

    args = parse_args()

    SEARCH_SEMAPHORE = threading.BoundedSemaphore(args.search_concurrency)

    log_file = configure_logging(args.log_dir)

    input_path = args.input.expanduser().resolve()
    output_path = output_path_for(input_path, args.output).expanduser().resolve()
    cv_dir = args.cv_dir.expanduser().resolve()
    cache_dir = args.cache_dir.expanduser().resolve()

    cv_dir.mkdir(parents=True, exist_ok=True)
    cache_dir.mkdir(parents=True, exist_ok=True)

    prepare_output(input_path, output_path)

    wb = load_workbook(output_path)
    if args.sheet not in wb.sheetnames:
        raise ValueError(
            f"Foglio '{args.sheet}' non trovato. Disponibili: {wb.sheetnames}"
        )

    ws = wb[args.sheet]
    headers = normalized_headers(ws)
    require_columns(headers)
    headers = ensure_output_columns(ws)
    atomic_save(wb, output_path)

    use_ai = ai_available(args.no_ai)

    logging.info("=" * 68)
    logging.info("%s", VERSION)
    logging.info("Input: %s", input_path)
    logging.info("Output: %s", output_path)
    logging.info("Foglio: %s", args.sheet)
    logging.info("CV dir: %s", cv_dir)
    logging.info("Cache dir: %s", cache_dir)
    logging.info("Workers: %s", args.workers)
    logging.info("Search concurrency: %s", args.search_concurrency)
    logging.info("Search backend fallback: %s", args.search_backend)
    logging.info("Serper: %s", "SI" if os.getenv("SERPER_API_KEY") else "NO")
    logging.info("Brave API: %s", "SI" if os.getenv("BRAVE_SEARCH_API_KEY") else "NO")
    logging.info("Gemini AI: %s", "SI" if use_ai else "NO")
    if use_ai:
        logging.info("Gemini model: %s", args.ai_model)
    logging.info("Log: %s", log_file)
    logging.info("=" * 68)

    people: list[Person] = []

    for row in range(max(2, args.start_row), ws.max_row + 1):
        if args.limit is not None and len(people) >= args.limit:
            break

        person = person_from_row(ws, row, headers)
        if not person.pers_id or not person.surname or not person.name:
            continue

        current_status = clean(getv(ws, row, headers, "Ricerca_Stato")).upper()

        if not args.retry_completed and current_status in {
            "COMPLETATO",
            "SPECIALITA_TROVATA",
        }:
            continue

        if current_status == "ERRORE" and not args.retry_errors:
            continue

        people.append(person)

    if not people:
        logging.info("Nessun medico da elaborare.")
        return 0

    total = len(people)
    started_all = time.perf_counter()
    finished = 0
    pending_save = 0
    counters: dict[str, int] = {}

    logging.info("Medici da elaborare: %s", total)

    with ThreadPoolExecutor(
        max_workers=args.workers,
        thread_name_prefix="medico",
    ) as executor:
        future_map = {
            executor.submit(
                research_person,
                person,
                cv_dir,
                cache_dir,
                args.search_results,
                args.search_backend,
                use_ai,
                args.ai_model,
            ): person
            for person in people
        }

        for future in as_completed(future_map):
            person = future_map[future]

            try:
                result = future.result()
            except Exception as exc:
                logging.exception(
                    "ERROR | row=%s | Pers_Id=%s | %s",
                    person.row, person.pers_id, person.full_name
                )
                result = {
                    "row": person.row,
                    "status": "ERRORE",
                    "specialty": "",
                    "specialty_confidence": "nessuna",
                    "cv_path": "",
                    "cv_url": "",
                    "sources": "",
                    "notes": "",
                    "updated": utc_now(),
                    "error": f"{type(exc).__name__}: {exc}",
                    "elapsed": 0.0,
                }

            write_result(ws, headers, result)

            status = clean(result.get("status")) or "?"
            counters[status] = counters.get(status, 0) + 1

            finished += 1
            pending_save += 1

            if pending_save >= args.save_every:
                atomic_save(wb, output_path)
                pending_save = 0
                logging.info("CHECKPOINT | %s/%s", finished, total)

            if finished % 10 == 0 or finished == total:
                elapsed = time.perf_counter() - started_all
                rate = finished / elapsed * 60 if elapsed else 0.0
                eta = (total - finished) / rate if rate else 0.0
                logging.info(
                    "PROGRESS | %s/%s | %.1f medici/min | ETA %.1f min | %s",
                    finished, total, rate, eta, counters
                )

    atomic_save(wb, output_path)

    elapsed = time.perf_counter() - started_all
    rate = finished / elapsed * 60 if elapsed else 0.0

    logging.info("=" * 68)
    logging.info("FINE %s", VERSION)
    logging.info("Elaborati: %s", finished)
    logging.info("Tempo: %.1f min", elapsed / 60)
    logging.info("Velocita media: %.1f medici/min", rate)
    logging.info("Stati: %s", counters)
    logging.info("Output: %s", output_path)
    logging.info("CV: %s", cv_dir)
    logging.info("Cache: %s", cache_dir)
    logging.info("=" * 68)

    return 0


if __name__ == "__main__":
    import sys

    early_input, early_output, early_log_dir = parse_early_paths(sys.argv)
    startup_log = bootstrap_log_path(early_log_dir)

    try:
        initial_output = create_initial_output_copy(early_input, early_output)
        if initial_output:
            write_bootstrap_log(
                startup_log,
                f"INFO | Copia iniziale Excel disponibile: {initial_output}"
            )
    except Exception as exc:
        write_bootstrap_log(
            startup_log,
            f"ERRORE | Impossibile creare la copia iniziale Excel: "
            f"{type(exc).__name__}: {exc}"
        )

    ok, missing = import_dependencies(startup_log)
    if not ok:
        print("")
        print("ERRORE: mancano dipendenze obbligatorie.")
        print("Installa con:")
        print("  python -m pip install -r requirements_medici.txt")
        print("")
        print(f"Log diagnostico: {startup_log.resolve()}")
        if early_input:
            resolved_out = early_output or early_input.with_name(
                f"{early_input.stem}_specialita_cv.xlsx"
            )
            print(f"Excel di output/copia iniziale: {Path(resolved_out).resolve()}")
        print("")
        print("Dipendenze mancanti:")
        for item in missing:
            print(f" - {item}")
        raise SystemExit(2)

    try:
        # Ora che python-dotenv è disponibile.
        load_dotenv()
        raise SystemExit(main())
    except KeyboardInterrupt:
        write_bootstrap_log(startup_log, "INTERRUZIONE | KeyboardInterrupt")
        print("\\nInterrotto.")
        raise SystemExit(130)
    except Exception as exc:
        write_bootstrap_log(
            startup_log,
            f"FATAL | {type(exc).__name__}: {exc}"
        )
        print("")
        print(f"ERRORE FATALE: {type(exc).__name__}: {exc}")
        print(f"Log diagnostico: {startup_log.resolve()}")
        raise

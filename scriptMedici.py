from __future__ import annotations

import argparse
import base64
import hashlib
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
from dataclasses import dataclass, asdict
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import quote_plus, urljoin, urlparse, parse_qs

VERSION = "V3.4 MEDICI - RELEVANCE FALLBACK + GOOGLE FIRST"

# Import caricati dopo il bootstrap, così i log esistono anche se manca una dipendenza.
requests = None
BeautifulSoup = None
load_dotenv = None
load_workbook = None
Alignment = Font = PatternFill = None
PdfReader = None
sync_playwright = None

DEFAULT_SHEET = "Foglio1"
DEFAULT_ENGINE = "auto"
DEFAULT_SEARCH_RESULTS = 8
DEFAULT_SAVE_EVERY = 10
DEFAULT_DELAY_MIN = 1.2
DEFAULT_DELAY_MAX = 2.8

HTTP_TIMEOUT = 12
MAX_PDF_BYTES = 25 * 1024 * 1024
MAX_HTML_BYTES = 5 * 1024 * 1024
MAX_PDF_PAGES = 100
MAX_PDF_TEXT = 120_000
MAX_PAGE_TEXT = 70_000
MAX_LANDING_PAGES_PER_PERSON = 8
MAX_PDF_CANDIDATES_PER_PERSON = 16
MAX_PDF_LINKS_PER_LANDING = 25

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/151.0.0.0 Safari/537.36"
)

OUTPUT_COLUMNS = (
    "Ricerca_Stato",
    "Specialita",
    "Specialita_Confidenza",
    "Specialita_Evidenza",
    "CV_Salvato",
    "CV_URL",
    "CV_Confidenza",
    "CV_Da_Verificare",
    "CV_Da_Verificare_URL",
    "Fonti_Ricerca",
    "Ricerca_Metodo",
    "Ricerca_Note",
    "Ultimo_Aggiornamento_UTC",
    "Ricerca_Errore",
)


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
        mid = numericish(clean(self.medico_id))
        if mid and mid not in {"0", "0.0"}:
            return mid
        return numericish(clean(self.pers_id))


@dataclass(frozen=True)
class WebHit:
    title: str
    url: str
    snippet: str
    engine: str


# ============================================================
# UTIL / BOOTSTRAP
# ============================================================

def clean(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if s.casefold() in {"none", "null", "nan", "n/a", "n.d.", "nd"}:
        return ""
    return s


def numericish(value: str) -> str:
    s = clean(value)
    if re.fullmatch(r"\d+\.0", s):
        return s[:-2]
    return s


def normalize(text: str) -> str:
    s = clean(text).casefold()
    s = s.replace("’", "'").replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", s).strip()


def safe_part(value: str) -> str:
    s = re.sub(r'[<>:"/\\|?*]', "_", clean(value))
    s = re.sub(r"\s+", "_", s).strip(" ._")
    return s or "ND"


def canonical_url(url: str) -> str:
    return clean(url).split("#", 1)[0].rstrip("/")


def decode_bing_redirect(url: str) -> str:
    """
    Decodifica i redirect Bing del tipo /ck/a?...&u=a1BASE64...
    Restituisce l'URL reale quando possibile.
    """
    raw = clean(url)
    try:
        p = urlparse(raw)
        host = (p.hostname or "").casefold()
        if "bing.com" not in host:
            return raw

        qs = parse_qs(p.query)
        values = qs.get("u", [])
        if not values:
            return raw

        value = clean(values[0])
        # Bing usa spesso prefisso a1 + Base64 URL-safe.
        payload = value[2:] if value.startswith("a1") else value
        padding = "=" * (-len(payload) % 4)

        try:
            decoded = base64.urlsafe_b64decode(payload + padding).decode("utf-8", errors="ignore")
            if decoded.startswith(("http://", "https://")):
                return decoded
        except Exception:
            pass

        # In alcuni casi il valore è già percent-encoded o contiene direttamente URL.
        if value.startswith(("http://", "https://")):
            return value
    except Exception:
        pass

    return raw


def normalize_result_url(url: str) -> str:
    u = decode_bing_redirect(url)
    return canonical_url(u)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def domain(url: str) -> str:
    try:
        return (urlparse(url).hostname or "").casefold()
    except Exception:
        return ""


def bootstrap_log_path(log_dir: Path) -> Path:
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir / f"startup_medici_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"


def write_bootstrap_log(path: Path, message: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8") as f:
        f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {message}\n")


def parse_early_paths(argv: list[str]) -> tuple[Path | None, Path | None, Path]:
    input_path = None
    output_path = None
    log_dir = Path("logs")

    for arg in argv[1:]:
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


def default_output_path(input_path: Path) -> Path:
    return Path.cwd() / "output" / f"{input_path.stem}_specialita_cv_v3_4.xlsx"


def create_initial_output_copy(input_path: Path | None, output_path: Path | None) -> Path | None:
    if input_path is None:
        return None
    inp = input_path.expanduser().resolve()
    out = output_path.expanduser().resolve() if output_path else default_output_path(inp).resolve()

    if inp.exists() and inp.suffix.casefold() == ".xlsx":
        out.parent.mkdir(parents=True, exist_ok=True)
        if not out.exists() and inp != out:
            shutil.copy2(inp, out)
    return out


def import_dependencies(log_path: Path) -> tuple[bool, list[str]]:
    global requests, BeautifulSoup, load_dotenv
    global load_workbook, Alignment, Font, PatternFill, PdfReader, sync_playwright

    missing = []

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
        from dotenv import load_dotenv as _load_dotenv
        load_dotenv = _load_dotenv
    except Exception as exc:
        missing.append(f"python-dotenv ({exc})")

    try:
        from openpyxl import load_workbook as _load_workbook
        from openpyxl.styles import Alignment as _Alignment, Font as _Font, PatternFill as _PatternFill
        load_workbook = _load_workbook
        Alignment, Font, PatternFill = _Alignment, _Font, _PatternFill
    except Exception as exc:
        missing.append(f"openpyxl ({exc})")

    try:
        from pypdf import PdfReader as _PdfReader
        PdfReader = _PdfReader
    except Exception as exc:
        missing.append(f"pypdf ({exc})")

    try:
        from playwright.sync_api import sync_playwright as _sync_playwright
        sync_playwright = _sync_playwright
    except Exception as exc:
        missing.append(f"playwright ({exc})")

    if missing:
        write_bootstrap_log(log_path, "ERRORE | Dipendenze mancanti: " + "; ".join(missing))
        return False, missing

    write_bootstrap_log(log_path, "OK | Dipendenze caricate.")
    return True, []


# ============================================================
# CLI / LOG
# ============================================================

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description=VERSION)
    p.add_argument("input", type=Path)
    p.add_argument("--output", type=Path)
    p.add_argument("--sheet", default=DEFAULT_SHEET)
    p.add_argument("--limit", "--max-rows", dest="limit", type=int)
    p.add_argument("--start-row", type=int, default=2)

    p.add_argument("--engine", choices=("bing", "google", "auto"), default=DEFAULT_ENGINE,
                   help="Motore preferito; l'altro viene usato automaticamente come fallback.")
    p.add_argument("--search-results", type=int, default=DEFAULT_SEARCH_RESULTS)
    p.add_argument("--headless", action="store_true",
                   help="Avvia Chromium senza finestra. Per i primi test consiglio di NON usarlo.")
    p.add_argument("--deep", action="store_true",
                   help="Aggiunge query e pagine da analizzare per i casi incompleti.")
    p.add_argument("--save-every", type=int, default=DEFAULT_SAVE_EVERY)
    p.add_argument("--delay-min", type=float, default=DEFAULT_DELAY_MIN)
    p.add_argument("--delay-max", type=float, default=DEFAULT_DELAY_MAX)

    p.add_argument("--retry-all", action="store_true")
    p.add_argument("--retry-errors", action="store_true")
    p.add_argument("--ignore-cache", action="store_true")

    p.add_argument("--cv-dir", type=Path, default=Path("cv_medici"))
    p.add_argument("--cv-review-dir", type=Path, default=Path("cv_medici_da_verificare"))
    p.add_argument("--cache-dir", type=Path, default=Path("cache_medici"))
    p.add_argument("--browser-data-dir", type=Path, default=Path("browser_medici"))
    p.add_argument("--log-dir", type=Path, default=Path("logs"))

    args = p.parse_args()

    if args.limit is not None and args.limit <= 0:
        p.error("--limit deve essere > 0")
    if args.start_row < 2:
        p.error("--start-row deve essere >= 2")
    if args.search_results <= 0:
        p.error("--search-results deve essere > 0")
    if args.save_every <= 0:
        p.error("--save-every deve essere > 0")
    if args.delay_min < 0 or args.delay_max < args.delay_min:
        p.error("Delay non valido")
    return args


def configure_logging(log_dir: Path) -> Path:
    log_dir.mkdir(parents=True, exist_ok=True)
    path = log_dir / f"run_medici_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

    root = logging.getLogger()
    root.handlers.clear()
    root.setLevel(logging.INFO)

    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    fh = logging.FileHandler(path, encoding="utf-8")
    sh = logging.StreamHandler()
    fh.setFormatter(fmt)
    sh.setFormatter(fmt)
    root.addHandler(fh)
    root.addHandler(sh)
    return path


# ============================================================
# EXCEL
# ============================================================

def headers_map(ws) -> dict[str, int]:
    out = {}
    for c in ws[1]:
        label = clean(c.value)
        if label:
            out[label.casefold()] = c.column
    return out


def require_input_columns(headers: dict[str, int]) -> None:
    req = ("Pers_Id", "Pers_Cognome", "Pers_Nome")
    missing = [x for x in req if x.casefold() not in headers]
    if missing:
        raise ValueError("Colonne obbligatorie mancanti: " + ", ".join(missing))


def ensure_output_columns(ws) -> dict[str, int]:
    headers = headers_map(ws)
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


def getv(ws, row: int, headers: dict[str, int], label: str) -> Any:
    col = headers.get(label.casefold())
    return ws.cell(row, col).value if col else ""


def person_from_row(ws, row: int, headers: dict[str, int]) -> Person:
    dob = getv(ws, row, headers, "Pers_DataNascita")
    if isinstance(dob, (datetime, date)):
        dob = dob.strftime("%d/%m/%Y")

    return Person(
        row=row,
        pers_id=numericish(clean(getv(ws, row, headers, "Pers_Id"))),
        medico_id=numericish(clean(getv(ws, row, headers, "Medico_Id"))),
        surname=clean(getv(ws, row, headers, "Pers_Cognome")),
        name=clean(getv(ws, row, headers, "Pers_Nome")),
        birth_date=clean(dob),
        fiscal_code=clean(getv(ws, row, headers, "Pers_CodFis")).upper(),
        city=clean(getv(ws, row, headers, "Indirizzi_Citta")),
        email=clean(getv(ws, row, headers, "emailPredefinita")) or
              clean(getv(ws, row, headers, "Email")),
    )


def prepare_output(input_path: Path, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if not output_path.exists():
        shutil.copy2(input_path, output_path)


def atomic_save(wb, path: Path) -> None:
    fd, tmp_name = tempfile.mkstemp(prefix=f".{path.stem}_", suffix=".xlsx", dir=path.parent)
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    finally:
        tmp.unlink(missing_ok=True)


def set_result(ws, headers: dict[str, int], row: int, result: dict[str, Any]) -> None:
    mapping = {
        "Ricerca_Stato": result.get("status", ""),
        "Specialita": result.get("specialty", ""),
        "Specialita_Confidenza": result.get("specialty_confidence", ""),
        "Specialita_Evidenza": result.get("specialty_evidence", ""),
        "CV_Salvato": result.get("cv_path", ""),
        "CV_URL": result.get("cv_url", ""),
        "CV_Confidenza": result.get("cv_confidence", ""),
        "CV_Da_Verificare": result.get("cv_review_paths", ""),
        "CV_Da_Verificare_URL": result.get("cv_review_urls", ""),
        "Fonti_Ricerca": result.get("sources", ""),
        "Ricerca_Metodo": result.get("method", ""),
        "Ricerca_Note": result.get("notes", ""),
        "Ultimo_Aggiornamento_UTC": result.get("updated", utc_now()),
        "Ricerca_Errore": result.get("error", ""),
    }
    for label, value in mapping.items():
        ws.cell(row, headers[label.casefold()], value)


# ============================================================
# CACHE
# ============================================================

def cache_path(person: Person, cache_dir: Path) -> Path:
    identity = f"{VERSION}|{person.pers_id}|{person.full_name}|{person.birth_date}|{person.fiscal_code}"
    digest = hashlib.sha1(identity.encode("utf-8")).hexdigest()[:14]
    return cache_dir / f"{safe_part(person.pers_id)}_{digest}.json"


def read_cache(person: Person, cache_dir: Path) -> dict[str, Any] | None:
    p = cache_path(person, cache_dir)
    try:
        if p.exists():
            data = json.loads(p.read_text(encoding="utf-8"))
            if data.get("version") == VERSION:
                return data.get("result")
    except Exception:
        pass
    return None


def write_cache(person: Person, cache_dir: Path, result: dict[str, Any]) -> None:
    cache_dir.mkdir(parents=True, exist_ok=True)
    p = cache_path(person, cache_dir)
    payload = {"version": VERSION, "person": asdict(person), "result": result}
    fd, tmp_name = tempfile.mkstemp(prefix=".cache_", suffix=".json", dir=cache_dir)
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, p)
    finally:
        tmp.unlink(missing_ok=True)


# ============================================================
# BROWSER SEARCH
# ============================================================

BLOCK_MARKERS = (
    "unusual traffic", "our systems have detected unusual traffic",
    "verify you are human", "captcha", "detected unusual traffic",
    "access denied", "robot or human", "challenge",
)


class BrowserSearch:
    def __init__(self, data_dir: Path, headless: bool, engine: str,
                 delay_min: float, delay_max: float, max_results: int):
        self.data_dir = data_dir
        self.headless = headless
        self.engine = engine
        self.delay_min = delay_min
        self.delay_max = delay_max
        self.max_results = max_results
        self.pw = None
        self.context = None
        self.page = None
        self.blocked_engines: set[str] = set()

    def __enter__(self):
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.pw = sync_playwright().start()

        self.context = self.pw.chromium.launch_persistent_context(
            user_data_dir=str(self.data_dir),
            headless=self.headless,
            viewport={"width": 1440, "height": 900},
            user_agent=USER_AGENT,
            locale="it-IT",
            timezone_id="Europe/Rome",
            args=["--disable-blink-features=AutomationControlled"],
        )
        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()
        self.page.set_default_timeout(12_000)
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if self.context:
                self.context.close()
        finally:
            if self.pw:
                self.pw.stop()

    def _delay(self):
        time.sleep(random.uniform(self.delay_min, self.delay_max))

    def _blocked(self) -> bool:
        try:
            body = normalize(self.page.locator("body").inner_text(timeout=3000))
            return any(x in body for x in BLOCK_MARKERS)
        except Exception:
            return False

    def search(self, query: str) -> list[WebHit]:
        # AUTO: Google prima. Bing resta fallback tecnico.
        if self.engine in {"google", "auto"}:
            engines = ["google", "bing"]
        else:
            engines = ["bing", "google"]

        for engine in engines:
            if engine in self.blocked_engines:
                continue
            try:
                hits = self._search_engine(engine, query)
                if hits:
                    return hits
                logging.warning("SEARCH EMPTY | %s | %s", engine, query)
            except RuntimeError as exc:
                if "BLOCCO_BROWSER" in str(exc):
                    self.blocked_engines.add(engine)
                    logging.error("BROWSER BLOCCATO | %s | %s", engine, exc)
                else:
                    logging.warning("SEARCH FAIL | %s | %s | %s", engine, query, exc)
            except Exception as exc:
                logging.warning("SEARCH FAIL | %s | %s | %s: %s",
                                engine, query, type(exc).__name__, exc)
        return []

    def search_specific(self, engine: str, query: str) -> list[WebHit]:
        """
        Ricerca esplicita su un singolo motore. Serve quando il motore primario
        restituisce risultati tecnicamente validi ma semanticamente irrilevanti.
        """
        if engine in self.blocked_engines:
            return []
        try:
            return self._search_engine(engine, query)
        except RuntimeError as exc:
            if "BLOCCO_BROWSER" in str(exc):
                self.blocked_engines.add(engine)
                logging.error("BROWSER BLOCCATO | %s | %s", engine, exc)
            else:
                logging.warning("SEARCH FAIL | %s | %s | %s", engine, query, exc)
        except Exception as exc:
            logging.warning(
                "SEARCH FAIL | %s | %s | %s: %s",
                engine, query, type(exc).__name__, exc
            )
        return []

    def _search_engine(self, engine: str, query: str) -> list[WebHit]:
        self._delay()

        if engine == "bing":
            url = f"https://www.bing.com/search?q={quote_plus(query)}&setlang=it-IT&cc=it"
        else:
            url = f"https://www.google.com/search?q={quote_plus(query)}&hl=it&gl=it&num=10"

        self.page.goto(url, wait_until="domcontentloaded", timeout=20_000)

        if self._blocked():
            raise RuntimeError(f"BLOCCO_BROWSER: CAPTCHA/anti-bot rilevato su {engine}")

        hits = self._parse_bing() if engine == "bing" else self._parse_google()

        logging.info("SEARCH OK | %s | %s | %s risultati", engine, query, len(hits))
        return hits[:self.max_results]

    def _parse_bing(self) -> list[WebHit]:
        out = []
        for li in self.page.locator("li.b_algo").all()[:20]:
            try:
                a = li.locator("h2 a").first
                title = clean(a.inner_text(timeout=1500))
                url = clean(a.get_attribute("href"))
                url = normalize_result_url(url)
                snippet = ""
                p = li.locator(".b_caption p")
                if p.count():
                    snippet = clean(p.first.inner_text(timeout=1000))
                if url.startswith(("http://", "https://")):
                    out.append(WebHit(title, url, snippet, "bing"))
            except Exception:
                continue
        return dedupe_hits(out)

    def _parse_google(self) -> list[WebHit]:
        out = []
        # Google cambia spesso markup: partiamo dagli h3 e risaliamo all'anchor.
        for h3 in self.page.locator("#search h3").all()[:30]:
            try:
                a = h3.locator("xpath=ancestor::a[1]")
                if not a.count():
                    continue
                title = clean(h3.inner_text(timeout=1000))
                url = clean(a.get_attribute("href"))
                if url.startswith("/url?"):
                    q = parse_qs(urlparse(url).query).get("q", [])
                    url = q[0] if q else ""
                if not url.startswith(("http://", "https://")):
                    continue
                container = h3.locator("xpath=ancestor::div[contains(@class,'MjjYud')][1]")
                snippet = ""
                try:
                    snippet = clean(container.inner_text(timeout=1000))[:900]
                except Exception:
                    pass
                out.append(WebHit(title, url, snippet, "google"))
            except Exception:
                continue
        return dedupe_hits(out)


def dedupe_hits(hits: Iterable[WebHit]) -> list[WebHit]:
    out = []
    seen = set()
    for hit in hits:
        real_url = normalize_result_url(hit.url)
        key = canonical_url(real_url)
        if key and key not in seen:
            seen.add(key)
            out.append(WebHit(hit.title, real_url, hit.snippet, hit.engine))
    return out


# ============================================================
# HTTP / PDF / PAGINE
# ============================================================

def is_public_http_url(url: str) -> bool:
    try:
        p = urlparse(url)
        if p.scheme not in {"http", "https"} or not p.hostname:
            return False

        try:
            infos = socket.getaddrinfo(
                p.hostname,
                p.port or (443 if p.scheme == "https" else 80),
                type=socket.SOCK_STREAM,
            )
            for info in infos:
                ip = ipaddress.ip_address(info[4][0])
                if ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved or ip.is_multicast:
                    return False
        except OSError:
            pass
        return True
    except Exception:
        return False


def get_bytes(url: str, max_bytes: int) -> tuple[Any, bytes] | None:
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
        buf = bytearray()
        for chunk in r.iter_content(65536):
            if chunk:
                buf.extend(chunk)
                if len(buf) > max_bytes:
                    r.close()
                    return None
        return r, bytes(buf)
    except Exception:
        return None


def pdf_text(raw: bytes) -> str:
    import io
    try:
        reader = PdfReader(io.BytesIO(raw))
        chunks = []
        for page in reader.pages[:MAX_PDF_PAGES]:
            t = page.extract_text() or ""
            if t:
                chunks.append(t)
        return "\n".join(chunks)[:MAX_PDF_TEXT]
    except Exception:
        return ""


def html_text(raw: bytes) -> str:
    try:
        soup = BeautifulSoup(raw, "html.parser")
        for tag in soup(["script", "style", "noscript", "svg"]):
            tag.decompose()
        return clean(soup.get_text(" ", strip=True))[:MAX_PAGE_TEXT]
    except Exception:
        return ""


def page_text(url: str) -> str:
    got = get_bytes(url, MAX_HTML_BYTES)
    if not got:
        return ""
    r, raw = got
    try:
        ctype = normalize(r.headers.get("Content-Type", ""))
        if raw.startswith(b"%PDF") or "application/pdf" in ctype:
            return pdf_text(raw)
        return html_text(raw)
    finally:
        r.close()


# ============================================================
# IDENTITÀ / CV
# ============================================================

CV_POSITIVE = (
    "curriculum vitae", "curriculum professionale", "curriculum formativo",
    "europass", "esperienza professionale", "esperienze professionali",
    "istruzione e formazione", "titoli di studio", "attività professionale",
)

CV_NEGATIVE = (
    "graduatoria", "elenco candidati", "candidati ammessi", "verbale",
    "bando", "avviso pubblico", "prova orale", "prova scritta",
    "commissione esaminatrice", "delibera", "deliberazione",
)

MEDICAL_TERMS = (
    "medico", "medicina", "chirurgo", "specialista", "specializzazione",
    "ospedale", "azienda sanitaria", "asl", "ausl", "asst", "aou",
    "irccs", "policlinico", "ordine dei medici",
)


TRUSTED_MEDICAL_DOMAIN_HINTS = (
    "asl", "ausl", "asst", "ats", "aou", "ao-", "osped", "policlin",
    "irccs", "sanita", "salute", "regione", "univ", "universita",
    "ordinemedici", "fnomceo", "gov.it",
)

JUNK_DOMAIN_HINTS = (
    "amazon.", "reddit.", "zhihu.", "baidu.", "qiwa.", "spartex.",
    "facebook.", "instagram.", "tiktok.", "pinterest.", "youtube.",
    "linkedin.", "wikipedia.", "ebay.", "aliexpress.",
)


def identity_score(text: str, person: Person) -> int:
    n = normalize(text)
    full = normalize(person.full_name)
    reverse = normalize(f"{person.surname} {person.name}")
    score = 0
    if full and full in n:
        score += 180
    if reverse and reverse in n:
        score += 180
    if normalize(person.surname) in n:
        score += 50
    if normalize(person.name) in n:
        score += 35
    if person.fiscal_code and normalize(person.fiscal_code) in n:
        score += 600
    if person.city and normalize(person.city) in n:
        score += 35
    return score


def date_variants(dob: str) -> set[str]:
    if not dob:
        return set()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
        try:
            d = datetime.strptime(dob[:10], fmt)
            return {
                d.strftime("%d/%m/%Y"), d.strftime("%d-%m-%Y"),
                d.strftime("%d.%m.%Y"), d.strftime("%Y-%m-%d"),
            }
        except Exception:
            continue
    return {dob}


def verify_cv(text: str, person: Person) -> tuple[bool, str, str]:
    n = normalize(text)

    if identity_score(text, person) < 180:
        return False, "bassa", "Nome/cognome non verificati."

    if person.fiscal_code:
        fiscal_codes = re.findall(r"\b[A-Z]{6}\d{2}[A-Z]\d{2}[A-Z]\d{3}[A-Z]\b", text.upper())
        if fiscal_codes and person.fiscal_code not in fiscal_codes:
            return False, "bassa", "Codice fiscale incompatibile."

    if person.birth_date:
        m = re.search(
            r"(?:data\s+di\s+nascita|nato\s+il|nata\s+il).{0,50}"
            r"(\d{1,2}[./-]\d{1,2}[./-]\d{4})",
            text, re.I | re.S,
        )
        if m:
            found = m.group(1).replace(".", "/").replace("-", "/")
            wanted = {x.replace(".", "/").replace("-", "/") for x in date_variants(person.birth_date)}
            if found not in wanted:
                return False, "bassa", "Data di nascita incompatibile."

    pos = sum(1 for x in CV_POSITIVE if x in n)
    neg = sum(1 for x in CV_NEGATIVE if x in n)
    med = sum(1 for x in MEDICAL_TERMS if x in n)

    if pos < 2 and "curriculum vitae" not in n and "europass" not in n:
        return False, "bassa", "Il documento non appare un CV."
    if med == 0:
        return False, "bassa", "Manca contesto medico."
    if neg >= 3 and "curriculum vitae" not in n:
        return False, "bassa", "Documento concorsuale/amministrativo."

    conf = "alta" if (
        person.fiscal_code and normalize(person.fiscal_code) in n
    ) or pos >= 4 else "media"
    return True, conf, "CV verificato sul contenuto."


def cv_destination(person: Person, cv_dir: Path) -> Path:
    return cv_dir / f"{safe_part(person.output_code)}-{safe_part(person.surname)}-{safe_part(person.name)}.pdf"


def unique_review_destination(person: Person, review_dir: Path, raw: bytes) -> Path:
    review_dir.mkdir(parents=True, exist_ok=True)
    base = f"{safe_part(person.output_code)}-{safe_part(person.surname)}-{safe_part(person.name)}"
    digest = hashlib.sha1(raw).hexdigest()[:10]
    candidate = review_dir / f"{base}-{digest}.pdf"
    return candidate


def looks_like_cv(text: str) -> tuple[bool, str]:
    n = normalize(text)
    pos = sum(1 for x in CV_POSITIVE if x in n)
    neg = sum(1 for x in CV_NEGATIVE if x in n)
    med = sum(1 for x in MEDICAL_TERMS if x in n)

    if "curriculum vitae" in n or "europass" in n:
        return True, "Titolo/struttura CV rilevata."
    if pos >= 2 and med >= 1 and neg < 4:
        return True, "Struttura compatibile con un CV medico."
    if pos >= 3 and neg < 3:
        return True, "Documento con più sezioni tipiche da CV."
    return False, "Documento non sufficientemente compatibile con un CV."


def save_review_candidate(raw: bytes, text: str, url: str, person: Person,
                          review_dir: Path, reason: str) -> tuple[str, str] | None:
    plausible, why = looks_like_cv(text)
    if not plausible:
        logging.info("CV SCARTATO | %s | %s | %s", person.pers_id, url, reason or why)
        return None

    dest = unique_review_destination(person, review_dir, raw)
    if not dest.exists():
        dest.write_bytes(raw)
    logging.info("CV DA VERIFICARE SALVATO | %s | %s | %s", person.pers_id, dest, reason or why)
    return str(dest), url


def try_pdf_url(url: str, person: Person, cv_dir: Path,
                review_dir: Path) -> tuple[str, str, str, str, str] | None:
    got = get_bytes(url, MAX_PDF_BYTES)
    if not got:
        return None
    r, raw = got
    try:
        ctype = normalize(r.headers.get("Content-Type", ""))
        if not (raw.startswith(b"%PDF") or "application/pdf" in ctype):
            return None

        text = pdf_text(raw)
        if not text:
            logging.info("CV SCARTATO | %s | %s | PDF senza testo estraibile", person.pers_id, url)
            return None

        ok, conf, reason = verify_cv(text, person)
        if ok:
            cv_dir.mkdir(parents=True, exist_ok=True)
            dest = cv_destination(person, cv_dir)
            dest.write_bytes(raw)
            logging.info("CV VERIFICATO | %s | %s | confidenza=%s", person.pers_id, dest, conf)
            return "verified", str(dest), text, conf, ""

        review = save_review_candidate(raw, text, url, person, review_dir, reason)
        if review:
            review_path, _ = review
            return "review", review_path, text, "bassa", reason

        return None
    finally:
        r.close()


def try_cv_candidate(url: str, person: Person, cv_dir: Path,
                     review_dir: Path) -> tuple[str, str, str, str, str] | None:
    direct = try_pdf_url(url, person, cv_dir, review_dir)
    if direct:
        kind, path, text, conf, reason = direct
        return kind, path, url, text, conf

    got = get_bytes(url, MAX_HTML_BYTES)
    if not got:
        return None
    r, raw = got
    try:
        soup = BeautifulSoup(raw, "html.parser")
        links = []

        # Link espliciti
        for a in soup.select("a[href]"):
            href = clean(a.get("href"))
            if not href:
                continue
            candidate = urljoin(url, href)
            label = normalize(f"{a.get_text(' ', strip=True)} {candidate}")

            score = 0
            if ".pdf" in candidate.casefold():
                score += 60
            if "curriculum" in label or "europass" in label:
                score += 120
            if normalize(person.surname) in label:
                score += 50
            if normalize(person.name) in label:
                score += 30
            if any(x in label for x in CV_NEGATIVE):
                score -= 120

            if score >= 60:
                links.append((score, candidate))

        # URL PDF scritti nel markup/testo ma non necessariamente dentro <a>
        html = raw.decode("utf-8", errors="ignore")
        for m in re.findall(r'https?://[^"\'<> ]+?\.pdf(?:\?[^"\'<> ]*)?', html, flags=re.I):
            links.append((70, m))

        seen = set()
        for _, candidate in sorted(links, reverse=True)[:20]:
            key = normalize_result_url(candidate)
            if key in seen:
                continue
            seen.add(key)
            found = try_pdf_url(candidate, person, cv_dir, review_dir)
            if found:
                kind, path, text, conf, reason = found
                return kind, path, candidate, text, conf
    finally:
        r.close()
    return None


def hit_looks_pdf(hit: WebHit) -> bool:
    blob = normalize(f"{hit.title} {hit.snippet} {hit.url}")
    return ".pdf" in hit.url.casefold() or "curriculum" in blob or "europass" in blob


def landing_pdf_links(url: str, person: Person) -> list[str]:
    """
    Apre una landing page via HTTP e raccoglie PDF plausibili.
    Non richiede che nome e cognome siano già nello snippet del motore.
    """
    got = get_bytes(url, MAX_HTML_BYTES)
    if not got:
        return []
    r, raw = got
    try:
        ctype = normalize(r.headers.get("Content-Type", ""))
        if raw.startswith(b"%PDF") or "application/pdf" in ctype:
            return [url]

        soup = BeautifulSoup(raw, "html.parser")
        scored = []

        for a in soup.select("a[href]"):
            href = clean(a.get("href"))
            if not href:
                continue
            candidate = urljoin(url, href)
            if not candidate.startswith(("http://", "https://")):
                continue

            label = normalize(f"{a.get_text(' ', strip=True)} {candidate}")
            score = 0
            if ".pdf" in candidate.casefold():
                score += 80
            if "curriculum" in label or "europass" in label or " cv " in f" {label} ":
                score += 120
            if normalize(person.surname) in label:
                score += 55
            if normalize(person.name) in label:
                score += 35
            if any(x in label for x in CV_NEGATIVE):
                score -= 100

            if score >= 70:
                scored.append((score, candidate))

        html = raw.decode("utf-8", errors="ignore")
        pdf_url_pattern = r"https?://[^\"'<> ]+?\.pdf(?:\?[^\"'<> ]*)?"
        for m in re.findall(pdf_url_pattern, html, flags=re.I):
            score = 75
            nm = normalize(m)
            if normalize(person.surname) in nm:
                score += 40
            if normalize(person.name) in nm:
                score += 25
            if "curriculum" in nm or "cv" in nm:
                score += 60
            scored.append((score, m))

        out = []
        seen = set()
        for _, candidate in sorted(scored, reverse=True):
            key = normalize_result_url(candidate)
            if key and key not in seen:
                seen.add(key)
                out.append(candidate)
            if len(out) >= MAX_PDF_LINKS_PER_LANDING:
                break
        return out
    finally:
        r.close()


def inspect_pdf_candidate(url: str, person: Person, cv_dir: Path, review_dir: Path):
    return try_pdf_url(url, person, cv_dir, review_dir)


# ============================================================
# SPECIALITÀ
# ============================================================

# Alias e titoli comuni. Non è usata per "indovinare": serve solo per normalizzare
# frasi esplicite trovate nelle fonti.
SPECIALTY_ALIASES = {
    "anestesia e rianimazione": "Anestesia e Rianimazione",
    "anestesia rianimazione": "Anestesia e Rianimazione",
    "cardiologia": "Cardiologia",
    "chirurgia generale": "Chirurgia Generale",
    "chirurgia vascolare": "Chirurgia Vascolare",
    "dermatologia": "Dermatologia e Venereologia",
    "dermatologia e venereologia": "Dermatologia e Venereologia",
    "ematologia": "Ematologia",
    "endocrinologia": "Endocrinologia",
    "gastroenterologia": "Gastroenterologia",
    "geriatria": "Geriatria",
    "ginecologia e ostetricia": "Ginecologia e Ostetricia",
    "malattie infettive": "Malattie Infettive",
    "medicina del lavoro": "Medicina del Lavoro",
    "medicina dello sport": "Medicina dello Sport",
    "medicina fisica e riabilitativa": "Medicina Fisica e Riabilitativa",
    "medicina interna": "Medicina Interna",
    "medicina legale": "Medicina Legale",
    "nefrologia": "Nefrologia",
    "neurologia": "Neurologia",
    "neurochirurgia": "Neurochirurgia",
    "oculistica": "Oftalmologia",
    "oftalmologia": "Oftalmologia",
    "oncologia": "Oncologia Medica",
    "oncologia medica": "Oncologia Medica",
    "ortopedia": "Ortopedia e Traumatologia",
    "ortopedia e traumatologia": "Ortopedia e Traumatologia",
    "otorinolaringoiatria": "Otorinolaringoiatria",
    "pediatria": "Pediatria",
    "pneumologia": "Malattie dell'Apparato Respiratorio",
    "psichiatria": "Psichiatria",
    "radiodiagnostica": "Radiodiagnostica",
    "radiologia": "Radiodiagnostica",
    "reumatologia": "Reumatologia",
    "urologia": "Urologia",
}

EXPLICIT_PATTERNS = (
    re.compile(r"\bspecialista\s+in\s+([^.;:\n]{3,100})", re.I),
    re.compile(r"\bspecializzato(?:a)?\s+in\s+([^.;:\n]{3,100})", re.I),
    re.compile(r"\bspecializzazione\s+in\s+([^.;:\n]{3,100})", re.I),
    re.compile(r"\bdiploma\s+di\s+specializzazione\s+in\s+([^.;:\n]{3,100})", re.I),
    re.compile(r"\bscuola\s+di\s+specializzazione\s+in\s+([^.;:\n]{3,100})", re.I),
)


def normalize_specialty(raw: str) -> str:
    r = normalize(raw)
    r = re.split(r"\b(?:presso|conseguita|conseguito|università|universita|nel|nell'|anno)\b", r)[0]
    r = r.strip(" ,.-;:")
    if not r:
        return ""

    for alias, canonical in sorted(SPECIALTY_ALIASES.items(), key=lambda x: len(x[0]), reverse=True):
        if alias in r:
            return canonical

    # Se non è in tassonomia, conserviamo solo una forma breve e plausibile.
    raw2 = clean(raw)
    raw2 = re.split(r"[.;:\n|•]", raw2)[0].strip(" ,.-")
    if 3 <= len(raw2) <= 80:
        invalid = {"medico", "medico chirurgo", "medicina e chirurgia", "dirigente medico"}
        if normalize(raw2) not in invalid:
            return raw2
    return ""


def specialty_candidates(text: str, person: Person, source_url: str) -> list[tuple[int, str, str, str]]:
    if not text or identity_score(text, person) < 180:
        return []

    out = []
    ntext = normalize(text)
    authoritative = any(
        x in domain(source_url)
        for x in ("asl", "ausl", "asst", "ats-", "aou", "irccs", "osped",
                  "policlin", "univ", "ordinemedici", "salute")
    )

    for pat in EXPLICIT_PATTERNS:
        for m in pat.finditer(text):
            raw = clean(m.group(1))
            spec = normalize_specialty(raw)
            if not spec:
                continue
            score = 250 + identity_score(text, person)
            if authoritative:
                score += 120
            evidence = clean(m.group(0))[:350]
            out.append((score, spec, source_url, evidence))

    # Nel CV accettiamo anche alias vicino a parole chiave "specializzazione".
    if "specializz" in ntext:
        for alias, canonical in SPECIALTY_ALIASES.items():
            idx = ntext.find(alias)
            if idx >= 0:
                window = ntext[max(0, idx - 160): idx + len(alias) + 160]
                if "specializz" in window:
                    score = 220 + identity_score(text, person)
                    if authoritative:
                        score += 100
                    out.append((score, canonical, source_url,
                                f"Rilevata specializzazione: {canonical}"))
    return out


def choose_specialty(candidates: list[tuple[int, str, str, str]]) -> tuple[str, str, str, str]:
    if not candidates:
        return "", "nessuna", "", ""

    candidates = sorted(candidates, key=lambda x: x[0], reverse=True)
    best = candidates[0]

    # Conflitto forte tra due specialità quasi equivalenti -> non decidere.
    for second in candidates[1:4]:
        if normalize(second[1]) != normalize(best[1]) and second[0] >= best[0] - 40:
            return "", "nessuna", "", ""

    conf = "alta" if best[0] >= 500 else "media"
    return best[1], conf, best[2], best[3]


# ============================================================
# QUERY
# ============================================================

def email_domain(email: str) -> str:
    if "@" not in email:
        return ""
    d = email.rsplit("@", 1)[1].strip().casefold()
    if d in {"gmail.com", "hotmail.com", "outlook.com", "libero.it", "virgilio.it", "yahoo.it", "yahoo.com"}:
        return ""
    return d


def research_queries(person: Person, deep: bool) -> list[str]:
    """
    Query più selettive: identità esatta + contesto sanitario + CV.
    """
    full = f'"{person.full_name}"'
    reverse = f'"{person.surname} {person.name}"'
    qs = []

    d = email_domain(person.email)
    if d:
        qs.append(f'{full} site:{d} curriculum OR specialista OR medico')

    qs.extend([
        f'{full} medico ospedale OR ASL OR AUSL OR ASST OR AOU OR IRCCS',
        f'{full} "curriculum vitae" filetype:pdf medico',
        f'{reverse} medico specialista curriculum',
    ])

    if deep and person.city:
        qs.append(f'{full} medico "{person.city}" specializzazione OR curriculum')

    # Massimo 4 query: meno rumore e meno rischio blocco.
    return qs[:4]


# ============================================================
# PIPELINE
# ============================================================

def is_junk_domain(url: str) -> bool:
    d = domain(url)
    return any(x in d for x in JUNK_DOMAIN_HINTS)


def trusted_medical_domain(url: str) -> bool:
    d = domain(url)
    return any(x in d for x in TRUSTED_MEDICAL_DOMAIN_HINTS)


def hit_relevance(hit: WebHit, person: Person) -> tuple[bool, int, str]:
    """
    Filtra i risultati palesemente estranei prima di aprire pagine/PDF.
    """
    real_url = normalize_result_url(hit.url)
    d = domain(real_url)
    blob = normalize(f"{hit.title} {hit.snippet} {real_url}")

    if not d:
        return False, -999, "dominio assente"

    if is_junk_domain(real_url):
        return False, -500, f"dominio irrilevante: {d}"

    surname = normalize(person.surname)
    name = normalize(person.name)
    full = normalize(person.full_name)
    reverse = normalize(f"{person.surname} {person.name}")

    has_full = bool(full and full in blob) or bool(reverse and reverse in blob)
    has_surname = bool(surname and surname in blob)
    has_name = bool(name and name in blob)
    medical_context = any(x in blob for x in MEDICAL_TERMS)
    trusted = trusted_medical_domain(real_url)
    professional_domain = email_domain(person.email)
    same_prof_domain = bool(professional_domain and (d == professional_domain or d.endswith("." + professional_domain)))

    score = 0
    if has_full:
        score += 220
    elif has_surname and has_name:
        score += 130
    elif has_surname:
        score += 45

    if medical_context:
        score += 80
    if trusted:
        score += 120
    if same_prof_domain:
        score += 220
    if "curriculum" in blob or "europass" in blob or ".pdf" in real_url.casefold():
        score += 70
    if person.city and normalize(person.city) in blob:
        score += 30

    # Accettiamo:
    # - identità forte;
    # - cognome + fonte sanitaria affidabile;
    # - dominio professionale noto.
    relevant = has_full or (has_surname and trusted) or same_prof_domain

    if not relevant:
        return False, score, "identità/fonte insufficienti"

    return True, score, "ok"


def score_hit_for_person(hit: WebHit, person: Person) -> int:
    ok, score, _ = hit_relevance(hit, person)
    if not ok:
        score -= 300
    return score


def research_person(person: Person, args: argparse.Namespace,
                    browser: BrowserSearch, cv_dir: Path, review_dir: Path,
                    cache_dir: Path) -> dict[str, Any]:
    if not args.ignore_cache:
        cached = read_cache(person, cache_dir)
        if cached:
            cached["row"] = person.row
            cached["method"] = clean(cached.get("method")) + " + CACHE"
            return cached

    started = time.perf_counter()
    sources = []
    notes = []
    specialty_pool = []

    cv_path = ""
    cv_url = ""
    cv_conf = "nessuna"
    cv_review_paths = []
    cv_review_urls = []

    # -------- DISCOVERY UNIFICATA: 3-4 QUERY + ANALISI PROFONDA --------
    queries = research_queries(person, args.deep)
    all_hits = []

    # Primo passaggio: motore preferito (AUTO = Google).
    for q in queries:
        hits = browser.search(q)
        all_hits.extend(hits)

    all_hits = dedupe_hits(all_hits)

    def filter_relevant(hits_to_filter):
        accepted = []
        for hit in hits_to_filter:
            ok, rel_score, reason = hit_relevance(hit, person)
            if ok:
                accepted.append(hit)
                logging.info(
                    "HIT ACCETTATO | Pers_Id=%s | engine=%s | score=%s | %s | %s",
                    person.pers_id, hit.engine, rel_score, hit.title[:120], hit.url
                )
            else:
                logging.info(
                    "HIT SCARTATO | Pers_Id=%s | engine=%s | score=%s | %s | %s | %s",
                    person.pers_id, hit.engine, rel_score, reason, hit.title[:100], hit.url
                )
        return accepted

    relevant_hits = filter_relevant(all_hits)

    # Fallback SEMANTICO: se il motore ha restituito 10 risultati ma tutti
    # irrilevanti, proviamo esplicitamente l'altro motore. Nella V3.3 questo
    # non succedeva perché Bing "aveva risultati" e impediva il fallback.
    if not relevant_hits:
        engines_seen = {h.engine for h in all_hits}
        if "google" in engines_seen:
            alternate = "bing"
        else:
            alternate = "google"

        if alternate not in browser.blocked_engines:
            logging.warning(
                "RELEVANCE FALLBACK | Pers_Id=%s | nessun hit rilevante; provo %s",
                person.pers_id, alternate
            )
            alt_hits = []
            for q in queries:
                alt_hits.extend(browser.search_specific(alternate, q))

            alt_hits = dedupe_hits(alt_hits)
            # Evitiamo duplicati fra i due motori.
            existing = {normalize_result_url(h.url) for h in all_hits}
            alt_hits = [h for h in alt_hits if normalize_result_url(h.url) not in existing]
            all_hits.extend(alt_hits)
            relevant_hits.extend(filter_relevant(alt_hits))

    ranked_hits = sorted(
        dedupe_hits(relevant_hits),
        key=lambda h: (
            1 if hit_looks_pdf(h) else 0,
            score_hit_for_person(h, person)
        ),
        reverse=True,
    )

    # Solo fonti realmente rilevanti finiscono nell'Excel.
    sources.extend(h.url for h in ranked_hits[:args.search_results * 2])
    engine_counts = {}
    for h in all_hits:
        engine_counts[h.engine] = engine_counts.get(h.engine, 0) + 1

    logging.info(
        "HIT SUMMARY | Pers_Id=%s | grezzi=%s | rilevanti=%s | engines=%s",
        person.pers_id, len(all_hits), len(ranked_hits), engine_counts
    )

    pdf_seen = set()
    pdf_checked = 0

    # 1) PDF diretti dai risultati.
    for hit in ranked_hits:
        if pdf_checked >= MAX_PDF_CANDIDATES_PER_PERSON:
            break
        if ".pdf" not in hit.url.casefold():
            continue

        key = normalize_result_url(hit.url)
        if not key or key in pdf_seen:
            continue
        pdf_seen.add(key)
        pdf_checked += 1

        found = inspect_pdf_candidate(hit.url, person, cv_dir, review_dir)
        if not found:
            continue

        kind, found_path, cv_text, found_conf, reason = found
        sources.insert(0, hit.url)

        if kind == "verified":
            cv_path, cv_url, cv_conf = found_path, hit.url, found_conf
            notes.append("CV verificato trovato direttamente dai risultati.")
            specialty_pool.extend(specialty_candidates(cv_text, person, hit.url))
            break
        else:
            if found_path not in cv_review_paths:
                cv_review_paths.append(found_path)
            if hit.url not in cv_review_urls:
                cv_review_urls.append(hit.url)
            notes.append("PDF candidato salvato per verifica manuale.")
            specialty_pool.extend(specialty_candidates(cv_text, person, hit.url))

    # 2) Landing page: analizziamo i migliori risultati anche senza identità
    # già perfettamente visibile nello snippet.
    landing_checked = 0
    for hit in ranked_hits:
        if cv_path:
            break
        if landing_checked >= MAX_LANDING_PAGES_PER_PERSON:
            break
        if ".pdf" in hit.url.casefold():
            continue

        rank_score = score_hit_for_person(hit, person)
        blob = normalize(f"{hit.title} {hit.snippet} {hit.url}")
        if rank_score < 80 and not any(
            x in blob for x in (
                "curriculum", "medico", "osped", "asl", "ausl",
                "asst", "aou", "irccs", "policlin"
            )
        ):
            continue

        landing_checked += 1
        sources.append(hit.url)

        text = page_text(hit.url)
        if text:
            specialty_pool.extend(specialty_candidates(text, person, hit.url))

        for candidate in landing_pdf_links(hit.url, person):
            if pdf_checked >= MAX_PDF_CANDIDATES_PER_PERSON:
                break

            key = normalize_result_url(candidate)
            if not key or key in pdf_seen:
                continue
            pdf_seen.add(key)
            pdf_checked += 1

            found = inspect_pdf_candidate(candidate, person, cv_dir, review_dir)
            if not found:
                continue

            kind, found_path, cv_text, found_conf, reason = found
            sources.insert(0, candidate)

            if kind == "verified":
                cv_path, cv_url, cv_conf = found_path, candidate, found_conf
                notes.append("CV verificato trovato tramite landing page.")
                specialty_pool.extend(specialty_candidates(cv_text, person, candidate))
                break
            else:
                if found_path not in cv_review_paths:
                    cv_review_paths.append(found_path)
                if candidate not in cv_review_urls:
                    cv_review_urls.append(candidate)
                notes.append("CV candidato da landing page salvato per verifica manuale.")
                specialty_pool.extend(specialty_candidates(cv_text, person, candidate))

        if cv_path:
            break

    # 3) Specialità anche dagli snippet raccolti.
    for hit in ranked_hits:
        blob = f"{person.full_name}\n{hit.title}\n{hit.snippet}"
        specialty_pool.extend(specialty_candidates(blob, person, hit.url))

    if all_hits and not ranked_hits:
        notes.append("I motori hanno restituito risultati, ma nessuno ha superato il filtro identità/fonte sanitaria.")

    specialty, specialty_conf, specialty_source, specialty_evidence = choose_specialty(specialty_pool)
    if specialty_source:
        sources.insert(0, specialty_source)

    # Se il browser è stato bloccato su tutti i motori disponibili, segnala chiaramente.
    browser_blocked = {"bing", "google"}.issubset(browser.blocked_engines)

    if specialty and cv_path:
        status = "COMPLETATO"
    elif specialty:
        status = "SPECIALITA_TROVATA"
    elif cv_path:
        status = "CV_TROVATO_SPECIALITA_DA_VERIFICARE"
    elif browser_blocked:
        status = "BLOCCATO_BROWSER"
    elif cv_review_paths:
        status = "DA_VERIFICARE"
    elif sources and ranked_hits:
        status = "DA_VERIFICARE"
    else:
        status = "NESSUN_RISULTATO"

    sources = list(dict.fromkeys(normalize_result_url(u) for u in sources if u.startswith(("http://", "https://"))))

    result = {
        "row": person.row,
        "status": status,
        "specialty": specialty,
        "specialty_confidence": specialty_conf,
        "specialty_evidence": specialty_evidence,
        "cv_path": cv_path,
        "cv_url": cv_url,
        "cv_confidence": cv_conf,
        "cv_review_paths": "\n".join(cv_review_paths[:20]),
        "cv_review_urls": "\n".join(cv_review_urls[:20]),
        "sources": "\n".join(sources[:12]),
        "method": f"Playwright browser search V3.4 ({args.engine})",
        "notes": " ".join(notes),
        "updated": utc_now(),
        "error": "",
        "elapsed": time.perf_counter() - started,
    }

    write_cache(person, cache_dir, result)

    logging.info(
        "DONE | Pers_Id=%s | %.1fs | %s | specialita=%s | cv=%s | cv_review=%s",
        person.pers_id, result["elapsed"], status,
        specialty or "N/D", "SI" if cv_path else "NO", len(cv_review_paths)
    )
    return result


# ============================================================
# MAIN
# ============================================================

def main() -> int:
    args = parse_args()
    log_file = configure_logging(args.log_dir)

    input_path = args.input.expanduser().resolve()
    output_path = (
        args.output.expanduser().resolve()
        if args.output
        else default_output_path(input_path).resolve()
    )
    cv_dir = args.cv_dir.expanduser().resolve()
    review_dir = args.cv_review_dir.expanduser().resolve()
    cache_dir = args.cache_dir.expanduser().resolve()
    browser_data_dir = args.browser_data_dir.expanduser().resolve()

    if not input_path.exists():
        raise FileNotFoundError(input_path)

    prepare_output(input_path, output_path)

    wb = load_workbook(output_path)
    if args.sheet not in wb.sheetnames:
        raise ValueError(f"Foglio '{args.sheet}' non trovato. Disponibili: {wb.sheetnames}")

    ws = wb[args.sheet]
    headers = headers_map(ws)
    require_input_columns(headers)
    headers = ensure_output_columns(ws)
    atomic_save(wb, output_path)

    logging.info("=" * 78)
    logging.info("%s", VERSION)
    logging.info("Input: %s", input_path)
    logging.info("Output: %s", output_path)
    logging.info("Engine: %s", args.engine)
    logging.info("Headless: %s", args.headless)
    logging.info("Deep: %s", args.deep)
    logging.info("Search results: %s", args.search_results)
    logging.info("Delay: %.1f - %.1f sec", args.delay_min, args.delay_max)
    logging.info("Browser profile: %s", browser_data_dir)
    logging.info("CV sicuri dir: %s", cv_dir)
    logging.info("CV da verificare dir: %s", review_dir)
    logging.info("Cache dir: %s", cache_dir)
    logging.info("Log: %s", log_file)
    logging.info("=" * 78)

    print(f"Excel output: {output_path}")
    print(f"Log: {log_file.resolve()}")

    people = []
    for row in range(args.start_row, ws.max_row + 1):
        if args.limit is not None and len(people) >= args.limit:
            break

        p = person_from_row(ws, row, headers)
        if not p.pers_id or not p.surname or not p.name:
            continue

        status = clean(getv(ws, row, headers, "Ricerca_Stato")).upper()
        if not args.retry_all and status in {"COMPLETATO", "SPECIALITA_TROVATA"}:
            continue
        if status == "ERRORE" and not args.retry_errors:
            continue

        people.append(p)

    logging.info("Medici da elaborare: %s", len(people))

    if not people:
        atomic_save(wb, output_path)
        return 0

    counts = {}
    start_all = time.perf_counter()
    unsaved = 0

    with BrowserSearch(
        data_dir=browser_data_dir,
        headless=args.headless,
        engine=args.engine,
        delay_min=args.delay_min,
        delay_max=args.delay_max,
        max_results=args.search_results,
    ) as browser:

        for i, person in enumerate(people, start=1):
            try:
                result = research_person(person, args, browser, cv_dir, review_dir, cache_dir)
            except Exception as exc:
                logging.exception("ERRORE PERSONA | %s | %s", person.pers_id, person.full_name)
                result = {
                    "row": person.row,
                    "status": "ERRORE",
                    "specialty": "",
                    "specialty_confidence": "nessuna",
                    "specialty_evidence": "",
                    "cv_path": "",
                    "cv_url": "",
                    "cv_confidence": "nessuna",
                    "cv_review_paths": "",
                    "cv_review_urls": "",
                    "sources": "",
                    "method": f"Playwright browser search V3.4 ({args.engine})",
                    "notes": "",
                    "updated": utc_now(),
                    "error": f"{type(exc).__name__}: {exc}",
                }

            set_result(ws, headers, person.row, result)
            status = clean(result.get("status")) or "?"
            counts[status] = counts.get(status, 0) + 1
            unsaved += 1

            if unsaved >= args.save_every:
                atomic_save(wb, output_path)
                unsaved = 0
                logging.info("CHECKPOINT | %s/%s | %s", i, len(people), output_path)

            if i % 5 == 0 or i == len(people):
                elapsed = time.perf_counter() - start_all
                rate = i / elapsed * 60 if elapsed else 0.0
                eta = (len(people) - i) / rate if rate else 0.0
                logging.info(
                    "PROGRESS | %s/%s | %.2f medici/min | ETA %.1f min | %s",
                    i, len(people), rate, eta, counts
                )

            # Se il browser viene bloccato, salviamo subito e interrompiamo:
            if {"bing", "google"}.issubset(browser.blocked_engines):
                logging.error("STOP | Tutti i motori richiesti risultano bloccati/CAPTCHA.")
                break

    atomic_save(wb, output_path)

    elapsed = time.perf_counter() - start_all
    logging.info("=" * 78)
    logging.info("FINE | %.1f min | %s", elapsed / 60, counts)
    logging.info("OUTPUT | %s", output_path)
    logging.info("=" * 78)

    print("")
    print(f"Excel salvato in: {output_path}")
    print(f"Log esecuzione: {log_file.resolve()}")
    print(f"CV sicuri salvati in: {cv_dir}")
    print(f"CV da verificare salvati in: {review_dir}")
    return 0


if __name__ == "__main__":
    import sys

    early_input, early_output, early_log_dir = parse_early_paths(sys.argv)
    startup_log = bootstrap_log_path(early_log_dir)

    try:
        initial_output = create_initial_output_copy(early_input, early_output)
        if initial_output:
            write_bootstrap_log(startup_log, f"INFO | Excel iniziale: {initial_output}")
            print(f"Excel di output: {initial_output}")
            print(f"Log startup: {startup_log.resolve()}")
    except Exception as exc:
        write_bootstrap_log(
            startup_log,
            f"ERRORE | Creazione output iniziale: {type(exc).__name__}: {exc}"
        )

    ok, missing = import_dependencies(startup_log)
    if not ok:
        print("\nERRORE: mancano dipendenze obbligatorie.")
        print("Installa con: python -m pip install -r requirements_medici_v3.txt")
        print("Poi esegui: python -m playwright install chromium")
        print(f"Log diagnostico: {startup_log.resolve()}")
        for item in missing:
            print(f" - {item}")
        raise SystemExit(2)

    try:
        load_dotenv()
        raise SystemExit(main())
    except KeyboardInterrupt:
        write_bootstrap_log(startup_log, "INTERRUZIONE | KeyboardInterrupt")
        print("\nInterrotto.")
        raise SystemExit(130)
    except Exception as exc:
        write_bootstrap_log(startup_log, f"FATAL | {type(exc).__name__}: {exc}")
        print(f"\nERRORE FATALE: {type(exc).__name__}: {exc}")
        print(f"Log diagnostico: {startup_log.resolve()}")
        raise

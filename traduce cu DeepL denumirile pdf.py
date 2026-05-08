#!/usr/bin/env python3
r"""
Redenumeste PDF-uri tradusand denumirea lor in romana cu DeepL Web.

Flux:
- ia numele PDF-ului fara extensie
- inlocuieste "-" si "_" cu spatii
- elimina cuvintele "compress" si "FINALIZAT"
- trimite textul in DeepL, cu sursa auto-detectata si tinta romana
- redenumeste PDF-ul cu rezultatul din romana

Exemplu:
la-naturaleza-de-la-conciencia_compress_FINALIZAT.pdf
=> Natura constiintei.pdf

Implicit lucreaza in folderul final_pdf al proiectului. Foloseste --folder
daca PDF-urile sunt in alta parte.
"""

import argparse
import json
import logging
import re
import subprocess
import sys
import time
import unicodedata
import urllib.request
from datetime import datetime
from pathlib import Path

from selenium import webdriver
from selenium.common.exceptions import InvalidSessionIdException, TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


PROJECT_DIR = Path(__file__).resolve().parent
PDF_DIR = PROJECT_DIR / "final_pdf"
LOG_DIR = PROJECT_DIR / "logs"
START_CHROME_PS1 = PROJECT_DIR / "PowerShell" / "Start-ChromeDebug.ps1"
STATE_FILE = PROJECT_DIR / "state_deepl_pdf_rename.json"

DEEPL_URL = "https://www.deepl.com/en/translator"
CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CHROME_PROFILE_DIR = r"C:\Users\necul\AppData\Local\Google\Chrome\User Data\Default"
DEBUG_PORT = 9222
DEFAULT_DELAY_SECONDS = 10

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

SOURCE_SELECTOR = (
    'd-textarea[name="source"] div[contenteditable="true"][data-content="true"], '
    'div[contenteditable="true"][aria-labelledby="translation-source-heading"]'
)
TARGET_JS = r"""
const candidates = [
  ...document.querySelectorAll('[aria-labelledby="translation-target-heading"] span.container-target'),
  ...document.querySelectorAll('span.container-target'),
  ...document.querySelectorAll('[aria-labelledby="translation-target-heading"][contenteditable="true"]'),
  ...document.querySelectorAll('d-textarea[name="target"] div[contenteditable="true"][data-content="true"]')
];
for (const el of candidates) {
  const text = (el.innerText || el.textContent || '').trim();
  if (text) return text;
}
return '';
"""


def setup_logging() -> tuple[logging.Logger, Path]:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_path = LOG_DIR / f"deepl_pdf_rename_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    logger = logging.getLogger("deepl_pdf_rename")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(formatter)
    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)
    logger.info("Log: %s", log_path)
    return logger, log_path


logger, LOG_PATH = setup_logging()


def wait_cdp(port: int, timeout: int = 8) -> bool:
    deadline = time.time() + timeout
    url = f"http://127.0.0.1:{port}/json/version"
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=2) as response:
                response.read(2000)
            return True
        except Exception:
            time.sleep(0.5)
    return False


def start_chrome_debug(port: int) -> None:
    if not Path(CHROME_PATH).exists():
        raise FileNotFoundError(f"Nu gasesc Chrome: {CHROME_PATH}")
    if not START_CHROME_PS1.exists():
        raise FileNotFoundError(f"Lipseste scriptul PowerShell: {START_CHROME_PS1}")

    logger.info("Pornesc Chrome debug pentru DeepL pe portul %s.", port)
    result = subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(START_CHROME_PS1),
            "-ChromePath",
            CHROME_PATH,
            "-ProfileDir",
            CHROME_PROFILE_DIR,
            "-DebugPort",
            str(port),
            "-Url",
            DEEPL_URL,
        ],
        capture_output=True,
        text=True,
        timeout=90,
    )
    if result.stdout:
        logger.info("Start-ChromeDebug stdout:\n%s", result.stdout.strip())
    if result.stderr:
        logger.warning("Start-ChromeDebug stderr:\n%s", result.stderr.strip())
    if result.returncode != 0:
        raise RuntimeError(f"Start-ChromeDebug.ps1 a esuat cu cod {result.returncode}")


def connect_driver(port: int, start_chrome: bool = True) -> webdriver.Chrome:
    if not wait_cdp(port, timeout=3):
        if not start_chrome:
            raise RuntimeError(f"Chrome debug nu raspunde pe portul {port}.")
        start_chrome_debug(port)
        if not wait_cdp(port, timeout=60):
            raise RuntimeError(f"Chrome debug nu raspunde pe portul {port}.")

    options = ChromeOptions()
    options.add_argument("--remote-allow-origins=*")
    options.add_experimental_option("debuggerAddress", f"127.0.0.1:{port}")
    driver = webdriver.Chrome(service=ChromeService(), options=options)
    driver.set_page_load_timeout(90)
    driver.set_script_timeout(90)
    return driver


def driver_alive(driver: webdriver.Chrome | None) -> bool:
    if driver is None:
        return False
    try:
        _ = driver.current_url
        return True
    except Exception:
        return False


def is_session_error(exc: Exception) -> bool:
    if isinstance(exc, InvalidSessionIdException):
        return True
    message = str(exc).casefold()
    return "invalid session id" in message or "chrome not reachable" in message


def reset_driver(driver: webdriver.Chrome | None, args: argparse.Namespace) -> webdriver.Chrome:
    try:
        if driver is not None:
            driver.quit()
    except Exception:
        pass
    if not args.no_start_chrome:
        start_chrome_debug(args.debug_port)
        if not wait_cdp(args.debug_port, timeout=60):
            raise RuntimeError(f"Chrome debug nu raspunde pe portul {args.debug_port}.")
    driver = connect_driver(args.debug_port, start_chrome=not args.no_start_chrome)
    keep_single_deepl_tab(driver)
    return driver


def keep_single_deepl_tab(driver: webdriver.Chrome) -> None:
    handles = list(driver.window_handles)
    if not handles:
        return

    keep = handles[0]
    for handle in handles:
        try:
            driver.switch_to.window(handle)
            if "deepl.com" in (driver.current_url or "").casefold():
                keep = handle
                break
        except Exception:
            continue

    for handle in handles:
        if handle == keep:
            continue
        try:
            driver.switch_to.window(handle)
            driver.close()
        except Exception:
            continue

    driver.switch_to.window(keep)
    driver.get(f"{DEEPL_URL}#auto/ro/")
    time.sleep(2)
    dismiss_popups(driver)


def clean_source_name(pdf_path: Path) -> str:
    text = pdf_path.stem
    text = re.sub(r"[-_]+", " ", text)
    text = re.sub(r"\b(compress|finalizat)\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def history_key(text: str) -> str:
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^a-zA-Z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip().casefold()


def fix_mojibake(text: str) -> str:
    if not text:
        return text
    if not re.search(r"[ÃÂâÈÄ]", text):
        return text
    try:
        fixed = text.encode("cp1252").decode("utf-8")
    except UnicodeError:
        return text
    return fixed if fixed else text


def is_raw_pdf_name(pdf_path: Path) -> bool:
    stem = pdf_path.stem.casefold()
    return "compress" in stem or "finalizat" in stem


def safe_pdf_filename(text: str, suffix: str = ".pdf") -> str:
    text = collapse_repeated_text(text)
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", text)
    text = text.strip(" .")
    if text:
        text = text[0].upper() + text[1:]
    if not text:
        text = "Fisier redenumit"
    return f"{text}{suffix}"


def collapse_repeated_text(text: str) -> str:
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) % 2 == 0:
        half = len(text) // 2
        if text[:half].strip().casefold() == text[half:].strip().casefold():
            return text[:half].strip()

    words = text.split()
    if len(words) % 2 == 0:
        half = len(words) // 2
        if " ".join(words[:half]).casefold() == " ".join(words[half:]).casefold():
            return " ".join(words[:half])
    return text


def unique_destination(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    for index in range(1, 1000):
        candidate = path.with_name(f"{stem} ({index}){suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"Nu pot gasi un nume liber pentru: {path}")


def empty_state() -> dict:
    return {
        "version": 1,
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "records": [],
    }


def load_state(path: Path = STATE_FILE) -> dict:
    if not path.exists():
        return empty_state()
    try:
        with path.open("r", encoding="utf-8") as handle:
            state = json.load(handle)
    except Exception as exc:
        logger.warning("Nu pot citi JSON-ul de stare %s: %s. Pornesc cu stare goala.", path, exc)
        return empty_state()
    if not isinstance(state, dict):
        return empty_state()
    state.setdefault("version", 1)
    state.setdefault("records", [])
    return state


def save_state(state: dict, path: Path = STATE_FILE) -> None:
    state["updated_at"] = datetime.now().isoformat(timespec="seconds")
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    with tmp_path.open("w", encoding="utf-8") as handle:
        json.dump(state, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
    tmp_path.replace(path)


def state_source_keys(state: dict) -> set[str]:
    return {
        str(record.get("source_key", ""))
        for record in state.get("records", [])
        if record.get("status") == "renamed" and record.get("source_key")
    }


def upsert_state_record(
    state: dict,
    *,
    source_text: str,
    original_name: str,
    translated_text: str,
    final_name: str,
    status: str = "renamed",
    log_path: Path | None = None,
) -> None:
    source_key = history_key(source_text)
    records = state.setdefault("records", [])
    existing = next((record for record in records if record.get("source_key") == source_key), None)
    record = {
        "source_key": source_key,
        "source_text": source_text,
        "original_name": original_name,
        "translated_text": fix_mojibake(translated_text),
        "final_name": fix_mojibake(final_name),
        "status": status,
        "renamed_at": datetime.now().isoformat(timespec="seconds"),
    }
    if log_path is not None:
        record["log_path"] = str(log_path)
    if existing is None:
        records.append(record)
    else:
        existing.update(record)


def rebuild_state_from_logs(log_dir: Path, state_path: Path) -> dict:
    state = empty_state()
    log_files = sorted(log_dir.glob("deepl_pdf_rename_*.log"))
    for log_file in log_files:
        content = log_file.read_text(encoding="utf-8", errors="replace")
        if "Mod dry-run" in content:
            continue

        current_original = ""
        current_source = ""
        current_translated = ""
        for raw_line in content.splitlines():
            line = raw_line.strip()
            match = re.search(r"\[INFO\] \[\d+/\d+\] (.+)$", line)
            if match:
                current_original = match.group(1).strip()
                current_source = ""
                current_translated = ""
                continue

            match = re.search(r"\[INFO\] Text pentru DeepL: (.+)$", line)
            if match:
                current_source = match.group(1).strip()
                continue

            match = re.search(r"\[INFO\] Tradus: (.+)$", line)
            if match:
                current_translated = collapse_repeated_text(match.group(1).strip())
                continue

            match = re.search(r"\[INFO\] Nume nou: (.+)$", line)
            if match and current_original and current_source and current_translated:
                final_name = fix_mojibake(match.group(1).strip())
                upsert_state_record(
                    state,
                    source_text=current_source,
                    original_name=current_original,
                    translated_text=current_translated,
                    final_name=final_name,
                    log_path=log_file,
                )

    save_state(state, state_path)
    return state


def dismiss_popups(driver: webdriver.Chrome) -> None:
    def click_button(button) -> bool:
        if button.is_displayed() and button.is_enabled():
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
            time.sleep(0.1)
            driver.execute_script("arguments[0].click();", button)
            time.sleep(0.8)
            return True
        return False

    selectors = [
        '[data-testid="cookie-banner-strict-accept-all"]',
        '[data-testid="cookie-banner-strict-accept-selected"]',
        'button:has(span.__content)',
        'button[aria-label="Close"]',
    ]
    for selector in selectors:
        try:
            buttons = driver.find_elements(By.CSS_SELECTOR, selector)
            for button in buttons[:3]:
                if selector == "button:has(span.__content)":
                    text = re.sub(r"\s+", " ", button.text or "").strip()
                    if text != "Accept":
                        continue
                if click_button(button):
                    return
        except Exception:
            continue

    labels = [
        "Accept all",
        "Accept",
        "Reject",
        "Agree",
        "I agree",
        "Got it",
        "Close",
    ]
    for label in labels:
        try:
            buttons = driver.find_elements(
                By.XPATH,
                f"//button[.//span[normalize-space(.)='{label}'] or normalize-space(.)='{label}']",
            )
            for button in buttons[:4]:
                if click_button(button):
                    return
        except Exception:
            continue


def target_language_code(driver: webdriver.Chrome) -> str:
    try:
        return driver.execute_script(
            """
            const el = document.querySelector('[data-testid="translator-target-lang"]');
            if (!el) return '';
            return (el.getAttribute('dl-selected-lang') || el.innerText || '').trim();
            """
        )
    except WebDriverException:
        return ""


def ensure_target_romanian(driver: webdriver.Chrome) -> None:
    code = target_language_code(driver).casefold()
    if code == "ro" or code == "romanian":
        return

    wait = WebDriverWait(driver, 20)
    button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="translator-target-lang-btn"]')))
    driver.execute_script("arguments[0].click();", button)
    option = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '[data-testid="translator-lang-option-ro"]')))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", option)
    driver.execute_script("arguments[0].click();", option)

    deadline = time.time() + 10
    while time.time() < deadline:
        code = target_language_code(driver).casefold()
        if code == "ro" or code == "romanian":
            return
        time.sleep(0.5)
    raise TimeoutException("Nu am putut seta limba tinta DeepL pe Romanian.")


def set_source_text(driver: webdriver.Chrome, text: str) -> None:
    wait = WebDriverWait(driver, 30)
    source = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, SOURCE_SELECTOR)))
    driver.execute_script(
        """
        const el = arguments[0];
        const text = arguments[1];
        el.focus();
        el.innerHTML = '';
        const p = document.createElement('p');
        p.textContent = text;
        el.appendChild(p);
        el.dispatchEvent(new InputEvent('input', {
            bubbles: true,
            cancelable: true,
            inputType: 'insertText',
            data: text
        }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        """,
        source,
        text,
    )
    time.sleep(1)
    try:
        source.click()
        time.sleep(0.2)
        source.send_keys(Keys.CONTROL, "a")
        source.send_keys(Keys.BACKSPACE)
        time.sleep(0.2)
        source.send_keys(text)
    except WebDriverException:
        driver.execute_script(
            """
            const el = arguments[0];
            const text = arguments[1];
            el.focus();
            el.innerHTML = '';
            const p = document.createElement('p');
            p.textContent = text;
            el.appendChild(p);
            el.dispatchEvent(new InputEvent('input', {
                bubbles: true,
                cancelable: true,
                inputType: 'insertText',
                data: text
            }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            """,
            source,
            text,
        )


def current_target_text(driver: webdriver.Chrome) -> str:
    try:
        text = driver.execute_script(TARGET_JS)
        return re.sub(r"\s+", " ", (text or "")).strip()
    except WebDriverException:
        return ""


def read_target_text(
    driver: webdriver.Chrome,
    source_text: str,
    timeout: int = 45,
    previous_target: str = "",
) -> str:
    deadline = time.time() + timeout
    previous = ""
    stable_since = 0.0
    min_len = 2 if len(source_text) <= 5 else max(4, min(8, len(source_text) // 4))
    while time.time() < deadline:
        try:
            text = current_target_text(driver)
            if text != previous:
                previous = text
                stable_since = time.time()
            if (
                text
                and len(text) >= min_len
                and text.casefold() != source_text.casefold()
                and (not previous_target or text.casefold() != previous_target.casefold())
                and time.time() - stable_since >= 2.0
            ):
                return text
        except WebDriverException:
            pass
        time.sleep(0.7)
    raise TimeoutException(f"Nu am gasit traducerea DeepL pentru: {source_text}")


def translate_with_deepl(driver: webdriver.Chrome, text: str) -> str:
    driver.get(f"{DEEPL_URL}#auto/ro/")
    time.sleep(2)
    dismiss_popups(driver)
    ensure_target_romanian(driver)
    previous_target = current_target_text(driver)
    set_source_text(driver, text)
    time.sleep(1)
    if target_language_code(driver).casefold() not in {"ro", "romanian"}:
        logger.warning(
            "DeepL a schimbat limba tinta din Romanian. Pastrez numele curatat pentru: %s",
            text,
        )
        return text
    translated = read_target_text(driver, text, timeout=45, previous_target=previous_target)
    return collapse_repeated_text(translated)


def iter_pdfs(folder: Path, only_name: str | None) -> list[Path]:
    files = sorted(folder.glob("*.pdf"), key=lambda item: item.name.casefold())
    if only_name:
        needle = only_name.casefold()
        files = [path for path in files if needle in path.stem.casefold()]
    return files


def rename_pdf(pdf_path: Path, translated: str, dry_run: bool) -> Path:
    new_name = safe_pdf_filename(translated, pdf_path.suffix)
    destination = unique_destination(pdf_path.with_name(new_name))
    if destination.name == pdf_path.name:
        return pdf_path
    if not dry_run:
        pdf_path.rename(destination)
    return destination


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Traduce cu DeepL denumirile PDF-urilor si le redenumeste in romana."
    )
    parser.add_argument("--folder", default=str(PDF_DIR), help="Folderul cu PDF-uri.")
    parser.add_argument("--limit", type=int, default=0, help="Numar maxim de PDF-uri. 0 = toate.")
    parser.add_argument("--only-name", default="", help="Proceseaza doar PDF-urile care contin acest text in nume.")
    parser.add_argument("--dry-run", action="store_true", help="Afiseaza noul nume fara redenumire.")
    parser.add_argument("--state-file", default=str(STATE_FILE), help="JSON-ul cu istoricul redenumirilor.")
    parser.add_argument(
        "--rebuild-state-from-logs",
        action="store_true",
        help="Reconstruieste JSON-ul de stare din logurile DeepL si se opreste.",
    )
    parser.add_argument(
        "--list-pending",
        action="store_true",
        help="Afiseaza PDF-urile brute care nu exista in JSON, fara Chrome/DeepL.",
    )
    parser.add_argument(
        "--include-clean-names",
        action="store_true",
        help="Proceseaza si PDF-uri fara FINALIZAT/compress. Implicit sunt sarite.",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=DEFAULT_DELAY_SECONDS,
        help="Pauza in secunde intre doua denumiri. Default: 10.",
    )
    parser.add_argument("--no-start-chrome", action="store_true", help="Nu porni Chrome debug automat.")
    parser.add_argument("--debug-port", type=int, default=DEBUG_PORT, help="Portul Chrome debug.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    state_path = Path(args.state_file)

    if args.rebuild_state_from_logs:
        state = rebuild_state_from_logs(LOG_DIR, state_path)
        logger.info("JSON stare reconstruit: %s", state_path)
        logger.info("Intrari in JSON: %s", len(state.get("records", [])))
        return 0

    state = load_state(state_path)
    completed_source_keys = state_source_keys(state)
    logger.info("JSON stare: %s", state_path)
    logger.info("Intrari deja redenumite in JSON: %s", len(completed_source_keys))

    folder = Path(args.folder)
    if not folder.exists():
        logger.error("Folderul nu exista: %s", folder)
        return 2

    pdfs = iter_pdfs(folder, args.only_name or None)
    skipped_history = 0
    if not args.include_clean_names:
        before = len(pdfs)
        pdfs = [path for path in pdfs if is_raw_pdf_name(path)]
        skipped_clean = before - len(pdfs)
    else:
        skipped_clean = 0

    pending_pdfs: list[Path] = []
    for path in pdfs:
        source_text = clean_source_name(path)
        if history_key(source_text) in completed_source_keys:
            skipped_history += 1
            logger.info("Sar deja redenumit in JSON: %s", path.name)
            continue
        pending_pdfs.append(path)
    pdfs = pending_pdfs

    if args.limit > 0:
        pdfs = pdfs[: args.limit]

    if args.list_pending:
        logger.info("Folder PDF: %s", folder)
        logger.info("PDF-uri brute sarite fiindca sunt deja in JSON: %s", skipped_history)
        logger.info("PDF-uri deja curate/fara FINALIZAT/compress sarite: %s", skipped_clean)
        logger.info("PDF-uri noi de redenumit: %s", len(pdfs))
        for path in pdfs:
            logger.info("PENDING: %s | text: %s", path.name, clean_source_name(path))
        return 0

    if not pdfs:
        logger.warning("Nu am gasit PDF-uri de procesat in: %s", folder)
        if skipped_history:
            logger.info("Toate PDF-urile brute gasite erau deja in JSON: %s", skipped_history)
        return 0

    logger.info("Folder PDF: %s", folder)
    logger.info("PDF-uri selectate: %s", len(pdfs))
    if skipped_clean:
        logger.info("Sar PDF-uri deja curate/fara FINALIZAT/compress: %s", skipped_clean)
    if skipped_history:
        logger.info("Sar PDF-uri deja redenumite in JSON: %s", skipped_history)
    logger.info("Pauza intre denumiri: %.1f secunde", args.delay)
    if args.dry_run:
        logger.info("Mod dry-run: nu redenumesc fisierele.")

    driver = reset_driver(None, args)
    renamed = 0
    failed: list[tuple[Path, str]] = []

    try:
        for index, pdf_path in enumerate(pdfs, start=1):
            source_text = clean_source_name(pdf_path)
            logger.info("[%s/%s] %s", index, len(pdfs), pdf_path.name)
            logger.info("Text pentru DeepL: %s", source_text)
            last_error = ""
            for attempt in range(1, 3):
                try:
                    if not driver_alive(driver):
                        logger.warning("Sesiunea Chrome nu mai este activa. Reconectez.")
                        driver = reset_driver(driver, args)
                    translated = translate_with_deepl(driver, source_text)
                    destination = rename_pdf(pdf_path, translated, args.dry_run)
                    if destination.name != pdf_path.name and not args.dry_run:
                        renamed += 1
                        upsert_state_record(
                            state,
                            source_text=source_text,
                            original_name=pdf_path.name,
                            translated_text=translated,
                            final_name=destination.name,
                            log_path=LOG_PATH,
                        )
                        save_state(state, state_path)
                        completed_source_keys.add(history_key(source_text))
                    logger.info("Tradus: %s", translated)
                    logger.info("Nume nou: %s", destination.name)
                    last_error = ""
                    break
                except Exception as exc:
                    last_error = str(exc)
                    if is_session_error(exc):
                        logger.warning(
                            "Sesiune Chrome pierduta la %s, incerc reconectare (%s/2).",
                            pdf_path.name,
                            attempt,
                        )
                        driver = reset_driver(driver, args)
                        continue
                    if attempt == 1:
                        logger.warning("Prima incercare a esuat la %s. Mai incerc o data.", pdf_path.name)
                        time.sleep(3)
                        continue
                    failed.append((pdf_path, last_error))
                    logger.error("Eroare la %s: %s", pdf_path.name, last_error)

            if last_error:
                logger.info("Sar peste fisier dupa eroare: %s", pdf_path.name)
            if index < len(pdfs) and args.delay > 0:
                logger.info("Astept %.1f secunde pana la urmatoarea denumire.", args.delay)
                time.sleep(args.delay)

    finally:
        try:
            driver.quit()
        except Exception:
            pass

    logger.info("=== RAPORT FINAL ===")
    logger.info("Procesate: %s", len(pdfs))
    logger.info("Redenumite: %s", renamed)
    logger.info("Esuate: %s", len(failed))
    for pdf_path, reason in failed:
        logger.info("ESUAT: %s | cauza: %s", pdf_path.name, reason)
    logger.info("Log salvat: %s", LOG_PATH)
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())

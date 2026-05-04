#!/usr/bin/env python3
r"""
Google Translate Docs automation - Chrome version.

Flux:
- scaneaza recursiv g:\ARHIVA\C pentru .doc/.docx
- converteste .doc la .docx cu Microsoft Word
- imparte documentele mai mari de 5 MB in parti _partea1, _partea2, ...
- urca fiecare parte la Google Translate Docs, pe rand
- asteapta 60 secunde dupa traducere, apoi descarca traducerea
- reuneste partile traduse si exporta rezultatul final ca PDF *_FINALIZAT.pdf

Nu modifica fisierele originale din g:\ARHIVA\C.
"""

import argparse
import hashlib
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import time
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

from selenium import webdriver
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, WebDriverException


PROJECT_DIR = Path(__file__).resolve().parent
ARCHIVE_PATH = Path(os.environ.get("SIMPLU_GT_ARCHIVE_PATH", r"g:\ARHIVA\C"))
TRANSLATE_URL = "https://translate.google.ro/?hl=ro&sl=auto&tl=ro&op=docs"

WORK_DIR = PROJECT_DIR / "work"
PARTS_DIR = WORK_DIR / "parts"
CONVERTED_DIR = WORK_DIR / "converted_doc"
DOWNLOADS_DIR = PROJECT_DIR / "downloads"
FINAL_DIR = PROJECT_DIR / "final_pdf"
LOG_DIR = PROJECT_DIR / "logs"
POWERSHELL_DIR = PROJECT_DIR / "PowerShell"
START_CHROME_PS1 = POWERSHELL_DIR / "Start-ChromeDebug.ps1"
STATE_FILE = PROJECT_DIR / "state_google_translate_chrome.json"

MAX_UPLOAD_BYTES = int(os.environ.get("SIMPLU_GT_MAX_BYTES", "5000000"))
MIN_SOURCE_BYTES = int(os.environ.get("SIMPLU_GT_MIN_SOURCE_BYTES", str(50 * 1024)))
MAX_PAGES_PER_PART = int(os.environ.get("SIMPLU_GT_MAX_PAGES_PER_PART", "400"))
TRANSLATE_WAIT_SEC = int(os.environ.get("SIMPLU_GT_TRANSLATE_WAIT_SEC", "60"))
DOWNLOAD_WAIT_SEC = int(os.environ.get("SIMPLU_GT_DOWNLOAD_WAIT_SEC", "420"))
BETWEEN_PARTS_SEC = int(os.environ.get("SIMPLU_GT_BETWEEN_PARTS_SEC", "60"))
MAX_SPLIT_PARTS = int(os.environ.get("SIMPLU_GT_MAX_SPLIT_PARTS", "30"))
TRANSLATE_ERROR_RETRIES = int(os.environ.get("SIMPLU_GT_TRANSLATE_ERROR_RETRIES", "2"))
DOWNLOAD_ERROR_RETRIES = int(os.environ.get("SIMPLU_GT_DOWNLOAD_ERROR_RETRIES", "2"))
KEEP_INTERMEDIATE = os.environ.get("SIMPLU_GT_KEEP_INTERMEDIATE", "0") == "1"
TRANSLATED_DOC_EXTENSIONS = {".doc", ".docx", ".pdf"}

CHROME_PATH = os.environ.get(
    "SIMPLU_CHROME_PATH",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
)
CHROME_PROFILE_DIR = os.environ.get(
    "SIMPLU_CHROME_PROFILE_DIR",
    r"C:\Users\necul\AppData\Local\Google\Chrome\User Data\Default",
)
DEBUG_PORT = int(os.environ.get("SIMPLU_CHROME_DEBUG_PORT", "9222"))

WORD_FORMAT_DOCX = 16
WORD_EXPORT_PDF = 17
WD_STATISTIC_PAGES = 2
WD_GOTO_PAGE = 1
WD_GOTO_ABSOLUTE = 1
WD_PAGE_BREAK = 7


def setup_logging() -> logging.Logger:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_path = LOG_DIR / f"google_translate_chrome_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    logger = logging.getLogger("gt_docs")
    logger.info("Log: %s", log_path)
    return logger


logger = setup_logging()


def ensure_dirs() -> None:
    for directory in [WORK_DIR, PARTS_DIR, CONVERTED_DIR, DOWNLOADS_DIR, FINAL_DIR, LOG_DIR]:
        directory.mkdir(parents=True, exist_ok=True)


def file_mb(path: Path) -> float:
    return path.stat().st_size / (1024 * 1024)


def safe_name(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name[:180] or "document"


def normalize_for_match(text: str) -> str:
    text = unicodedata.normalize("NFKD", text or "")
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return text.lower()


def doc_id(path: Path) -> str:
    raw = str(path.resolve()).lower().encode("utf-8", errors="ignore")
    return hashlib.sha1(raw).hexdigest()[:12]


def alphabetical_key(path: Path) -> str:
    clean = re.sub(r"[^a-zA-Z0-9\s]", " ", path.name.lower())
    return re.sub(r"\s+", " ", clean).strip()


def scan_documents(root: Path) -> list[Path]:
    if not root.exists():
        raise FileNotFoundError(f"Directorul sursa nu exista: {root}")
    docs = [
        p for p in root.rglob("*")
        if p.is_file()
        and p.suffix.lower() in {".doc", ".docx"}
        and not p.name.startswith("~$")
    ]
    docs.sort(key=lambda p: (alphabetical_key(p.parent), alphabetical_key(p)))
    return docs


def load_state() -> dict:
    if STATE_FILE.exists():
        try:
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except Exception as exc:
            logger.warning("Nu pot citi state, pornesc unul nou: %s", exc)
    return {"documents": {}, "updated_at": ""}


def save_state(state: dict) -> None:
    state["updated_at"] = datetime.now().isoformat(timespec="seconds")
    STATE_FILE.write_text(json.dumps(state, indent=2, ensure_ascii=False), encoding="utf-8")


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def find_existing_translation_for_part(part: Path) -> Path | None:
    """Cauta o traducere deja descarcata pentru partea data, dupa numele fisierului."""
    expected_suffix = part.suffix.lower()
    expected_stem = part.stem
    candidates = [
        p for p in DOWNLOADS_DIR.rglob(f"*{part.suffix}")
        if (
            p.is_file()
            and p.suffix.lower() == expected_suffix
            and (
                p.stem == expected_stem
                or re.fullmatch(re.escape(expected_stem) + r" \(\d+\)", p.stem)
            )
        )
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda p: p.stat().st_mtime)


def has_existing_translation(parts: Iterable[Path]) -> bool:
    return any(find_existing_translation_for_part(part) for part in parts)


def collect_translated_parts_from_state_or_disk(
    parts: list[Path],
    existing: dict,
    allow_disk_lookup: bool = True,
) -> list[Path | None]:
    saved = [Path(p) for p in existing.get("translated_parts", []) if p]
    result: list[Path | None] = []

    for idx, part in enumerate(parts):
        found: Path | None = None
        if (
            idx < len(saved)
            and saved[idx].exists()
            and saved[idx].suffix.lower() in TRANSLATED_DOC_EXTENSIONS
        ):
            found = saved[idx]
        if found is None and allow_disk_lookup:
            found = find_existing_translation_for_part(part)
        result.append(found)

    return result


def is_rpc_unavailable(exc: Exception) -> bool:
    text = str(exc).lower()
    return "-2147023174" in text or "rpc server is unavailable" in text


def is_word_corrupt_or_unreadable(exc: Exception) -> bool:
    text = str(exc).lower()
    needles = [
        "appears to be corrupted",
        "file appears to be corrupted",
        "corrupted",
        "is corrupt",
        "word experienced an error trying to open the file",
        "-2146822496",
    ]
    return any(needle in text for needle in needles)


def is_download_timeout(exc: Exception) -> bool:
    return "download-ul traducerii nu a aparut la timp" in str(exc).lower()


def file_signature(path: Path) -> dict:
    try:
        st = path.stat()
        return {"source_size": st.st_size, "source_mtime": st.st_mtime}
    except OSError:
        return {}


def is_same_skipped_source(existing: dict, path: Path) -> bool:
    signature = file_signature(path)
    return (
        existing.get("status") == "skipped"
        and signature
        and existing.get("source_size") == signature.get("source_size")
        and existing.get("source_mtime") == signature.get("source_mtime")
    )


def mark_document_skipped(state: dict, key: str, original: Path, reason: str, detail: str) -> None:
    state["documents"][key] = {
        "original": str(original),
        "status": "skipped",
        "skip_reason": reason,
        "skip_detail": detail[:1000],
        **file_signature(original),
        "updated_at": now_iso(),
    }
    save_state(state)


class WordManager:
    def __init__(self) -> None:
        self.app = None
        self._pythoncom = None
        self._win32com = None

    def __enter__(self):
        self._start_app()
        return self

    def _start_app(self) -> None:
        import pythoncom
        import win32com.client

        self._pythoncom = pythoncom
        self._win32com = win32com.client
        pythoncom.CoInitialize()
        self.app = win32com.client.DispatchEx("Word.Application")
        self.app.Visible = False
        self.app.DisplayAlerts = 0

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    def close(self) -> None:
        if self.app is not None:
            try:
                self.app.Quit()
            except Exception:
                pass
            self.app = None

    def restart(self) -> None:
        logger.warning("Repornesc Microsoft Word COM...")
        self.close()
        time.sleep(2)
        self._start_app()

    def ensure_app(self) -> None:
        if self.app is None:
            self._start_app()
            return
        try:
            _ = self.app.Version
        except Exception as exc:
            logger.warning("Microsoft Word COM nu mai raspunde: %s", exc)
            self.restart()

    def convert_to_docx(self, source: Path) -> Path:
        self.ensure_app()
        if source.suffix.lower() == ".docx" and source.stat().st_size <= MAX_UPLOAD_BYTES:
            return source

        out_dir = CONVERTED_DIR / doc_id(source)
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / f"{safe_name(source.stem)}.docx"
        if source.suffix.lower() == ".docx":
            if out_path.exists() and out_path.stat().st_mtime >= source.stat().st_mtime:
                return out_path
            shutil.copyfile(source, out_path)
            return out_path

        logger.info("Convertesc .doc la .docx: %s", source)
        doc = self.app.Documents.Open(str(source), ReadOnly=True, AddToRecentFiles=False)
        try:
            doc.SaveAs2(str(out_path), FileFormat=WORD_FORMAT_DOCX)
        finally:
            doc.Close(False)
        return out_path

    def split_docx_if_needed(self, source_docx: Path, original: Path, on_part_saved=None) -> list[Path]:
        self.ensure_app()
        doc = self.app.Documents.Open(str(source_docx), ReadOnly=True, AddToRecentFiles=False)
        try:
            doc.Repaginate()
            pages = max(1, int(doc.ComputeStatistics(WD_STATISTIC_PAGES)))
            size = source_docx.stat().st_size
            size_parts = (size + MAX_UPLOAD_BYTES - 1) // MAX_UPLOAD_BYTES
            page_parts = (pages + MAX_PAGES_PER_PART - 1) // MAX_PAGES_PER_PART
            initial_parts = max(1, size_parts, page_parts)

            if initial_parts <= 1:
                logger.info(
                    "Fara split: %s (%.2f MB, %s pagini)",
                    source_docx.name,
                    file_mb(source_docx),
                    pages,
                )
                return [source_docx]

            part_dir = PARTS_DIR / doc_id(original)
            part_dir.mkdir(parents=True, exist_ok=True)
            base = safe_name(original.stem)
            logger.info(
                "Split necesar: %s (%.2f MB, %s pagini; limita %s MB si %s pagini/parte)",
                source_docx,
                file_mb(source_docx),
                pages,
                MAX_UPLOAD_BYTES / (1024 * 1024),
                MAX_PAGES_PER_PART,
            )
            source_stat = source_docx.stat()
            meta_path = part_dir / "_split_progress.json"

            for part_count in range(initial_parts, MAX_SPLIT_PARTS + 1):
                expected_meta = {
                    "source": str(source_docx),
                    "source_size": source_stat.st_size,
                    "source_mtime": source_stat.st_mtime,
                    "pages": pages,
                    "part_count": part_count,
                    "max_upload_bytes": MAX_UPLOAD_BYTES,
                    "max_pages_per_part": MAX_PAGES_PER_PART,
                }
                split_meta: dict = {}
                if meta_path.exists():
                    try:
                        split_meta = json.loads(meta_path.read_text(encoding="utf-8"))
                    except Exception:
                        split_meta = {}
                same_split = all(split_meta.get(k) == v for k, v in expected_meta.items())
                if same_split:
                    logger.info("Resume split existent: %s parti planificate.", part_count)
                else:
                    for old in part_dir.glob(f"{base}_partea*.docx"):
                        old.unlink(missing_ok=True)
                    split_meta = {
                        **expected_meta,
                        "done_parts": [],
                        "updated_at": now_iso(),
                    }
                    meta_path.write_text(json.dumps(split_meta, indent=2, ensure_ascii=False), encoding="utf-8")
                done_by_index = {
                    int(item.get("index")): item
                    for item in split_meta.get("done_parts", [])
                    if item.get("index")
                }

                paths: list[Path] = []
                ranges: list[tuple[int, int]] = []
                for idx in range(part_count):
                    start_page = int(idx * pages / part_count) + 1
                    end_page = int((idx + 1) * pages / part_count)
                    end_page = max(start_page, min(pages, end_page))
                    out_path = part_dir / f"{base}_partea{idx + 1}.docx"
                    done_item = done_by_index.get(idx + 1)
                    if done_item and out_path.exists() and out_path.stat().st_size > 0:
                        logger.info(
                            "Split deja existent: partea %s/%s | pagini %s-%s | %.2f MB | %s",
                            idx + 1,
                            part_count,
                            start_page,
                            end_page,
                            file_mb(out_path),
                            out_path,
                        )
                    else:
                        logger.info(
                            "Split incep: partea %s/%s | pagini %s-%s -> %s",
                            idx + 1,
                            part_count,
                            start_page,
                            end_page,
                            out_path.name,
                        )
                        self._save_page_range(doc, start_page, end_page, out_path)
                        split_meta["done_parts"] = [
                            item for item in split_meta.get("done_parts", [])
                            if int(item.get("index", 0)) != idx + 1
                        ]
                        split_meta["done_parts"].append(
                            {
                                "index": idx + 1,
                                "path": str(out_path),
                                "start_page": start_page,
                                "end_page": end_page,
                                "size": out_path.stat().st_size,
                                "updated_at": now_iso(),
                            }
                        )
                        split_meta["updated_at"] = now_iso()
                        meta_path.write_text(json.dumps(split_meta, indent=2, ensure_ascii=False), encoding="utf-8")
                        logger.info(
                            "Split terminat: partea %s/%s | pagini %s-%s | %.2f MB | %s",
                            idx + 1,
                            part_count,
                            start_page,
                            end_page,
                            file_mb(out_path),
                            out_path,
                        )
                    paths.append(out_path)
                    ranges.append((start_page, end_page))
                    if on_part_saved:
                        on_part_saved(paths.copy(), idx + 1, part_count, start_page, end_page, out_path)

                too_big = [p for p in paths if p.stat().st_size > MAX_UPLOAD_BYTES]
                too_many_pages = [
                    (p, start, end)
                    for p, (start, end) in zip(paths, ranges)
                    if end - start + 1 > MAX_PAGES_PER_PART
                ]
                if not too_big and not too_many_pages:
                    for p, (start, end) in zip(paths, ranges):
                        logger.info(
                            "Parte: %s (%.2f MB, pagini %s-%s, total %s)",
                            p.name,
                            file_mb(p),
                            start,
                            end,
                            end - start + 1,
                        )
                    return paths

                if too_big:
                    logger.info(
                        "Inca exista parti peste limita de MB la %s parti: %s",
                        part_count,
                        ", ".join(f"{p.name}={file_mb(p):.2f}MB" for p in too_big[:3]),
                    )
                if too_many_pages:
                    logger.info(
                        "Inca exista parti peste limita de pagini la %s parti: %s",
                        part_count,
                        ", ".join(f"{p.name}={end - start + 1} pagini" for p, start, end in too_many_pages[:3]),
                    )

            raise RuntimeError(
                f"Nu am reusit sa impart {source_docx.name} in parti sub "
                f"{MAX_UPLOAD_BYTES} bytes si {MAX_PAGES_PER_PART} pagini"
            )
        finally:
            doc.Close(False)

    def _save_page_range(self, doc, start_page: int, end_page: int, out_path: Path) -> None:
        start = doc.GoTo(What=WD_GOTO_PAGE, Which=WD_GOTO_ABSOLUTE, Count=start_page).Start
        pages = max(1, int(doc.ComputeStatistics(WD_STATISTIC_PAGES)))
        if end_page >= pages:
            end = doc.Content.End
        else:
            end = doc.GoTo(What=WD_GOTO_PAGE, Which=WD_GOTO_ABSOLUTE, Count=end_page + 1).Start - 1

        src_range = doc.Range(Start=start, End=max(start, end))
        new_doc = self.app.Documents.Add()
        try:
            new_doc.Range(0, 0).FormattedText = src_range.FormattedText
            new_doc.SaveAs2(str(out_path), FileFormat=WORD_FORMAT_DOCX)
        finally:
            new_doc.Close(False)

    def export_translated_parts_to_pdf(self, translated_parts: list[Path], original: Path) -> Path:
        last_error: Exception | None = None
        for attempt in range(1, 3):
            try:
                self.ensure_app()
                return self._export_translated_parts_to_pdf_once(translated_parts, original)
            except Exception as exc:
                last_error = exc
                if attempt == 1 and is_rpc_unavailable(exc):
                    logger.warning("Word COM/RPC a cazut la export PDF. Repornesc Word si reincerc.")
                    self.restart()
                    continue
                raise
        raise RuntimeError(f"Nu am putut exporta PDF pentru {original}") from last_error

    def _export_translated_parts_to_pdf_once(self, translated_parts: list[Path], original: Path) -> Path:
        final_base = f"{safe_name(original.stem)}_FINALIZAT"
        final_docx = WORK_DIR / f"{final_base}_{doc_id(original)}.docx"
        final_pdf = FINAL_DIR / f"{final_base}.pdf"

        if len(translated_parts) == 1:
            part = translated_parts[0]
            if part.suffix.lower() == ".pdf":
                shutil.copyfile(str(part), str(final_pdf))
                return final_pdf
            if part.suffix.lower() not in {".doc", ".docx"}:
                raise ValueError(f"Extensie tradusa nesuportata pentru PDF: {part}")

            src_doc = self.app.Documents.Open(str(part), ReadOnly=True, AddToRecentFiles=False)
            try:
                src_doc.ExportAsFixedFormat(OutputFileName=str(final_pdf), ExportFormat=WORD_EXPORT_PDF)
            finally:
                src_doc.Close(False)
            return final_pdf

        doc = self.app.Documents.Add()
        try:
            for idx, part in enumerate(translated_parts):
                if part.suffix.lower() not in {".doc", ".docx"}:
                    raise ValueError(f"Extensie tradusa nesuportata pentru merge: {part}")

                src_doc = self.app.Documents.Open(str(part), ReadOnly=True, AddToRecentFiles=False)
                try:
                    if idx > 0:
                        rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                        rng.InsertBreak(WD_PAGE_BREAK)
                    rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                    rng.FormattedText = src_doc.Content.FormattedText
                finally:
                    src_doc.Close(False)

            doc.SaveAs2(str(final_docx), FileFormat=WORD_FORMAT_DOCX)
            doc.ExportAsFixedFormat(OutputFileName=str(final_pdf), ExportFormat=WORD_EXPORT_PDF)
        finally:
            doc.Close(False)

        if not KEEP_INTERMEDIATE:
            final_docx.unlink(missing_ok=True)
        return final_pdf


@dataclass
class PreparedDocument:
    original: Path
    upload_parts: list[Path]


class GoogleTranslateRetryableError(RuntimeError):
    pass


class ChromeTranslateBot:
    def __init__(self, download_dir: Path):
        self.download_dir = download_dir
        self.driver = None
        self.wait = None

    def _wait_cdp(self, timeout: int = 60) -> bool:
        import urllib.request

        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                with urllib.request.urlopen(f"http://127.0.0.1:{DEBUG_PORT}/json/version", timeout=3) as resp:
                    resp.read(2000)
                return True
            except Exception:
                time.sleep(1)
        return False

    def _start_chrome_debug(self) -> None:
        if not Path(CHROME_PATH).exists():
            raise FileNotFoundError(f"Nu gasesc Chrome: {CHROME_PATH}")
        if not START_CHROME_PS1.exists():
            raise FileNotFoundError(f"Lipseste scriptul PowerShell: {START_CHROME_PS1}")

        logger.info("Pornesc Chrome debug pe profilul: %s", CHROME_PROFILE_DIR)
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
                str(DEBUG_PORT),
                "-Url",
                TRANSLATE_URL,
            ],
            capture_output=True,
            text=True,
            timeout=90,
        )
        if result.stdout:
            logger.info("PowerShell Start-ChromeDebug stdout:\n%s", result.stdout.strip())
        if result.stderr:
            logger.warning("PowerShell Start-ChromeDebug stderr:\n%s", result.stderr.strip())
        if result.returncode != 0:
            raise RuntimeError(f"Start-ChromeDebug.ps1 a esuat cu cod {result.returncode}")

    def start(self) -> None:
        if not self._wait_cdp(timeout=3):
            self._start_chrome_debug()
            if not self._wait_cdp(timeout=90):
                raise RuntimeError(f"Chrome debug nu raspunde pe portul {DEBUG_PORT}")

        options = ChromeOptions()
        options.add_argument("--remote-allow-origins=*")
        options.add_experimental_option("debuggerAddress", f"127.0.0.1:{DEBUG_PORT}")
        service = ChromeService()
        self.driver = webdriver.Chrome(service=service, options=options)
        self.wait = WebDriverWait(self.driver, 45)
        self.driver.set_page_load_timeout(90)
        self.driver.set_script_timeout(90)
        self.download_dir.mkdir(parents=True, exist_ok=True)
        try:
            self.driver.execute_cdp_cmd(
                "Page.setDownloadBehavior",
                {"behavior": "allow", "downloadPath": str(self.download_dir)},
            )
        except Exception as exc:
            logger.warning("Nu pot seta download dir prin CDP: %s", exc)
        self._open_translate_single_tab()

    def _open_translate_single_tab(self) -> None:
        assert self.driver is not None
        handles = list(self.driver.window_handles)
        if not handles:
            self.driver.switch_to.new_window("tab")
            handles = list(self.driver.window_handles)

        keep_handle = None
        for handle in handles:
            try:
                self.driver.switch_to.window(handle)
                if "translate.google" in (self.driver.current_url or ""):
                    keep_handle = handle
                    break
            except WebDriverException:
                continue
        if keep_handle is None and handles:
            keep_handle = handles[0]

        for handle in list(handles):
            if handle == keep_handle:
                continue
            try:
                self.driver.switch_to.window(handle)
                self.driver.close()
                logger.info("Am inchis un tab Chrome extra.")
            except WebDriverException:
                pass

        if keep_handle:
            self.driver.switch_to.window(keep_handle)
        self.driver.get(TRANSLATE_URL)

    def close(self) -> None:
        # Nu inchidem Chrome: descarcarile sau taburile pot ramane utile pentru inspectie.
        pass

    def translate_file(self, upload_path: Path) -> Path:
        if self.driver is None or self.wait is None:
            raise RuntimeError("ChromeTranslateBot nu este pornit")

        last_error: Exception | None = None
        max_attempts = max(TRANSLATE_ERROR_RETRIES, DOWNLOAD_ERROR_RETRIES) + 1
        for attempt in range(1, max_attempts + 1):
            try:
                return self._translate_file_once(upload_path, attempt)
            except GoogleTranslateRetryableError as exc:
                last_error = exc
                if attempt > TRANSLATE_ERROR_RETRIES:
                    raise
                logger.warning(
                    "Google Translate a refuzat temporar fisierul (%s/%s): %s",
                    attempt,
                    TRANSLATE_ERROR_RETRIES + 1,
                    exc,
                )
                self._open_translate_single_tab()
                time.sleep(10)
            except TimeoutException as exc:
                last_error = exc
                existing = find_existing_translation_for_part(upload_path)
                if existing and existing.exists():
                    logger.info(
                        "Download gasit dupa timeout, folosesc fisierul existent: %s",
                        existing,
                    )
                    return existing
                if is_download_timeout(exc) and attempt <= DOWNLOAD_ERROR_RETRIES:
                    logger.warning(
                        "Download-ul nu a aparut la timp pentru %s (%s/%s). Reincerc aceeasi parte.",
                        upload_path.name,
                        attempt,
                        DOWNLOAD_ERROR_RETRIES + 1,
                    )
                    self._open_translate_single_tab()
                    time.sleep(15)
                    continue
                raise

        raise RuntimeError(f"Nu am putut traduce dupa retry: {upload_path}") from last_error

    def _translate_file_once(self, upload_path: Path, attempt: int) -> Path:
        assert self.driver is not None and self.wait is not None
        logger.info("Traduc: %s (incercarea %s)", upload_path, attempt)
        self._open_translate_single_tab()
        self._dismiss_popups()

        input_el = self.wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']"))
        )
        input_el.send_keys(str(upload_path.resolve()))

        self._click_button(["Traduceti", "Translate", "Traducir"], timeout=90)
        logger.info("Traducere pornita; astept %s secunde inainte de download.", TRANSLATE_WAIT_SEC)
        self._wait_translate_delay_or_error(TRANSLATE_WAIT_SEC)

        before = self._download_snapshot()
        self._click_button(
            ["Descarcati traducerea", "Download translation", "Download"],
            timeout=180,
        )
        downloaded = self._wait_for_download(before, timeout=DOWNLOAD_WAIT_SEC)
        logger.info("Descarcat: %s", downloaded)
        return downloaded

    def _dismiss_popups(self) -> None:
        for texts in [
            ["Accepta tot", "Accept all", "Sunt de acord", "I agree"],
            ["Acceptă tot", "Accept all", "Sunt de acord", "I agree"],
            ["Nu acum", "Not now"],
        ]:
            try:
                self._click_button(texts, timeout=3, required=False)
            except Exception:
                pass

    def _click_button(self, texts: list[str], timeout: int, required: bool = True) -> bool:
        assert self.driver is not None
        deadline = time.time() + timeout
        wanted = texts
        while time.time() < deadline:
            error_text = self._translate_error_text()
            if error_text:
                raise GoogleTranslateRetryableError(error_text)

            clicked = self.driver.execute_script(
                """
                const normalize = (value) => (value || '')
                    .normalize('NFD')
                    .replace(/[\\u0300-\\u036f]/g, '')
                    .toLowerCase();
                const wanted = arguments[0].map(normalize);
                const nodes = Array.from(document.querySelectorAll('button, [role="button"]'));
                for (const el of nodes) {
                    const text = (el.innerText || el.textContent || '').trim();
                    const textNorm = normalize(text);
                    const disabled = el.disabled || el.getAttribute('aria-disabled') === 'true';
                    if (!disabled && wanted.some(w => textNorm.includes(w))) {
                        el.scrollIntoView({block: 'center', inline: 'center'});
                        el.click();
                        return text;
                    }
                }
                return null;
                """,
                wanted,
            )
            if clicked:
                logger.info("Click buton: %s", clicked)
                return True
            time.sleep(1)

        if required:
            raise TimeoutException(f"Nu am gasit butonul: {texts}")
        return False

    def _wait_translate_delay_or_error(self, seconds: int) -> None:
        deadline = time.time() + seconds
        while time.time() < deadline:
            error_text = self._translate_error_text()
            if error_text:
                raise GoogleTranslateRetryableError(error_text)
            time.sleep(2)

    def _translate_error_text(self) -> str | None:
        assert self.driver is not None
        try:
            text = self.driver.execute_script(
                "return document.body ? (document.body.innerText || document.body.textContent || '') : '';"
            )
        except WebDriverException:
            return None

        normalized = normalize_for_match(str(text))
        error_needles = [
            "momentan fisierul nu poate fi tradus",
            "fisierul nu poate fi tradus",
            "incercati din nou peste cateva minute",
            "currently the file cannot be translated",
            "this file cannot be translated",
            "try again in a few minutes",
            "the file could not be translated",
        ]
        if any(needle in normalized for needle in error_needles):
            for line in str(text).splitlines():
                if any(needle in normalize_for_match(line) for needle in error_needles):
                    return line.strip()
            return "Momentan, fisierul nu poate fi tradus. Incercati din nou peste cateva minute."
        return None

    def _download_snapshot(self) -> set[str]:
        return {p.name for p in self.download_dir.glob("*") if p.is_file()}

    def _wait_for_download(self, before: set[str], timeout: int) -> Path:
        start_time = time.time()
        deadline = time.time() + timeout
        while time.time() < deadline:
            candidates = [
                p for p in self.download_dir.glob("*")
                if p.is_file()
                and p.name not in before
                and p.suffix.lower() in TRANSLATED_DOC_EXTENSIONS
            ]
            if candidates:
                newest = max(candidates, key=lambda p: p.stat().st_mtime)
                if not (self.download_dir / f"{newest.name}.crdownload").exists():
                    return newest
            time.sleep(2)
        current_docs = [
            p.name for p in self.download_dir.glob("*")
            if p.is_file() and p.suffix.lower() in TRANSLATED_DOC_EXTENSIONS
        ]
        partials = [
            p.name for p in self.download_dir.glob("*")
            if p.is_file() and p.stat().st_mtime >= start_time and p.name.lower().endswith((".crdownload", ".tmp"))
        ]
        logger.warning(
            "Timeout download. Director=%s | documente=%s | partiale recente=%s",
            self.download_dir,
            current_docs[-10:],
            partials[-10:],
        )
        raise TimeoutException("Download-ul traducerii nu a aparut la timp")


def prepare_document(word: WordManager, path: Path, on_split_part_saved=None) -> PreparedDocument:
    docx_path = word.convert_to_docx(path)
    parts = word.split_docx_if_needed(docx_path, path, on_part_saved=on_split_part_saved)
    return PreparedDocument(original=path, upload_parts=parts)


def cleanup_document_intermediates(prepared: PreparedDocument, downloads: Iterable[Path]) -> None:
    if KEEP_INTERMEDIATE:
        return
    for p in prepared.upload_parts:
        try:
            if PROJECT_DIR in p.resolve().parents:
                p.unlink(missing_ok=True)
        except Exception:
            pass


def process_documents(args: argparse.Namespace) -> int:
    ensure_dirs()
    state = load_state()
    docs = scan_documents(ARCHIVE_PATH)
    known_docs = {doc_id(p): p for p in docs}
    resume_docs: list[Path] = []
    for key, entry in state.get("documents", {}).items():
        if entry.get("status") == "done" and not args.force:
            continue
        if key in known_docs:
            continue
        original_text = entry.get("original")
        parts = [Path(p) for p in entry.get("parts", []) if p]
        translated = [Path(p) for p in entry.get("translated_parts", []) if p]
        if original_text and (
            any(p.exists() for p in parts + translated)
            or has_existing_translation(parts)
        ):
            resume_docs.append(Path(original_text))
    if resume_docs:
        logger.info(
            "Resume: adaug %s documente din state, chiar daca nu mai sunt in sursa.",
            len(resume_docs),
        )
        docs = resume_docs + docs
    if args.only_name:
        needle = args.only_name.lower()
        docs = [p for p in docs if needle in p.name.lower()]
    if args.max_files:
        docs = docs[: args.max_files]

    logger.info("Documente gasite: %s", len(docs))
    if not docs:
        return 0

    run_download_dir = DOWNLOADS_DIR / datetime.now().strftime("%Y%m%d_%H%M%S")
    bot = None
    completed = 0

    with WordManager() as word:
        for index, original in enumerate(docs, 1):
            key = doc_id(original)
            existing = state["documents"].get(key, {})
            final_pdf = FINAL_DIR / f"{safe_name(original.stem)}_FINALIZAT.pdf"
            state_final_pdf = Path(existing.get("final_pdf", "")) if existing.get("final_pdf") else None
            if (
                existing.get("status") == "done"
                and not args.force
                and (final_pdf.exists() or (state_final_pdf and state_final_pdf.exists()))
            ):
                logger.info("[%s/%s] Skip deja finalizat: %s", index, len(docs), original)
                continue
            if original.exists() and is_same_skipped_source(existing, original) and not args.force:
                logger.info(
                    "[%s/%s] Skip deja marcat: %s | motiv=%s",
                    index,
                    len(docs),
                    original,
                    existing.get("skip_reason", "necunoscut"),
                )
                continue

            if original.exists():
                source_size = original.stat().st_size
                if source_size < MIN_SOURCE_BYTES:
                    detail = f"Fisier prea mic: {source_size} bytes, sub limita {MIN_SOURCE_BYTES} bytes"
                    logger.warning("[%s/%s] Skip fisier incomplet/prea mic: %s | %s", index, len(docs), original, detail)
                    mark_document_skipped(state, key, original, "too_small", detail)
                    continue

                logger.info("[%s/%s] Pregatesc: %s (%.2f MB)", index, len(docs), original, file_mb(original))

                def checkpoint_split(saved_parts, done_index, total_parts, start_page, end_page, out_path):
                    state["documents"][key] = {
                        "original": str(original),
                        "status": "splitting",
                        "parts": [str(p) for p in saved_parts],
                        "translated_parts": existing.get("translated_parts", []),
                        "split_progress": {
                            "done": done_index,
                            "total": total_parts,
                            "last_part": str(out_path),
                            "last_pages": [start_page, end_page],
                            "updated_at": now_iso(),
                        },
                        "updated_at": now_iso(),
                    }
                    save_state(state)

                try:
                    prepared = prepare_document(word, original, on_split_part_saved=checkpoint_split)
                except Exception as exc:
                    if is_rpc_unavailable(exc):
                        logger.warning(
                            "[%s/%s] Word COM/RPC a cazut la pregatire. Repornesc Word si reincerc: %s",
                            index,
                            len(docs),
                            original,
                        )
                        word.restart()
                        try:
                            prepared = prepare_document(word, original, on_split_part_saved=checkpoint_split)
                        except Exception as retry_exc:
                            if is_word_corrupt_or_unreadable(retry_exc):
                                detail = str(retry_exc)
                                logger.warning(
                                    "[%s/%s] Skip fisier corupt/necitibil dupa retry: %s | %s",
                                    index,
                                    len(docs),
                                    original,
                                    detail,
                                )
                                mark_document_skipped(state, key, original, "corrupt_or_unreadable", detail)
                                continue
                            raise
                    elif is_word_corrupt_or_unreadable(exc):
                        detail = str(exc)
                        logger.warning(
                            "[%s/%s] Skip fisier corupt/necitibil: %s | %s",
                            index,
                            len(docs),
                            original,
                            detail,
                        )
                        mark_document_skipped(state, key, original, "corrupt_or_unreadable", detail)
                        continue
                    else:
                        raise
            else:
                state_parts = [Path(p) for p in existing.get("parts", []) if p]
                resume_parts = [p for p in state_parts if p.exists()]
                if not resume_parts and args.finalize_existing and has_existing_translation(state_parts):
                    resume_parts = state_parts
                if not resume_parts:
                    logger.warning(
                        "[%s/%s] Nu gasesc originalul si nici partile pregatite, sar: %s",
                        index,
                        len(docs),
                        original,
                    )
                    continue
                logger.info(
                    "[%s/%s] Reiau din state fara original in sursa: %s | parti=%s",
                    index,
                    len(docs),
                    original,
                    len(resume_parts),
                )
                prepared = PreparedDocument(original=original, upload_parts=resume_parts)
            logger.info("Parti upload: %s", len(prepared.upload_parts))

            existing_part_paths = [Path(p) for p in existing.get("parts", []) if p]
            parts_changed = (
                bool(existing_part_paths)
                and [str(p) for p in existing_part_paths] != [str(p) for p in prepared.upload_parts]
            )
            resume_existing = dict(existing)
            allow_disk_resume = True
            if parts_changed:
                logger.info(
                    "Schema de split s-a schimbat pentru %s; ignor traducerile vechi si refac partile.",
                    original.name,
                )
                resume_existing["translated_parts"] = []
                allow_disk_resume = False

            state["documents"][key] = {
                "original": str(original),
                "status": "prepared" if args.prepare_only else existing.get("status", "running"),
                "parts": [str(p) for p in prepared.upload_parts],
                "translated_parts": resume_existing.get("translated_parts", []),
                "updated_at": now_iso(),
            }
            save_state(state)

            if args.prepare_only:
                for p in prepared.upload_parts:
                    logger.info("  parte: %s (%.2f MB)", p, file_mb(p))
                continue

            translated_slots = collect_translated_parts_from_state_or_disk(
                prepared.upload_parts,
                resume_existing,
                allow_disk_lookup=allow_disk_resume,
            )
            translated_parts: list[Path | None] = list(translated_slots)
            found_count = sum(1 for p in translated_parts if p and p.exists())
            if found_count:
                logger.info(
                    "Resume: am gasit deja %s/%s parti traduse pentru %s",
                    found_count,
                    len(prepared.upload_parts),
                    original.name,
                )

            if args.finalize_existing and found_count < len(prepared.upload_parts):
                logger.info(
                    "Finalize-only: nu am toate partile traduse pentru %s (%s/%s), sar.",
                    original.name,
                    found_count,
                    len(prepared.upload_parts),
                )
                continue

            for part_index, part in enumerate(prepared.upload_parts, 1):
                existing_download = translated_parts[part_index - 1]
                if existing_download and existing_download.exists():
                    logger.info(
                        "Skip upload parte %s/%s, traducere existenta: %s",
                        part_index,
                        len(prepared.upload_parts),
                        existing_download,
                    )
                    continue

                if args.finalize_existing:
                    continue

                if bot is None:
                    bot = ChromeTranslateBot(run_download_dir)
                    bot.start()

                logger.info("Upload parte %s/%s: %s", part_index, len(prepared.upload_parts), part.name)
                try:
                    downloaded = bot.translate_file(part)
                except (TimeoutException, GoogleTranslateRetryableError) as exc:
                    logger.error(
                        "Nu am putut finaliza upload/download pentru %s, partea %s/%s. Marchez pentru reluare: %s",
                        original.name,
                        part_index,
                        len(prepared.upload_parts),
                        exc,
                    )
                    state["documents"][key] = {
                        "original": str(original),
                        "status": "download_failed",
                        "parts": [str(p) for p in prepared.upload_parts],
                        "translated_parts": [str(p) for p in translated_parts if p],
                        "failed_part_index": part_index,
                        "failed_part": str(part),
                        "error": str(exc)[:1000],
                        "updated_at": now_iso(),
                    }
                    save_state(state)
                    break
                translated_parts[part_index - 1] = downloaded
                state["documents"][key] = {
                    "original": str(original),
                    "status": "running",
                    "parts": [str(p) for p in prepared.upload_parts],
                    "translated_parts": [str(p) for p in translated_parts if p],
                    "updated_at": now_iso(),
                }
                save_state(state)
                if part_index < len(prepared.upload_parts):
                    logger.info("Pauza intre parti: %s secunde", BETWEEN_PARTS_SEC)
                    time.sleep(BETWEEN_PARTS_SEC)

            ready_parts = [p for p in translated_parts if p and p.exists()]
            if len(ready_parts) < len(prepared.upload_parts):
                logger.warning(
                    "Nu pot face merge pentru %s: am %s/%s parti traduse.",
                    original.name,
                    len(ready_parts),
                    len(prepared.upload_parts),
                )
                continue

            pdf_path = word.export_translated_parts_to_pdf(ready_parts, original)
            logger.info("PDF final: %s", pdf_path)
            state["documents"][key] = {
                "original": str(original),
                "status": "done",
                "parts": [str(p) for p in prepared.upload_parts],
                "translated_parts": [str(p) for p in ready_parts],
                "final_pdf": str(pdf_path),
                "updated_at": now_iso(),
            }
            save_state(state)
            cleanup_document_intermediates(prepared, ready_parts)
            completed += 1

    logger.info("Gata. Documente finalizate in aceasta rulare: %s", completed)
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Google Translate Docs Chrome automation")
    parser.add_argument("--prepare-only", action="store_true", help="Doar converteste/spliteaza, fara browser/upload")
    parser.add_argument(
        "--finalize-existing",
        action="store_true",
        help="Doar face merge/PDF pentru traducerile deja descarcate, fara browser/upload",
    )
    parser.add_argument("--max-files", type=int, default=int(os.environ.get("SIMPLU_GT_MAX_FILES", "0")) or None)
    parser.add_argument("--force", action="store_true", help="Reproceseaza chiar daca PDF-ul final exista")
    parser.add_argument("--only-name", help="Proceseaza doar fisierele care contin acest text in nume")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    logger.info("ARCHIVE_PATH=%s", ARCHIVE_PATH)
    logger.info("MAX_UPLOAD_BYTES=%s", MAX_UPLOAD_BYTES)
    logger.info("MIN_SOURCE_BYTES=%s", MIN_SOURCE_BYTES)
    logger.info("MAX_PAGES_PER_PART=%s", MAX_PAGES_PER_PART)
    logger.info("TRANSLATE_WAIT_SEC=%s", TRANSLATE_WAIT_SEC)
    logger.info("BETWEEN_PARTS_SEC=%s", BETWEEN_PARTS_SEC)
    return process_documents(args)


if __name__ == "__main__":
    raise SystemExit(main())

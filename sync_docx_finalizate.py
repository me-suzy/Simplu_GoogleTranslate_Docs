from __future__ import annotations

import re
import shutil
import unicodedata
from datetime import datetime
from pathlib import Path


SOURCE_DIR = Path(r"G:\ARHIVA\C")
DONE_DIR = SOURCE_DIR / "GATA FINALIZAT"
PDF_DIRS = [
    Path(r"D:\Simplu_GoogleTranslate_Docs\final_pdf"),
    Path(r"D:\ENGLEZA\PDF-uri convertite"),
]
REPORT_DIR = Path(r"D:\Simplu_GoogleTranslate_Docs\logs")


def normalize_name(text: str) -> str:
    text = unicodedata.normalize("NFKD", text or "")
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def pdf_source_stem(pdf: Path) -> str:
    return re.sub(r"_FINALIZAT$", "", pdf.stem, flags=re.IGNORECASE)


def unique_destination(directory: Path, name: str) -> Path:
    dest = directory / name
    if not dest.exists():
        return dest

    stem = Path(name).stem
    suffix = Path(name).suffix
    counter = 1
    while True:
        candidate = directory / f"{stem} ({counter}){suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def source_docs() -> list[Path]:
    docs = []
    for path in SOURCE_DIR.glob("*"):
        if (
            path.is_file()
            and path.suffix.lower() in {".doc", ".docx"}
            and not path.name.startswith("~$")
        ):
            docs.append(path)
    return sorted(docs, key=lambda p: normalize_name(p.name))


def pdf_lookup() -> dict[str, Path]:
    lookup: dict[str, Path] = {}
    for pdf_dir in PDF_DIRS:
        if not pdf_dir.exists():
            continue
        for pdf in pdf_dir.glob("*.pdf"):
            key = normalize_name(pdf_source_stem(pdf))
            if key:
                lookup[key] = pdf
    return lookup


def main() -> int:
    if not SOURCE_DIR.exists():
        raise FileNotFoundError(f"Nu exista folderul sursa: {SOURCE_DIR}")

    DONE_DIR.mkdir(parents=True, exist_ok=True)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)

    pdfs = pdf_lookup()
    moved: list[tuple[Path, Path, Path]] = []
    remaining: list[Path] = []

    for doc in source_docs():
        pdf = pdfs.get(normalize_name(doc.stem))
        if not pdf:
            remaining.append(doc)
            continue

        dest = unique_destination(DONE_DIR, doc.name)
        shutil.move(str(doc), str(dest))
        moved.append((doc, dest, pdf))

    report_lines = [
        f"Raport sync DOCX finalizate - {datetime.now():%Y-%m-%d %H:%M:%S}",
        f"Sursa DOCX: {SOURCE_DIR}",
        f"Folder finalizate: {DONE_DIR}",
        "PDF-uri comparate:",
        *[f"  - {p}" for p in PDF_DIRS],
        "",
        f"Mutate in GATA FINALIZAT: {len(moved)}",
    ]

    for src, dest, pdf in moved:
        report_lines.append(f"  MUTAT: {src.name} -> {dest} | PDF: {pdf}")

    report_lines.extend(["", f"Ramase netraduse in sursa: {len(remaining)}"])
    for doc in remaining:
        report_lines.append(f"  RAMAS: {doc}")

    report = "\n".join(report_lines)
    report_path = REPORT_DIR / f"sync_docx_finalizate_{datetime.now():%Y%m%d_%H%M%S}.log"
    report_path.write_text(report + "\n", encoding="utf-8")

    print(report)
    print(f"\nRaport salvat: {report_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Application service for safe, in-process listing report generation.

This module is deliberately independent from Streamlit.  The UI, tests, and any
future API can call the same bounded processing pipeline.
"""

from __future__ import annotations

import io
import re
import tempfile
import zipfile
from collections.abc import Sequence
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Literal

import pymupdf as fitz

from .engine import (
    detect_listing_type_from_text,
    make_combined_imei_report,
    make_combined_report,
    make_imei_report,
    make_report,
    parse_shared_imei_users,
)

ListingKind = Literal["auto", "phone", "imei"]
OutputMode = Literal["combined", "separate"]

MAX_FILES = 10
MAX_FILE_BYTES = 25 * 1024 * 1024
MAX_TOTAL_BYTES = 75 * 1024 * 1024
MAX_PAGES_PER_FILE = 1_000
MAX_TOTAL_PAGES = 2_500
MAX_OUTPUT_BYTES = 30 * 1024 * 1024
DETECTION_PAGE_BATCH = 8
DOCX_MIME = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)


class ProcessingError(ValueError):
    """A user-correctable validation or parsing error."""


@dataclass(frozen=True)
class Upload:
    filename: str
    content: bytes


@dataclass(frozen=True)
class ProcessingOptions:
    listing_type: ListingKind = "auto"
    output_mode: OutputMode = "combined"
    start_hour: int = 22
    end_hour: int = 7
    top_n: int = 5
    add_extra_present: bool = True
    theft_date: str = ""
    enable_internet_search: bool = False
    serpapi_api_key: str = ""
    max_web_results: int = 3


@dataclass(frozen=True)
class DetectedFile:
    filename: str
    listing_type: str
    identifier: str = ""
    pages: int = 0


@dataclass(frozen=True)
class GeneratedReport:
    filename: str
    mime_type: str
    content: bytes
    listing_type: str
    detected_files: list[DetectedFile]
    summaries: list[dict[str, object]]


def safe_filename(value: str, fallback: str = "listing.pdf") -> str:
    """Strip paths and filesystem-significant characters from an upload name."""
    name = Path(value or fallback).name
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("._")
    return cleaned or fallback


def _validate_options(options: ProcessingOptions) -> None:
    if options.listing_type not in {"auto", "phone", "imei"}:
        raise ProcessingError("Type de listing invalide.")
    if options.output_mode not in {"combined", "separate"}:
        raise ProcessingError("Mode de sortie invalide.")
    if not 0 <= options.start_hour <= 23 or not 0 <= options.end_hour <= 23:
        raise ProcessingError("Les heures doivent être comprises entre 0 et 23.")
    if not 1 <= options.top_n <= 20:
        raise ProcessingError("Le Top correspondants doit être compris entre 1 et 20.")
    if not 1 <= options.max_web_results <= 5:
        raise ProcessingError("Le nombre de résultats web doit être compris entre 1 et 5.")
    if options.enable_internet_search and not options.serpapi_api_key.strip():
        raise ProcessingError(
            "La recherche web est activée, mais aucune clé SerpAPI n'est configurée."
        )


def _validate_pdf(content: bytes, filename: str) -> int:
    if not content.startswith(b"%PDF-"):
        raise ProcessingError(f"{filename} n'est pas un fichier PDF valide.")
    try:
        document = fitz.open(stream=content, filetype="pdf")
        if document.needs_pass:
            raise ProcessingError(f"{filename} est protégé par un mot de passe.")
        pages = document.page_count
        document.close()
    except ProcessingError:
        raise
    except Exception as exc:
        raise ProcessingError(f"Lecture impossible de {filename}.") from exc
    if pages < 1:
        raise ProcessingError(f"{filename} ne contient aucune page.")
    if pages > MAX_PAGES_PER_FILE:
        raise ProcessingError(
            f"{filename} dépasse la limite de {MAX_PAGES_PER_FILE} pages."
        )
    return pages


def detect_pdf_type(content: bytes, filename: str) -> tuple[str, str]:
    """Scan a PDF progressively until a supported document signature is found."""
    document = fitz.open(stream=content, filetype="pdf")
    try:
        parts: list[str] = []
        detected = "unknown"
        detection_text = ""
        for index in range(document.page_count):
            parts.append(document[index].get_text("text"))
            batch_end = (index + 1) % DETECTION_PAGE_BATCH == 0
            document_end = index + 1 == document.page_count
            if not batch_end and not document_end:
                continue
            detection_text = "\n".join(parts)
            detected = detect_listing_type_from_text(detection_text, filename)
            if detected != "unknown":
                break
        return detected, detection_text
    finally:
        document.close()


def _resolve_listing_type(
    requested: ListingKind, detected: Sequence[DetectedFile]
) -> Literal["phone", "imei"]:
    unknown = [item.filename for item in detected if item.listing_type == "unknown"]
    if unknown:
        raise ProcessingError(
            "Type de listing indéterminé pour : " + ", ".join(unknown) + "."
        )
    kinds = {item.listing_type for item in detected}
    if len(kinds) != 1:
        raise ProcessingError(
            "Lot mixte refusé : séparez les listings IMEI des listings téléphone."
        )
    actual = next(iter(kinds))
    if requested != "auto" and requested != actual:
        expected_label = "IMEI" if requested == "imei" else "téléphone"
        actual_label = "IMEI" if actual == "imei" else "téléphone"
        raise ProcessingError(
            f"Mode {expected_label} sélectionné, mais document {actual_label} détecté."
        )
    return "imei" if actual == "imei" else "phone"


def _summary_rows(summaries: Sequence[object]) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for summary in summaries:
        row = asdict(summary)
        row["input_file"] = Path(str(row["input_file"])).name
        row["output_file"] = Path(str(row["output_file"])).name
        rows.append(row)
    return rows


def process_uploads(
    uploads: Sequence[Upload], options: ProcessingOptions | None = None
) -> GeneratedReport:
    """Validate, classify, analyze, and render one bounded batch of PDF uploads."""
    options = options or ProcessingOptions()
    _validate_options(options)
    if not 1 <= len(uploads) <= MAX_FILES:
        raise ProcessingError(f"Sélectionnez entre 1 et {MAX_FILES} PDF.")

    normalized: list[Upload] = []
    detected: list[DetectedFile] = []
    total_bytes = 0
    total_pages = 0
    for index, upload in enumerate(uploads, start=1):
        filename = safe_filename(upload.filename, f"listing-{index}.pdf")
        if not filename.lower().endswith(".pdf"):
            raise ProcessingError(f"{filename} doit avoir l'extension .pdf.")
        if len(upload.content) > MAX_FILE_BYTES:
            raise ProcessingError(f"{filename} dépasse la limite de 25 Mo.")
        total_bytes += len(upload.content)
        if total_bytes > MAX_TOTAL_BYTES:
            raise ProcessingError("La sélection dépasse la limite totale de 75 Mo.")
        pages = _validate_pdf(upload.content, filename)
        total_pages += pages
        if total_pages > MAX_TOTAL_PAGES:
            raise ProcessingError(
                f"La sélection dépasse la limite totale de {MAX_TOTAL_PAGES} pages."
            )
        detected_type, text = detect_pdf_type(upload.content, filename)
        identifier = ""
        if detected_type == "imei":
            users = parse_shared_imei_users(text)
            if users:
                identifier = max(users, key=lambda item: item.occurrences).imei
        detected.append(DetectedFile(filename, detected_type, identifier, pages))
        normalized.append(Upload(filename, upload.content))

    listing_type = _resolve_listing_type(options.listing_type, detected)

    with tempfile.TemporaryDirectory(prefix="cdr-insight-") as temp_dir:
        root = Path(temp_dir)
        inputs_dir = root / "inputs"
        outputs_dir = root / "outputs"
        inputs_dir.mkdir()
        outputs_dir.mkdir()
        paths: list[Path] = []
        for index, upload in enumerate(normalized, start=1):
            path = inputs_dir / f"{index:02d}-{upload.filename}"
            path.write_bytes(upload.content)
            paths.append(path)

        combined = options.output_mode == "combined" and len(paths) > 1
        summaries: list[object] = []
        output_paths: list[Path] = []
        try:
            if listing_type == "imei" and combined:
                output, summaries = make_combined_imei_report(
                    paths,
                    outputs_dir,
                    top_n=options.top_n,
                    theft_date=options.theft_date,
                )
                output_paths = [Path(output)]
            elif listing_type == "imei":
                for path in paths:
                    summary = make_imei_report(
                        path,
                        outputs_dir,
                        top_n=options.top_n,
                        theft_date=options.theft_date,
                    )
                    summaries.append(summary)
                    output_paths.append(Path(summary.output_file))
            elif combined:
                output, summaries = make_combined_report(
                    paths,
                    outputs_dir,
                    start_hour=options.start_hour,
                    end_hour=options.end_hour,
                    top_n=options.top_n,
                    add_extra_present=options.add_extra_present,
                    enable_internet_search=options.enable_internet_search,
                    serpapi_api_key=options.serpapi_api_key,
                    max_web_results=options.max_web_results,
                )
                output_paths = [Path(output)]
            else:
                for path in paths:
                    summary = make_report(
                        path,
                        outputs_dir,
                        start_hour=options.start_hour,
                        end_hour=options.end_hour,
                        top_n=options.top_n,
                        add_extra_present=options.add_extra_present,
                        enable_internet_search=options.enable_internet_search,
                        serpapi_api_key=options.serpapi_api_key,
                        max_web_results=options.max_web_results,
                    )
                    summaries.append(summary)
                    output_paths.append(Path(summary.output_file))
        except ValueError as exc:
            raise ProcessingError(str(exc)) from exc

        if not output_paths:
            raise ProcessingError("Aucun rapport n'a été généré.")
        if len(output_paths) == 1:
            filename = output_paths[0].name
            mime_type = DOCX_MIME
            content = output_paths[0].read_bytes()
        else:
            buffer = io.BytesIO()
            with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
                for path in output_paths:
                    archive.write(path, arcname=path.name)
            filename = "rapports_cdr_insight.zip"
            mime_type = "application/zip"
            content = buffer.getvalue()
        if len(content) > MAX_OUTPUT_BYTES:
            raise ProcessingError("Le rapport généré dépasse la taille autorisée.")

    return GeneratedReport(
        filename=filename,
        mime_type=mime_type,
        content=content,
        listing_type=listing_type,
        detected_files=detected,
        summaries=_summary_rows(summaries),
    )

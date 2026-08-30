from __future__ import annotations

import json
import os
import re
import unicodedata
from collections.abc import Iterable
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from urllib.parse import urlencode, urlparse
from urllib.request import Request, urlopen

import pymupdf as fitz
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

PHONE_TOKEN_RE = re.compile(r"^(?:237)?\d{9,16}$")
DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}$")
DATE_WORD_RE = re.compile(r"^(?:\d{2}/\d{2}/\d{4}|\d{4}-\d{2}-\d{2})$")
TIME_WORD_RE = re.compile(r"^\d{1,2}:\d{2}:\d{2}$")
CALL_DURATION_RE = re.compile(r"^(?:\d{1,2}:){2}\d{1,2}$")
SMS_DURATION_RE = re.compile(r"^(?:SMS(?:/DATA)?|SMS-DATA|MISSEDCALL|MOBILEMONEY|MTN|ORANGE|DATA)$", re.I)
IMEI_RE = re.compile(r"^\d{14,17}$")
REGIONS = {
    "ADAMAOUA", "CENTRE", "EST", "EAST", "EXTREME-NORD", "FAR-NORTH",
    "LITTORAL", "NORD", "NORTH", "NORD-OUEST", "NORTH-WEST",
    "OUEST", "WEST", "SUD", "SOUTH", "SUD-OUEST", "SOUTH-WEST"
}
STOP_SECTION_RE = re.compile(
    r"\n\s*(?:III\.\s*Identification|IV\.\s*Fr[ée]quence|V\.\s*Fr[ée]quence|VI\.\s*Fr[ée]quence|VII\.\s*Fr[ée]quence)",
    re.I,
)

# The standalone package deliberately ships without customer or product branding.
# Applications may pass an explicit logo path when generating a report.
DEFAULT_LOGO_PATH: Path | None = None
IMEI_LOGO_PATH: Path | None = None
CORRESPONDENT_COLUMN_VERSION = "cdr-insight-correspondent-columns-v1"


@dataclass
class SubscriberInfo:
    number: str = ""
    name: str = ""
    operator: str = ""
    cni: str = ""
    imei: str = ""
    address: str = ""
    period: str = ""
    requester: str = ""


@dataclass
class IdentificationRecord:
    number: str = ""
    name: str = ""
    cni: str = ""
    address: str = ""


@dataclass
class CallRow:
    # caller = number to which the location belongs. For calls this is the real caller.
    # For SMS destination-only tables, caller is the localized destination number.
    caller_raw: str
    caller: str | None
    called_raw: str
    called: str | None
    location_raw: str
    location_clean: str
    district: str
    city: str
    lat: str
    lon: str
    cell: str
    imei: str
    dt: datetime
    duration: str
    source: str = "CALLS"

    @property
    def is_call(self) -> bool:
        return bool(CALL_DURATION_RE.match(self.duration or ""))

    @property
    def is_smsdata(self) -> bool:
        return not self.is_call

    @property
    def is_known_site(self) -> bool:
        text = (self.location_clean or "").lower()
        district = (self.district or "").lower()
        if "site inconnu" in text or "site inconnu" in district:
            return False
        if district in {"", "n/d", "nd", "n.a", "n/a"}:
            return False
        if re.match(r"^cell\s*:", text, re.I) or re.match(r"^cell\s*:", district, re.I):
            return False
        return True


@dataclass
class DistrictChoice:
    district: str = "N/D"
    city: str = "N/D"
    lat: str = ""
    lon: str = ""
    prefix_nd: bool = False
    source: str = "N/D"
    total_count: int = 0
    total_days: int = 0
    window_count: int = 0
    window_days: int = 0

    @property
    def display(self) -> str:
        if self.district == "N/D":
            return "N/D"
        return f"N/D-{self.district}" if self.prefix_nd else self.district

    @property
    def localisation(self) -> str:
        if self.lat and self.lon:
            return f"Lat : {self.lat} Long : {self.lon}"
        return "N/D"


@dataclass
class ReportSummary:
    input_file: str
    output_file: str
    suspect_number: str
    suspect_name: str
    records: int
    correspondents: int
    retained_district: str
    parser: str = "hybrid"
    warnings: str = ""


@dataclass
class WebSearchResult:
    label: str
    url: str


@dataclass
class ListingAnalysis:
    input_file: str
    info: SubscriberInfo
    rows: list[CallRow]
    parser_desc: str
    warnings: list[str]
    suspect: str
    suspect_choice: DistrictChoice
    counts: dict[str, int]
    top: list[str]
    selected: list[str]
    identities: dict[str, IdentificationRecord]
    certified_call_counts: dict[str, int] = field(default_factory=dict)
    certified_sms_counts: dict[str, int] = field(default_factory=dict)
    certified_call_durations: dict[str, int] = field(default_factory=dict)


@dataclass
class ImeiUserObservation:
    number: str
    imei: str
    name: str
    occurrences: int
    first_seen: datetime
    last_seen: datetime


@dataclass
class ImeiListingAnalysis:
    input_file: str
    imei: str
    operator: str
    period: str
    requester: str
    shared_users: list[ImeiUserObservation]
    dominant_user: ImeiUserObservation
    rows: list[CallRow]
    identities: dict[str, IdentificationRecord]
    parser_desc: str
    warnings: list[str]
    theft_date: str = ""


@dataclass
class WinDevParseResult:
    """Structured data extracted from the WinDev requisition PDF variant."""

    info: SubscriberInfo
    rows: list[CallRow]
    identities: dict[str, IdentificationRecord]
    parser_desc: str


@dataclass
class PdfKitMtnParseResult:
    """Structured data extracted from the strict MTN/PDFKit requisition layout."""

    info: SubscriberInfo
    rows: list[CallRow]
    identities: dict[str, IdentificationRecord]
    parser_desc: str
    expected_events: int


def clean_space(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip(" :;,-\n\t")


def clean_null(value: str) -> str:
    v = clean_space(value)
    vl = v.lower().strip()
    vl_norm = re.sub(r"[^a-z0-9éèêàç/ ]+", " ", vl)
    vl_norm = clean_space(vl_norm)
    if vl_norm in {"null", "n/a", "na", "none", "nd", "n/d", "adresse n/a", "imei n/a", "numero cni n/a"}:
        return ""
    if vl_norm.startswith(("adresse n/a", "adresse na", "imei n/a", "numero cni n/a")):
        return ""
    if vl_norm in {"non identifie", "non identifié"}:
        return "Non Identifié"
    if vl_norm.startswith(("et prénom null", "et prenom null")) or vl_norm.endswith(" null"):
        return ""
    return v


def normalize_phone(value: str) -> str | None:
    digits = re.sub(r"\D", "", value or "")
    if not digits:
        return None
    if digits.startswith("237") and len(digits) >= 12:
        digits = digits[3:]
    if len(digits) == 9 and digits.startswith("6"):
        return digits
    # Some provider exports contain long route identifiers such as 2371600000000001.
    # The real 9-digit number is often the last 9 digits.
    if len(digits) > 9 and digits[-9:].startswith("6"):
        return digits[-9:]
    return None


def display_phone(number: str) -> str:
    n = normalize_phone(number) or re.sub(r"\D", "", number or "")
    if len(n) == 9:
        return f"{n[0]} {n[1:3]} {n[3:5]} {n[5:7]} {n[7:9]}"
    return number or "N/D"


def extract_text(pdf_path: str | Path) -> str:
    with fitz.open(str(pdf_path)) as doc:
        return "\n".join(page.get_text("text") for page in doc)


def _fold_windev(value: object) -> str:
    """Accent-insensitive normalization used only for WinDev table signatures."""
    raw = unicodedata.normalize("NFKD", str(value or ""))
    ascii_text = "".join(char for char in raw if not unicodedata.combining(char))
    return clean_space(re.sub(r"[^a-z0-9]+", " ", ascii_text.lower()))


def pdf_producer(pdf_path: str | Path) -> str:
    """Return the PDF producer metadata used to separate export families."""
    try:
        with fitz.open(str(pdf_path)) as doc:
            return clean_space((doc.metadata or {}).get("producer", ""))
    except Exception:
        return ""


def is_windev_requisition(text: str, producer: str = "") -> bool:
    """Return True only for the new WinDev telecom requisition layout.

    Keeping this detector strict prevents the table parser below from changing the
    existing ANTIC/Orange/MTN parsing paths.
    """
    folded = _fold_windev(text)
    folded_producer = _fold_windev(producer)
    # ReportLab exports can contain the same visible headings as WinDev but use
    # a different table geometry. Never route a known non-WinDev producer to the
    # WinDev parser. Unknown metadata still uses the structural signatures below.
    if folded_producer and "windev" not in folded_producer:
        return False
    required = (
        "requisition relative aux identites des abonnes" in folded,
        "numero appelant" in folded and "localisation numero appelant" in folded,
        "date debut appel" in folded and "duree de l appel" in folded,
    )
    supporting = (
        "identification des numeros" in folded,
        "frequence par cellule" in folded,
        "service demandeur" in folded,
    )
    return all(required) and sum(supporting) >= 2


def is_pdfkit_mtn_requisition(text: str, producer: str = "") -> bool:
    """Recognize only the MTN requisition family generated by PDFKit.

    The producer check and the two exact table signatures deliberately keep this
    parser away from the established ReportLab, WinDev and legacy PDF routes.
    """
    folded = _fold_windev(text)
    folded_producer = _fold_windev(producer)
    if "pdfkit" not in folded_producer:
        return False
    identity_signature = (
        "requisition relative aux identites des abonnes de telephonie mobile" in folded
        and "operateur mtn cameroon" in folded
    )
    calls_signature = all(
        marker in folded
        for marker in (
            "ii listing appels",
            "numero appellant",
            "imei appellant",
            "localisation appellant",
            "numero appele",
            "date debut appel",
            "duree",
        )
    )
    sms_signature = all(
        marker in folded
        for marker in (
            "iii listing sms",
            "numero d envoi",
            "localisation numero destination",
            "imei numero destination",
            "date sms",
            "numero destination",
        )
    )
    return identity_signature and calls_signature and sms_signature


def extract_text_until_tables_end(text: str) -> str:
    m = STOP_SECTION_RE.search(text)
    return text[: m.start()] if m else text


def _first_regex(pattern: str, text: str, flags: int = re.I | re.S) -> str:
    m = re.search(pattern, text, flags)
    return clean_space(m.group(1)) if m else ""


def _line_groups_from_words(pdf_path: str | Path, max_pages: int = 2) -> list[tuple[float, list[tuple[float, str]]]]:
    groups: list[tuple[float, list[tuple[float, str]]]] = []
    with fitz.open(str(pdf_path)) as doc:
        for page in doc[:max_pages]:
            words = sorted(page.get_text("words"), key=lambda w: (w[1], w[0]))
            current_y: float | None = None
            current: list[tuple[float, str]] = []
            for x0, y0, _x1, _y1, word, *_ in words:
                if current_y is None or abs(y0 - current_y) <= 3.0:
                    if current_y is None:
                        current_y = y0
                    current.append((x0, word))
                else:
                    groups.append((current_y, sorted(current)))
                    current_y = y0
                    current = [(x0, word)]
            if current_y is not None and current:
                groups.append((current_y, sorted(current)))
    return groups


def parse_identity_from_words(pdf_path: str | Path) -> SubscriberInfo:
    info = SubscriberInfo()
    lines = _line_groups_from_words(pdf_path, max_pages=2)
    for _, items in lines:
        text = clean_space(" ".join(w for _, w in items))
        low = text.lower()
        if "téléphone" in low or "telephone" in low:
            for _, w in items:
                n = normalize_phone(w)
                if n:
                    info.number = info.number or n
                    break
        if "nom" in low and "prénom" in low or "nom" in low and "prenom" in low:
            # take content after the label, usually right of x=300 or after ':'
            right = [w for x, w in items if x > 300 and not re.fullmatch(r":", w)]
            if right:
                info.name = info.name or clean_null(" ".join(right))
        if "opérateur" in low or "operateur" in low:
            if re.search(r"\bORANGE\b", text, re.I):
                info.operator = info.operator or "ORANGE"
            elif re.search(r"\bMTN\b", text, re.I):
                info.operator = info.operator or "MTN"
        if "numéro" in low and "cni" in low or "numero" in low and "cni" in low:
            right = [w for x, w in items if x > 300 and not re.fullmatch(r":", w)]
            if right:
                candidate = clean_null(" ".join(right))
                if candidate and "num" not in candidate.lower() and "cni" not in candidate.lower():
                    info.cni = info.cni or candidate
        if "adresse" in low:
            right = [w for x, w in items if x > 300 and not re.fullmatch(r":", w)]
            if right:
                info.address = info.address or clean_null(" ".join(right))
    return info


def parse_subscriber_info(text: str, pdf_path: str | Path | None = None) -> SubscriberInfo:
    flat = clean_space(text)
    period = _first_regex(r"Période\s*:\s*du\s*([0-9/]+\s*(?:au|Au)\s*[0-9/]+)", flat)
    requester = _first_regex(r"Service\s+Demandeur\s*:?\s*(.*?)(?:\s+I\.\s*Ident|\s+I\.\s+Ident|\s+Téléphone\s*:|\s+REPUBLIQUE)", flat)
    phone = normalize_phone(_first_regex(r"Téléphone\s*:?\s*((?:237)?\d{9,12})", flat)) or ""
    name = _first_regex(r"Nom\s+et\s+Pr[ée]nom\s*:?\s*(.*?)\s+Op[ée]rateur\s+t[ée]l[ée]phonique", flat)
    operator = _first_regex(r"Op[ée]rateur\s+t[ée]l[ée]phonique\s*:?\s*(.*?)\s+Num[ée]ro\s+CNI", flat)
    cni = _first_regex(r"Num[ée]ro\s+CNI\s*:?\s*([^\s]+)", flat)
    imei = _first_regex(r"IMEI\s*:?\s*(.*?)\s+Adresse\s*: ?", flat)
    address = _first_regex(
        r"Adresse\s*:?\s*(.*?)(?:\s+II\.\s*Listing|\s+A\.\s*Listing|\s+Num[eé]ro\s+appell?ant|\s+Numero\s+appell?ant|\s+Localisation\s+num[eé]ro|\Z)",
        flat,
    )
    info = SubscriberInfo(
        phone, clean_null(name), clean_null(operator), clean_null(cni), clean_null(imei), clean_null(address), clean_null(period), clean_null(requester)
    )
    if pdf_path:
        w = parse_identity_from_words(pdf_path)
        # Fill missing or obviously polluted fields from coordinate-based identity rows.
        for field in ("number", "name", "operator", "cni", "address"):
            v = getattr(info, field)
            wv = getattr(w, field)
            if (not v or len(v) > 80 or "Date debut appel" in v or "Numero appelé" in v) and wv:
                setattr(info, field, wv)
    return info




def _is_hexish_identifier(value: str) -> bool:
    v = clean_space(value)
    if not v:
        return False
    if normalize_phone(v):
        return False
    # In several ANTIC/Orange PDFs, non-phone service identifiers appear as hex-like strings.
    return bool(re.fullmatch(r"[A-F0-9]{6,24}", v, flags=re.I))


def _looks_like_name_token(value: str) -> bool:
    v = clean_space(value)
    if not v or v.lower() in {"inconnu", "null", "n/a", "na"}:
        return True
    if DATE_WORD_RE.match(v) or TIME_WORD_RE.match(v):
        return False
    if normalize_phone(v) or IMEI_RE.match(v):
        return False
    if re.fullmatch(r"\d+", v):
        return False
    if _is_hexish_identifier(v):
        return False
    if re.search(r"REQ-GN|20777|page", v, re.I):
        return False
    return True


def _clean_ident_name(value: str) -> str:
    v = clean_null(value)
    if not v:
        return ""
    # Remove accidental header/noise fragments and collapse repeated spaces.
    v = re.sub(r"\b(?:Nom|pr[ée]nom|Numero|Num[ée]ro|Appelant|Adresse|Nationalit[ée])\b", " ", v, flags=re.I)
    v = re.sub(r"\bREQ-GN[-\w]*\b|\b20777\b", " ", v, flags=re.I)
    return clean_null(v)


def _clean_ident_address(value: str) -> str:
    v = clean_null(value)
    if not v:
        return ""
    v = re.sub(r"\bREQ-GN[-\w]*\b|\b20777\b", " ", v, flags=re.I)
    return clean_null(v)


def parse_identification_map_from_words(pdf_path: str | Path) -> dict[str, IdentificationRecord]:
    """Parse the 'III. Identification' table to enrich correspondents.

    This targets the common ANTIC/Orange/MTN horizontal table where each row starts
    with 'Numero Appelant' followed by name, CNI and address columns. It is more
    reliable than plain-text regex because names and addresses can be split over
    two visual lines in the PDF.
    """
    records: dict[str, IdentificationRecord] = {}
    try:
        doc = fitz.open(str(pdf_path))
    except Exception:
        return records
    with doc:
        for page in doc:
            page_text = page.get_text("text") or ""
            if not re.search(r"Identification", page_text, re.I):
                continue
            words = page.get_text("words")
            if not words:
                continue
            # Locate the identification title/header area.
            title_ys = [w[1] for w in words if re.search(r"Identification", w[4], re.I)]
            if not title_ys:
                continue
            min_y = min(title_ys) - 5
            # The horizontal row format uses a left phone-number column. Candidate rows
            # start with a phone token on the left side, after the header.
            candidates = []
            for w in words:
                x0, y0, x1, y1, word, *_ = w
                n = normalize_phone(word)
                if not n:
                    continue
                if y0 < min_y + 20:
                    continue
                # Restrict to the left identification number column; this avoids
                # CNI/IMEI/date/frequency numbers elsewhere in the same page.
                if x0 <= 95:
                    candidates.append((y0, x0, n))
            # Keep the first phone token per visual row.
            row_starts: list[tuple[float, str]] = []
            for y, _x, n in sorted(candidates):
                if row_starts and abs(y - row_starts[-1][0]) <= 6:
                    continue
                row_starts.append((y, n))
            for y, n in row_starts:
                # One identification row normally fits in about 16-20 PDF points;
                # continuation lines for name/address sit immediately below it.
                band = [w for w in words if y - 3 <= w[1] <= y + 19]
                if not band:
                    continue
                name_tokens = [w[4] for w in sorted(band, key=lambda z: (z[1], z[0])) if 90 <= w[0] < 205 and _looks_like_name_token(w[4])]
                cni_tokens = [w[4] for w in sorted(band, key=lambda z: (z[1], z[0])) if 255 <= w[0] < 315 and re.search(r"\d", w[4])]
                address_tokens = [w[4] for w in sorted(band, key=lambda z: (z[1], z[0])) if 365 <= w[0] < 505 and not re.search(r"REQ-GN|20777", w[4], re.I)]
                name = _clean_ident_name(" ".join(name_tokens))
                cni = clean_null("".join(cni_tokens))
                address = _clean_ident_address(" ".join(address_tokens))
                # Avoid saving pure technical/unknown rows unless they still carry useful address.
                if not name and not address and not cni:
                    continue
                if name and name.lower() == "inconnu" and not address and not cni:
                    continue
                if name and re.search(r"fr[ée]quence|cellule|correspondant|total|entrant|sortant|identit[ée]", name, re.I) and not address and not cni:
                    continue
                old = records.get(n)
                if old:
                    # Prefer the richest occurrence for a number.
                    if (not old.name or old.name.lower() == "inconnu") and name:
                        old.name = name
                    if not old.cni and cni:
                        old.cni = cni
                    if not old.address and address:
                        old.address = address
                else:
                    records[n] = IdentificationRecord(n, name, cni, address)
    return records


def parse_identification_map_from_text(text: str) -> dict[str, IdentificationRecord]:
    """Fallback parser for identification data when coordinates are unavailable.

    Handles plain-text sequences such as:
    600000001 | SAMPLE NAME | 2000-01-01 | time | CNI part | CNI part | ExpDate | time | SAMPLE ADDRESS
    """
    records: dict[str, IdentificationRecord] = {}
    m = re.search(r"III\.\s*Identification(?:\s+des\s+num[ée]ros)?", text, flags=re.I)
    if not m:
        return records
    section = text[m.end():]
    stop = re.search(r"\n\s*(?:IV\.\s*Fr[ée]quence|V\.\s*Fr[ée]quence|VI\.\s*Fr[ée]quence|VII\.\s*Fr[ée]quence|Agence\s+Nationale)", section, flags=re.I)
    if stop:
        section = section[: stop.start()]
    tokens = [clean_space(t) for t in section.splitlines()]
    tokens = [t for t in tokens if t and not t.startswith("REQ-GN-") and t not in {"RE", "R", "77", "20777"}]
    # Remove common header tokens.
    header_tokens = {"Numero", "Numéro", "Appelant", "Nom et prénom", "Nom et prenom", "Date", "Naissance", "Numero CNI", "Numéro CNI", "Exp CNI", "Adresse", "Nationalité"}
    i = 0
    while i < len(tokens):
        n = normalize_phone(tokens[i])
        if not n:
            i += 1
            continue
        # Avoid route/service fragments before actual records.
        if tokens[i] in header_tokens:
            i += 1
            continue
        j = i + 1
        name_parts: list[str] = []
        while j < len(tokens) and not DATE_WORD_RE.match(tokens[j]) and not normalize_phone(tokens[j]) and len(name_parts) < 6:
            if tokens[j] not in header_tokens and _looks_like_name_token(tokens[j]):
                name_parts.append(tokens[j])
            j += 1
        # Need a birth-date marker to treat this as a real identification row.
        if j >= len(tokens) or not DATE_WORD_RE.match(tokens[j]):
            i += 1
            continue
        # Skip birth date/time.
        j += 1
        if j < len(tokens) and TIME_WORD_RE.match(tokens[j]):
            j += 1
        cni_parts: list[str] = []
        while j < len(tokens) and not DATE_WORD_RE.match(tokens[j]) and len(cni_parts) < 3:
            if re.search(r"\d", tokens[j]) and not normalize_phone(tokens[j]):
                cni_parts.append(tokens[j])
            j += 1
        # Skip CNI expiration date/time.
        if j < len(tokens) and DATE_WORD_RE.match(tokens[j]):
            j += 1
        if j < len(tokens) and TIME_WORD_RE.match(tokens[j]):
            j += 1
        addr_parts: list[str] = []
        while j < len(tokens) and not normalize_phone(tokens[j]) and not DATE_WORD_RE.match(tokens[j]) and len(addr_parts) < 8:
            tok = tokens[j]
            if tok not in header_tokens and not _is_hexish_identifier(tok) and not re.search(r"REQ-GN|20777", tok, re.I):
                if tok.lower() != "inconnu":
                    addr_parts.append(tok)
            j += 1
        name = _clean_ident_name(" ".join(name_parts))
        cni = clean_null("".join(cni_parts))
        address = _clean_ident_address(" ".join(addr_parts))
        if name and name.lower() == "inconnu" and not address and not cni:
            i = max(j, i + 1)
            continue
        if name and re.search(r"fr[ée]quence|cellule|correspondant|total|entrant|sortant|identit[ée]", name, re.I) and not address and not cni:
            i = max(j, i + 1)
            continue
        if name or address or cni:
            old = records.get(n)
            if old:
                if (not old.name or old.name.lower() == "inconnu") and name:
                    old.name = name
                if not old.cni and cni:
                    old.cni = cni
                if not old.address and address:
                    old.address = address
            else:
                records[n] = IdentificationRecord(n, name, cni, address)
        i = max(j, i + 1)
    return records


def parse_identification_map(pdf_path: str | Path, text: str) -> dict[str, IdentificationRecord]:
    records = parse_identification_map_from_words(pdf_path)
    fallback = parse_identification_map_from_text(text)
    for n, rec in fallback.items():
        old = records.get(n)
        if old:
            if (not old.name or old.name.lower() == "inconnu") and rec.name:
                old.name = rec.name
            if not old.cni and rec.cni:
                old.cni = rec.cni
            if not old.address and rec.address:
                old.address = rec.address
        else:
            records[n] = rec
    return records

def _looks_like_next_row(lines: list[str], index: int) -> bool:
    if index >= len(lines):
        return False
    if not PHONE_TOKEN_RE.match(lines[index]):
        return False
    for j in range(index + 2, min(index + 14, len(lines))):
        if DATE_RE.match(lines[j]):
            return True
    return False


def _is_duration_token(value: str) -> bool:
    v = clean_space(value)
    return bool(CALL_DURATION_RE.match(v) or SMS_DURATION_RE.match(v))


def _make_row(caller_raw: str, called_raw: str, location_raw: str, imei: str, date_text: str, duration: str, source: str) -> CallRow | None:
    caller = normalize_phone(caller_raw)
    called = normalize_phone(called_raw)
    try:
        dt = datetime.strptime(date_text, "%d/%m/%Y %H:%M:%S")
    except ValueError:
        try:
            date_part, time_part = date_text.split()
            h, m, s = time_part.split(":")
            dt = datetime.strptime(f"{date_part} {int(h):02d}:{int(m):02d}:{int(s):02d}", "%d/%m/%Y %H:%M:%S")
        except Exception:
            return None
    location_clean, district, city, lat, lon, cell = parse_location(location_raw)
    return CallRow(
        caller_raw=clean_space(caller_raw), caller=caller,
        called_raw=clean_space(called_raw), called=called,
        location_raw=clean_space(location_raw), location_clean=location_clean,
        district=district, city=city, lat=lat, lon=lon, cell=cell,
        imei=clean_space(imei), dt=dt, duration=clean_space(duration), source=source,
    )


def parse_call_rows_from_lines(text: str) -> list[CallRow]:
    text = extract_text_until_tables_end(text)
    raw_lines = [clean_space(line) for line in text.splitlines()]
    noise_values = {"RE", "R", "77", "20777", "|", ""}
    lines = [
        line
        for line in raw_lines
        if line and not line.startswith("REQ-GN-") and line not in noise_values
    ]
    rows: list[CallRow] = []
    header_noise = (
        "Nom et Prénom", "Nom et Prenom", "Opérateur téléphonique", "Operateur telephonique", "Numéro CNI", "Numero CNI", "Adresse :",
        "Numéro appelant", "Numero appelant", "Numéro appellant", "Numero appellant", "Localisation numéro appelant", "Localisation numero appelant",
        "Date début", "Date debut", "Durée de", "Duree de", "Numéro appelé", "Numero appelé", "Numero appele", "I. Identité", "I. Identite", "II. Listing",
    )
    i = 0
    while i < len(lines):
        line = lines[i]
        if not PHONE_TOKEN_RE.match(line):
            i += 1
            continue
        date_idx = None
        for j in range(i + 2, min(i + 30, len(lines))):
            if DATE_RE.match(lines[j]):
                date_idx = j
                break
        if date_idx is None or date_idx + 1 >= len(lines):
            i += 1
            continue
        duration_line = lines[date_idx + 1]
        if not _is_duration_token(duration_line):
            i += 1
            continue
        imei = lines[date_idx - 1]
        location_lines = lines[i + 1:date_idx - 1]
        location_raw = clean_space(" ".join(location_lines))
        if any(marker.lower() in location_raw.lower() for marker in header_noise):
            i += 1
            continue
        if "cell" not in location_raw.lower() and "site inconnu" not in location_raw.lower():
            i += 1
            continue
        called_parts: list[str] = []
        k = date_idx + 2
        if k < len(lines):
            called_parts.append(lines[k])
            k += 1
            while k < len(lines) and len(called_parts) < 3 and not _looks_like_next_row(lines, k):
                if re.fullmatch(r"[A-Za-z0-9]+", lines[k]) and len("".join(called_parts)) < 28:
                    called_parts.append(lines[k])
                    k += 1
                else:
                    break
        called_raw = clean_space("".join(called_parts))
        row = _make_row(line, called_raw, location_raw, imei, lines[date_idx], duration_line, "CALL_OR_INLINE_SMS")
        if row:
            rows.append(row)
            i = max(k, i + 1)
        else:
            i += 1
    return rows


def _looks_like_sms_sender_token(value: str) -> bool:
    v = clean_space(value)
    if not v:
        return False
    if DATE_RE.match(v) or _is_duration_token(v):
        return False
    if re.search(r"localisation|destination|imei|date sms|num[eé]ro|numero|listing|identification", v, re.I):
        return False
    # sender can be a phone, short code, or hex-like service identifier
    return bool(re.fullmatch(r"[A-Za-z0-9]{1,40}", v))


def parse_sms_destination_rows_from_lines(text: str) -> list[CallRow]:
    """Parse separate 'B. Listing SMS' tables where the location belongs to the destination number."""
    m = re.search(r"B\.\s*Listing\s+SMS", text, flags=re.I)
    if not m:
        return []
    section = text[m.end():]
    stop = STOP_SECTION_RE.search(section)
    if stop:
        section = section[: stop.start()]
    raw_lines = [clean_space(line) for line in section.splitlines()]
    lines = [
        line
        for line in raw_lines
        if line
        and not line.startswith("REQ-GN-")
        and line not in {"RE", "R", "77", "20777"}
    ]
    rows: list[CallRow] = []
    date_indices = [i for i, line in enumerate(lines) if DATE_RE.match(line)]
    for date_idx in date_indices:
        if date_idx < 3 or date_idx + 1 >= len(lines):
            continue
        imei = lines[date_idx - 1]
        destination_raw = lines[date_idx + 1]
        dest = normalize_phone(destination_raw)
        if not dest:
            continue
        # Walk backwards from the IMEI to find the beginning of the location block.
        loc_start = date_idx - 2
        while loc_start > 0:
            prev = lines[loc_start - 1]
            if _looks_like_sms_sender_token(prev):
                break
            if DATE_RE.match(prev):
                break
            loc_start -= 1
        if loc_start <= 0:
            continue
        sender_idx = loc_start - 1
        sender_parts = [lines[sender_idx]]
        # Some sender IDs are split across two lines, e.g. '466972...E6' + 'B'.
        if len(sender_parts[0]) <= 4 and sender_idx - 1 >= 0:
            prev = lines[sender_idx - 1]
            if _looks_like_sms_sender_token(prev) and not normalize_phone(prev):
                sender_parts.insert(0, prev)
                sender_idx -= 1
        sender_raw = clean_space("".join(sender_parts))
        location_raw = clean_space(" ".join(lines[loc_start:date_idx - 1]))
        if "cell" not in location_raw.lower() and "site inconnu" not in location_raw.lower():
            continue
        row = _make_row(destination_raw, sender_raw, location_raw, imei, lines[date_idx], "SMS/DATA", "SMS_DESTINATION")
        if row:
            rows.append(row)
    return rows

def parse_rows_from_words(pdf_path: str | Path) -> list[CallRow]:
    rows: list[CallRow] = []
    with fitz.open(str(pdf_path)) as doc:
        for page in doc:
            words = page.get_text("words")
            if not words:
                continue
            # Drop page headers/footers; keep table body.
            words = [w for w in words if 35 <= w[0] <= page.rect.width - 20]
            words_sorted = sorted(words, key=lambda w: (w[1], w[0]))
            date_anchors: list[tuple[float, str, float]] = []
            for idx, w in enumerate(words_sorted[:-1]):
                if DATE_WORD_RE.match(w[4]) and TIME_WORD_RE.match(words_sorted[idx + 1][4]):
                    date_anchors.append((w[1], f"{w[4]} {words_sorted[idx + 1][4]}", w[0]))
            if not date_anchors:
                continue
            date_anchors.sort()
            ys = [y for y, _, _ in date_anchors]
            for pos, (y, date_text, _date_x) in enumerate(date_anchors):
                # Ensure this is a table date, not a header/date in a frequency section: require a left participant and a location/cell.
                same_line = [w for w in words_sorted if abs(w[1] - y) <= 4.2]
                caller_tokens = [w[4] for w in same_line if 35 <= w[0] <= 105 and (PHONE_TOKEN_RE.match(w[4]) or normalize_phone(w[4]))]
                if not caller_tokens:
                    continue
                caller_raw = "".join(caller_tokens[:2])
                imei_tokens = [w[4] for w in same_line if 260 <= w[0] <= 335 and (IMEI_RE.match(w[4]) or re.match(r"^[A-Za-z0-9]{10,}$", w[4]))]
                imei = "".join(imei_tokens[:2])
                duration_tokens = [w[4] for w in same_line if 395 <= w[0] <= 475 and _is_duration_token(w[4])]
                duration = clean_space(" ".join(duration_tokens)) if duration_tokens else "SMS/DATA"
                # The certified ReportLab/Orange export places the destination
                # number around x=444 and the textual role (``CIBLE`` or
                # ``CORRESPONDANT``) after x=500.  The previous fixed x>=470
                # rule therefore discarded the number and parsed the role as
                # ``called_raw``.  Restrict the primary candidates to tokens
                # that normalize as phone numbers, anywhere to the right of
                # the duration column.  Keep the legacy range only as a
                # compatibility fallback for older column layouts.
                called_phone_tokens = [
                    w[4]
                    for w in same_line
                    if w[0] >= 430 and normalize_phone(w[4])
                ]
                if called_phone_tokens:
                    called_raw = "".join(called_phone_tokens[:2])
                else:
                    # Preserve a numeric network/service identifier when it is
                    # not a normalizable mobile number.  It must not be
                    # replaced by the following role label.
                    called_numeric_tokens = [
                        w[4]
                        for w in same_line
                        if w[0] >= 430 and re.fullmatch(r"\+?\d{6,20}", w[4])
                    ]
                    called_tokens = called_numeric_tokens or [
                        w[4]
                        for w in same_line
                        if w[0] >= 470
                        and not DATE_WORD_RE.match(w[4])
                        and not TIME_WORD_RE.match(w[4])
                    ]
                    called_raw = "".join(called_tokens[:3])

                next_y = ys[pos + 1] if pos + 1 < len(ys) else y + 24
                lower = max(y + 9, min(next_y - 8, y + 16))
                loc_words = [w for w in words_sorted if 95 <= w[0] <= 270 and (y - 8) <= w[1] < lower]
                location_raw = clean_space(" ".join(w[4] for w in sorted(loc_words, key=lambda w: (w[1], w[0]))))
                if "cell" not in location_raw.lower() and "site inconnu" not in location_raw.lower():
                    continue
                row = _make_row(caller_raw, called_raw, location_raw, imei, date_text, duration, "COORDINATE_TABLE")
                if row:
                    rows.append(row)
    return rows


def dedupe_rows(rows: Iterable[CallRow]) -> list[CallRow]:
    seen = set()
    out: list[CallRow] = []
    for r in rows:
        key = (
            r.caller or r.caller_raw,
            r.called or r.called_raw,
            r.dt.strftime("%Y-%m-%d %H:%M:%S"),
            r.duration,
            r.cell,
            r.imei,
            r.district,
        )
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    out.sort(key=lambda r: r.dt, reverse=True)
    return out


def parse_listing_rows(text: str, pdf_path: str | Path | None = None) -> tuple[list[CallRow], str]:
    line_rows = parse_call_rows_from_lines(text)
    sms_rows = parse_sms_destination_rows_from_lines(text)
    rows_no_coords = [*line_rows, *sms_rows]
    date_count = len(re.findall(r"\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}", text))
    word_rows: list[CallRow] = []
    used_coords = False
    # Use coordinate extraction as a fallback for column-stacked PDFs, not as a duplicate source.
    if pdf_path and (not rows_no_coords or (date_count and len(rows_no_coords) < max(3, int(date_count * 0.35)))):
        word_rows = parse_rows_from_words(pdf_path)
        if len(word_rows) > len(rows_no_coords):
            rows = word_rows
            used_coords = True
        else:
            rows = rows_no_coords
    else:
        rows = rows_no_coords
    parser_desc = f"hybrid lines={len(line_rows)}, sms={len(sms_rows)}, coords={len(word_rows)}, used_coords={used_coords}"
    return rows, parser_desc


def parse_location(location: str) -> tuple[str, str, str, str, str, str]:
    loc = clean_space(location)
    cell = _first_regex(r"Cell\s*:?\s*([^\s\)]+)", loc, flags=re.I)
    lon = _first_regex(r"Long\s*:?\s*([^\s\)]+)", loc, flags=re.I)
    lat = _first_regex(r"Lat\s*:?\s*([^\s\)]+)", loc, flags=re.I)
    before_cell = clean_space(re.sub(r"\(?\s*Cell\s*:?\s*.*", "", loc, flags=re.I))
    if not before_cell:
        before_cell = loc
    if "site inconnu" in before_cell.lower() or re.match(r"^cell\s*:", before_cell, re.I):
        return loc, "Site Inconnu", "N/D", lat, lon, cell
    district, city = split_district_city(before_cell)
    return loc, district, city, lat, lon, cell


def split_district_city(before_cell: str) -> tuple[str, str]:
    s = clean_space(before_cell)
    # Normalize odd spaces inserted by PDF extraction around hyphens.
    s = re.sub(r"\s*-\s*", "-", s)
    parts = [p.strip() for p in s.split("-") if p.strip()]
    if not parts:
        return s, "N/D"
    last = parts[-1]
    tokens = last.split()
    if len(tokens) >= 2 and tokens[-1].upper() in REGIONS:
        city_candidate = " ".join(tokens[:-1])
        region = tokens[-1].upper()
        if len(parts) >= 3:
            return "-".join(parts[:-1]), city_candidate or "N/D"
        readable_without_region = clean_space(re.sub(rf"\s+{re.escape(region)}$", "", s, flags=re.I))
        return readable_without_region or s, "N/D"
    if last.upper() in REGIONS and len(parts) >= 2:
        return "-".join(parts[:-1]), parts[-2]
    if len(parts) >= 2:
        return "-".join(parts[:-1]), parts[-1]
    return s, "N/D"


def parse_windev_location(location: str) -> tuple[str, str, str, str, str, str]:
    """Parse a location from the WinDev requisition tables.

    WinDev exports the value as ``site-ville-département région(cell: ...)``.
    The generic parser historically treated the department as the city and also
    considered bare cell identifiers to be real districts. This specialized
    parser corrects both cases without changing the generic location logic.
    """
    loc = clean_space(location)
    cell = _first_regex(r"Cell\s*:?\s*([^\s\)]+)", loc, flags=re.I)
    lon = _first_regex(r"Long\s*:?\s*([^\s\)]+)", loc, flags=re.I)
    lat = _first_regex(r"Lat\s*:?\s*([^\s\)]+)", loc, flags=re.I)
    lon = "" if lon in {"-", "N/D"} else lon
    lat = "" if lat in {"-", "N/D"} else lat
    before_cell = clean_space(re.sub(r"\(?\s*Cell\s*:?\s*.*", "", loc, flags=re.I))
    before_cell = re.sub(r"\s*-\s*$", "", before_cell).strip()

    # Examples such as "624-01-32-15955 -(cell: ...)" contain no named site.
    if (
        not before_cell
        or "site inconnu" in before_cell.lower()
        or re.fullmatch(r"\d{3}(?:-\d+){2,}", before_cell)
    ):
        return loc, "Site Inconnu", "N/D", lat, lon, cell

    protected = before_cell
    placeholders: dict[str, str] = {}
    for index, region in enumerate(sorted(REGIONS, key=len, reverse=True)):
        if "-" not in region:
            continue
        token = f"__REGION_{index}__"
        replaced = re.sub(re.escape(region), token, protected, flags=re.I)
        if replaced != protected:
            protected = replaced
            placeholders[token] = region
    protected = re.sub(r"\s*-\s*", "-", protected)
    parts = [part.strip() for part in protected.split("-") if part.strip()]
    parts = [next((value for token, value in placeholders.items() if token in part), part) if part in placeholders else part for part in parts]
    # Restore placeholders that can be embedded after a department name.
    for i, _part in enumerate(parts):
        for token, value in placeholders.items():
            parts[i] = parts[i].replace(token, value)

    if len(parts) >= 3:
        last_tokens = parts[-1].split()
        if last_tokens and last_tokens[-1].upper() in REGIONS:
            district = clean_space("-".join(parts[:-2]))
            city = clean_space(parts[-2])
            if district and city:
                return loc, district, city, lat, lon, cell

    district, city = split_district_city(before_cell)
    return loc, district, city, lat, lon, cell


def _windev_table_rows(page: object) -> list[list[str]]:
    """Extract all grid tables on a page while tolerating empty/malformed cells."""
    extracted: list[list[str]] = []
    try:
        finder = page.find_tables()
        tables = getattr(finder, "tables", []) or []
    except Exception:
        return extracted
    for table in tables:
        try:
            rows = table.extract() or []
        except Exception:
            continue
        for row in rows:
            if not row:
                continue
            extracted.append([clean_space(cell or "") for cell in row])
    return extracted


def _apply_windev_location(row: CallRow) -> CallRow:
    location_clean, district, city, lat, lon, cell = parse_windev_location(
        row.location_raw
    )
    row.location_clean = location_clean
    row.district = district
    row.city = city
    row.lat = lat
    row.lon = lon
    row.cell = cell
    return row


def _parse_windev_rows_from_all_tables(pdf_path: str | Path) -> list[CallRow]:
    """Slow fallback used only when coordinate extraction is incomplete."""
    rows: list[CallRow] = []
    with fitz.open(str(pdf_path)) as doc:
        for page in doc:
            for cells in _windev_table_rows(page):
                if len(cells) != 6 or not DATE_RE.match(cells[3]):
                    continue
                row = _make_row(
                    cells[0], cells[5], cells[1], cells[2], cells[3], cells[4],
                    "WINDEV_TABLE_FALLBACK",
                )
                if row:
                    rows.append(_apply_windev_location(row))
    rows.sort(key=lambda row: row.dt, reverse=True)
    return rows


def _read_windev_identity_row(
    cells: list[str],
    info: SubscriberInfo,
    identities: dict[str, IdentificationRecord],
) -> None:
    folded = [_fold_windev(cell) for cell in cells]
    if len(cells) == 4:
        if folded[0] == "telephone":
            info.number = normalize_phone(cells[1]) or info.number
            if folded[2].startswith("nom et prenom"):
                info.name = clean_null(cells[3]) or info.name
        elif folded[0].startswith("operateur telephonique"):
            info.operator = clean_null(cells[1]) or info.operator
            if folded[2].startswith("numero cni"):
                info.cni = clean_null(cells[3]) or info.cni
        elif folded[0] == "imei":
            info.imei = clean_null(cells[1]) or info.imei
            if folded[2] == "adresse":
                info.address = clean_null(cells[3]) or info.address
        return

    if len(cells) != 7:
        return
    number = normalize_phone(cells[0])
    if not number or folded[0] == "numero appelant":
        return
    name = clean_null(cells[1])
    cni = clean_null(cells[3])
    address = clean_null(cells[5])
    if name or cni or address:
        identities[number] = IdentificationRecord(number, name, cni, address)


def parse_windev_requisition(pdf_path: str | Path, text: str) -> WinDevParseResult:
    """Extract WinDev PDFs with a fast coordinate/table hybrid.

    Event rows are read from positioned words, which is typically tens of times
    faster than running grid detection on every page. Grid detection is limited
    to the main identity and identification sections. A full table scan remains
    available only when the fast event count fails its integrity check.
    """
    info = SubscriberInfo()
    identities: dict[str, IdentificationRecord] = {}

    flat = clean_space(text)
    period = _first_regex(r"P[ée]riode\s*:\s*du\s*([0-9/]+\s*(?:au|Au)\s*[0-9/]+)", flat)
    if period:
        period = re.sub(r"\s+[Aa]u\s+", " au ", period)
    info.period = clean_null(period)
    info.requester = clean_null(_first_regex(
        r"Service\s+Demandeur\s*:?\s*(.*?)(?:\s+P[ée]riode\s*:|\s+I\.\s*Ident|\s+T[ée]l[ée]phone\s*:|\s+REPUBLIQUE)",
        flat,
    ))

    # Fast event extraction preserves all repeated calls/SMS and exact coordinates.
    rows = [_apply_windev_location(row) for row in parse_rows_from_words(pdf_path)]
    for row in rows:
        row.source = "WINDEV_COORDINATE"
    rows.sort(key=lambda row: row.dt, reverse=True)

    # Only two small sections need expensive table-grid detection.
    identification_mode = False
    with fitz.open(str(pdf_path)) as doc:
        for page_index, page in enumerate(doc):
            page_folded = _fold_windev(page.get_text("text") or "")
            if "identification des numeros" in page_folded:
                identification_mode = True
            scan_page = page_index == 0 or identification_mode
            if scan_page:
                for cells in _windev_table_rows(page):
                    _read_windev_identity_row(cells, info, identities)
            if identification_mode and "frequence par cellule" in page_folded:
                identification_mode = False

    # Fill safe gaps without allowing the generic flat-text parser to pollute
    # already structured values such as CNI or address.
    word_info = parse_identity_from_words(pdf_path)
    for attribute in ("number", "name", "operator", "cni", "address"):
        if not getattr(info, attribute) and getattr(word_info, attribute):
            setattr(info, attribute, getattr(word_info, attribute))

    listing_text = extract_text_until_tables_end(text)
    expected_events = len(re.findall(
        r"(?m)^\s*\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s*$",
        listing_text,
    ))
    coverage = (len(rows) / expected_events) if expected_events else 1.0
    used_full_table_fallback = False
    if expected_events and not 0.90 <= coverage <= 1.10:
        fallback_rows = _parse_windev_rows_from_all_tables(pdf_path)
        if len(fallback_rows) > len(rows):
            rows = fallback_rows
            coverage = len(rows) / expected_events
            used_full_table_fallback = True

    parser_desc = (
        f"windev_hybrid rows={len(rows)}, identities={len(identities)}, "
        f"coverage={coverage:.3f}, full_table_fallback={used_full_table_fallback}"
    )
    return WinDevParseResult(info, rows, identities, parser_desc)


def parse_pdfkit_mtn_location(location: str) -> tuple[str, str, str, str, str, str]:
    """Normalize an MTN/PDFKit site while preserving cell and GPS metadata."""
    loc = clean_space((location or "").replace("|", " "))
    if not loc or loc == "-":
        return loc, "Site Inconnu", "N/D", "", "", ""
    cell = _first_regex(r"Cell\s*:?\s*([^\s\)]+)", loc, flags=re.I)
    lon = _first_regex(r"Long\s*:?\s*([^\s\)]+)", loc, flags=re.I)
    lat = _first_regex(r"Lat\s*:?\s*([^\s\)]+)", loc, flags=re.I)
    lon = "" if lon in {"-", "N/D"} else lon
    lat = "" if lat in {"-", "N/D"} else lat
    before_cell = clean_space(re.sub(r"\(?\s*Cell\s*:?.*", "", loc, flags=re.I))
    if not before_cell or "site inconnu" in before_cell.lower():
        return loc, "Site Inconnu", "N/D", lat, lon, cell

    # PDFKit MTN locations generally follow:
    # SITE-DEPARTEMENT, VILLE(ARRONDISSEMENT)-REGION.
    without_region = re.sub(
        r"\s*-\s*(?:ADAMAOUA|CENTRE|EST|EAST|EXTREME-NORD|FAR-NORTH|LITTORAL|NORD|NORTH|NORD-OUEST|NORTH-WEST|OUEST|WEST|SUD|SOUTH|SUD-OUEST|SOUTH-WEST)\s*$",
        "",
        before_cell,
        flags=re.I,
    )
    district = without_region
    city = "N/D"
    if "," in without_region:
        left, right = without_region.rsplit(",", 1)
        district = clean_space(left.split("-", 1)[0]) or clean_space(left)
        city = clean_space(re.sub(r"\(.*", "", right)) or "N/D"
        city = re.sub(r"\s+(?:I|II|III|IV|V|VI|VII|VIII|IX|X|\d+)\s*$", "", city, flags=re.I).strip() or "N/D"
    if not district or re.fullmatch(r"\d+(?:-\d+){2,}", district):
        district = "Site Inconnu"
        city = "N/D"
    return loc, district, city, lat, lon, cell


def _apply_pdfkit_mtn_location(row: CallRow) -> CallRow:
    location_clean, district, city, lat, lon, cell = parse_pdfkit_mtn_location(
        row.location_raw
    )
    row.location_clean = location_clean
    row.district = district
    row.city = city
    row.lat = lat
    row.lon = lon
    row.cell = cell
    return row


def _pdfkit_tables(page: object) -> list[list[list[str]]]:
    """Return clean table matrices for one PDFKit page."""
    matrices: list[list[list[str]]] = []
    try:
        tables = getattr(page.find_tables(), "tables", []) or []
    except Exception:
        return matrices
    for table in tables:
        try:
            extracted = table.extract() or []
        except Exception:
            continue
        matrix = [
            [clean_space(cell or "") for cell in row]
            for row in extracted
            if row
        ]
        if matrix:
            matrices.append(matrix)
    return matrices


def _pdfkit_header(matrix: list[list[str]]) -> str:
    return _fold_windev(" ".join(matrix[0])) if matrix else ""


def _pdfkit_mtn_read_subscriber_table(matrix: list[list[str]], info: SubscriberInfo) -> None:
    for cells in matrix:
        if len(cells) != 4:
            continue
        left = _fold_windev(cells[0])
        right = _fold_windev(cells[2])
        if left == "telephone":
            info.number = normalize_phone(cells[1]) or info.number
        elif left == "operateur":
            info.operator = clean_null(cells[1]) or info.operator
        elif left == "imei":
            info.imei = clean_null(cells[1]) or info.imei
        if right.startswith("nom et prenom"):
            info.name = clean_null(cells[3]) or info.name
        elif right.startswith("numero cni"):
            info.cni = clean_null(cells[3]) or info.cni
        elif right == "adresse":
            info.address = clean_null(cells[3]) or info.address


def _pdfkit_mtn_read_identification_row(
    cells: list[str], identities: dict[str, IdentificationRecord]
) -> None:
    if len(cells) < 6:
        return
    number = normalize_phone(cells[0])
    if not number:
        return
    name = clean_null(cells[1])
    cni = clean_null(cells[3])
    expiration_and_overflow = clean_space(cells[4])
    address = clean_null(cells[5])
    overflow = re.sub(r"^\d{4}-\d{2}-\d{2}\s*", "", expiration_and_overflow)
    if overflow:
        address = clean_null(f"{overflow} {address}")
    if name == "-":
        name = ""
    if cni == "-":
        cni = ""
    if address == "-":
        address = ""
    if name or cni or address:
        identities[number] = IdentificationRecord(number, name, cni, address)


def parse_pdfkit_mtn_requisition(
    pdf_path: str | Path, text: str
) -> PdfKitMtnParseResult:
    """Parse the MTN/PDFKit call, SMS and identification grids.

    Only pages containing the recognized grids are scanned. Later frequency and
    certificate tables are intentionally ignored, both for speed and to avoid
    treating aggregates as call-detail records.
    """
    info = SubscriberInfo()
    identities: dict[str, IdentificationRecord] = {}
    rows: list[CallRow] = []
    call_count = 0
    sms_count = 0

    flat = clean_space(text)
    info.period = clean_null(_first_regex(
        r"P[ée]riode\s*:\s*du\s*([0-9/]+\s*(?:au|Au)\s*[0-9/]+)", flat
    ))
    info.period = re.sub(r"\s+[Aa]u\s+", " au ", info.period)
    info.requester = clean_null(_first_regex(
        r"Service\s+Demandeur\s*:?\s*(.*?)(?:\s+I\.\s*Identit[ée]|\s+T[ée]l[ée]phone\s*:)",
        flat,
    ))

    with fitz.open(str(pdf_path)) as doc:
        for page in doc:
            page_folded = _fold_windev(page.get_text("text") or "")
            if "v frequence par cellule" in page_folded:
                break
            for matrix in _pdfkit_tables(page):
                header = _pdfkit_header(matrix)
                if "telephone" in header and "nom et prenom" in header and len(matrix[0]) == 4:
                    _pdfkit_mtn_read_subscriber_table(matrix, info)
                    continue
                if all(marker in header for marker in (
                    "numero appellant", "imei appellant", "localisation appellant",
                    "numero appele", "date debut appel", "duree",
                )):
                    for cells in matrix[1:]:
                        if len(cells) < 8:
                            continue
                        date_text = clean_space(cells[6])
                        row = _make_row(
                            cells[0], cells[3], cells[2], cells[1], date_text,
                            cells[7], "PDFKIT_MTN_CALL",
                        )
                        if row:
                            rows.append(_apply_pdfkit_mtn_location(row))
                            call_count += 1
                    continue
                if all(marker in header for marker in (
                    "numero d envoi", "localisation numero destination",
                    "imei numero destination", "date sms", "numero destination",
                )):
                    for cells in matrix[1:]:
                        if len(cells) < 5:
                            continue
                        destination = cells[4]
                        if not normalize_phone(destination):
                            continue
                        imei = _first_regex(r"(\d{14,17})", cells[2]) or cells[2]
                        row = _make_row(
                            destination, cells[0].replace("|", " "), cells[1], imei,
                            clean_space(cells[3]), "SMS/DATA", "PDFKIT_MTN_SMS",
                        )
                        if row:
                            rows.append(_apply_pdfkit_mtn_location(row))
                            sms_count += 1
                    continue
                if all(marker in header for marker in (
                    "numero", "nom prenom", "date de naissance", "numero cni",
                    "date expiration cni", "quartier",
                )):
                    for cells in matrix[1:]:
                        _pdfkit_mtn_read_identification_row(cells, identities)

    rows.sort(key=lambda row: row.dt, reverse=True)
    expected_text = _first_regex(r"Nombre\s+d['’]enregistrements\s*:\s*(\d+)", flat)
    expected_events = int(expected_text) if expected_text.isdigit() else len(rows)
    coverage = len(rows) / expected_events if expected_events else 1.0
    parser_desc = (
        f"pdfkit_mtn_v1 rows={len(rows)}, calls={call_count}, sms={sms_count}, "
        f"identities={len(identities)}, coverage={coverage:.3f}"
    )
    return PdfKitMtnParseResult(
        info, rows, identities, parser_desc, expected_events
    )


def in_time_window(dt: datetime, start_hour: int, end_hour: int) -> bool:
    h = dt.hour
    if start_hour == end_hour:
        return True
    if start_hour < end_hour:
        return start_hour <= h < end_hour
    return h >= start_hour or h < end_hour


def choose_district(events: Iterable[CallRow], start_hour: int, end_hour: int) -> DistrictChoice:
    known = [r for r in events if r.is_known_site and r.district]
    if not known:
        return DistrictChoice()
    grouped: dict[str, dict[str, object]] = {}
    for r in known:
        key = r.district
        g = grouped.setdefault(key, {
            "district": r.district, "city": r.city, "lat": r.lat, "lon": r.lon,
            "total_count": 0, "total_days": set(), "window_count": 0, "window_days": set(),
        })
        g["total_count"] = int(g["total_count"]) + 1
        g["total_days"].add(r.dt.date())
        if in_time_window(r.dt, start_hour, end_hour):
            g["window_count"] = int(g["window_count"]) + 1
            g["window_days"].add(r.dt.date())

    def as_choice(g: dict[str, object], prefix_nd: bool, source: str) -> DistrictChoice:
        return DistrictChoice(
            district=str(g["district"]), city=str(g["city"]), lat=str(g.get("lat") or ""), lon=str(g.get("lon") or ""),
            prefix_nd=prefix_nd, source=source,
            total_count=int(g["total_count"]), total_days=len(g["total_days"]),
            window_count=int(g["window_count"]), window_days=len(g["window_days"]),
        )

    window_candidates = [g for g in grouped.values() if int(g["window_count"]) > 0]
    if window_candidates:
        best_window = max(window_candidates, key=lambda g: (len(g["window_days"]), int(g["window_count"]), len(g["total_days"]), int(g["total_count"])))
        if len(best_window["window_days"]) >= 2 or int(best_window["window_count"]) >= 3:
            return as_choice(best_window, False, f"{start_hour:02d}h-{end_hour:02d}h")
    best_global = max(grouped.values(), key=lambda g: (len(g["total_days"]), int(g["total_count"]), len(g["window_days"]), int(g["window_count"])))
    return as_choice(best_global, True, "hors plage / fenêtre insuffisante")


def count_occurrences_in_district(events: Iterable[CallRow], district: str) -> tuple[int, int]:
    calls = smsdata = 0
    for r in events:
        if r.district != district:
            continue
        if r.is_call:
            calls += 1
        else:
            smsdata += 1
    return calls, smsdata


def count_exchanges_between(
    rows: Iterable[CallRow],
    suspect_number: str,
    correspondent_number: str,
) -> tuple[int, int]:
    """Count every direct exchange between two numbers, split by event type.

    This intentionally does not depend on a PDF parser or a retained district.
    All parsers normalize their output to ``CallRow`` before report generation,
    so the same bidirectional rule is applied to legacy, WinDev and MTN/PDFKit
    listings without changing any parser-specific extraction logic.
    """
    calls = smsdata = 0
    for r in rows:
        same_exchange = (
            (r.caller == suspect_number and r.called == correspondent_number)
            or (r.caller == correspondent_number and r.called == suspect_number)
        )
        if not same_exchange:
            continue
        if r.is_call:
            calls += 1
        else:
            smsdata += 1
    return calls, smsdata


def dates_in_suspect_district(events: Iterable[CallRow], suspect_district: str, suspect_number: str) -> list[str]:
    # Sort real dates before formatting them. Sorting dd/mm strings incorrectly
    # placed 02/07 before 27/06 in reports spanning several months.
    dates = sorted({r.dt.date() for r in events if r.called == suspect_number and r.district == suspect_district})
    return [date.strftime("%d/%m") for date in dates]


def top_districts(events: Iterable[CallRow], limit: int = 3) -> list[tuple[str, str, int, int, int]]:
    grouped: dict[str, dict[str, object]] = {}
    for r in events:
        if not r.is_known_site:
            continue
        g = grouped.setdefault(r.district, {"city": r.city, "count": 0, "days": set()})
        g["count"] = int(g["count"]) + 1
        g["days"].add(r.dt.date())
    ranked = sorted(grouped.items(), key=lambda kv: (len(kv[1]["days"]), int(kv[1]["count"])), reverse=True)
    return [(dist, str(v["city"]), int(v["count"]), len(v["days"]), i + 1) for i, (dist, v) in enumerate(ranked[:limit])]


def format_dates(dates: list[str]) -> str:
    if not dates:
        return "NON"
    if len(dates) <= 5:
        return "OUI, le " + ", ".join(dates)
    return "OUI, du " + dates[0] + " au " + dates[-1] + f" ({len(dates)} jours)"


def duration_to_seconds(duration: str) -> int:
    value = clean_space(duration)
    if not CALL_DURATION_RE.match(value):
        return 0
    try:
        h, m, sec = [int(part) for part in value.split(":")]
    except Exception:
        return 0
    return h * 3600 + m * 60 + sec


def format_duration(seconds: int) -> str:
    if seconds <= 0:
        return "00:00:00"
    h, rem = divmod(int(seconds), 3600)
    m, sec = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{sec:02d}"


def total_call_duration_between(rows: Iterable[CallRow], suspect_number: str, correspondent_number: str) -> int:
    total = 0
    for r in rows:
        if not r.is_call:
            continue
        same_exchange = (
            (r.caller == suspect_number and r.called == correspondent_number)
            or (r.caller == correspondent_number and r.called == suspect_number)
        )
        if same_exchange:
            total += duration_to_seconds(r.duration)
    return total




def _phone_variants_for_search(phone_number: str) -> set[str]:
    """Return common digit-only forms of a Cameroonian correspondent number."""
    n = normalize_phone(phone_number) or re.sub(r"\D", "", phone_number or "")
    variants: set[str] = set()
    if not n or len(n) < 6:
        return variants
    variants.add(n)
    if len(n) == 9:
        variants.update({f"237{n}", f"00237{n}"})
    if n.startswith("237") and len(n) == 12:
        variants.add(n[3:])
        variants.add(f"00237{n[3:]}")
    if n.startswith("00237") and len(n) == 14:
        variants.add(n[5:])
        variants.add(f"237{n[5:]}")
    return {v for v in variants if len(v) >= 6}


def _url_contains_phone_number(link: str, phone_number: str) -> bool:
    """Return True when the correspondent phone appears inside the URL itself.

    The final decision is not based on this alone. A URL such as
    https://website.example/600000009 is suspicious, but a social-network result can
    still be kept when the title/snippet contains a real supporting signal.
    """
    variants = _phone_variants_for_search(phone_number)
    if not variants:
        return False
    parsed = urlparse(link or "")
    url_material = " ".join([
        parsed.netloc or "",
        parsed.path or "",
        parsed.params or "",
        parsed.query or "",
        parsed.fragment or "",
    ])
    url_digits = re.sub(r"\D", "", url_material)
    if not url_digits:
        return False
    return any(v in url_digits for v in variants)


def _visible_text_contains_phone(text: str, phone_number: str) -> bool:
    """Check title/snippet/display text, not URL, for the correspondent number."""
    variants = _phone_variants_for_search(phone_number)
    if not variants:
        return False
    digits = re.sub(r"\D", "", text or "")
    return any(v in digits for v in variants)


def _name_tokens(name: str) -> list[str]:
    raw = re.findall(r"[A-Za-zÀ-ÖØ-öø-ÿ]{3,}", clean_space(name).upper())
    stop = {
        "INCONNU", "IDENTIFIE", "IDENTIFIÉ", "NON", "NULL", "NUL", "NOM", "PRENOM",
        "LE", "LA", "LES", "DES", "DU", "DE", "ET", "M", "MME", "MR",
    }
    return [t for t in raw if t not in stop]


def _visible_text_supports_name(text: str, name: str) -> bool:
    """Return True when visible SerpApi metadata contains enough name evidence.

    Without AI, we avoid validating a link on a single common token. For names with
    two or more useful words, at least two tokens must be visible. For a short one-
    word name, the exact token can be enough.
    """
    tokens = _name_tokens(name)
    if not tokens:
        return False
    text_up = clean_space(text).upper()
    if not text_up:
        return False
    present = [t for t in tokens if re.search(rf"\b{re.escape(t)}\b", text_up)]
    if len(tokens) >= 2:
        return len(set(present)) >= 2
    return len(present) >= 1


def _serpapi_item_visible_text(item: dict) -> str:
    fields: list[str] = []
    for key in ("title", "snippet", "displayed_link", "source"):
        value = item.get(key)
        if isinstance(value, str):
            fields.append(value)
    rich = item.get("rich_snippet") or item.get("about_this_result") or {}
    if isinstance(rich, dict):
        fields.append(str(rich))
    return clean_space(" ".join(fields))


def _is_generic_phone_lookup_domain(host: str) -> bool:
    """Detect reverse-lookup/spam-directory pages that usually match only by number."""
    host = (host or "").lower()
    generic_parts = (
        "nomorobo", "mapquest", "bbb.org", "whitepages", "truecaller", "tellows",
        "whocallsme", "who-called", "callercenter", "numlookup", "spytox", "sync.me",
        "reversephone", "phonebook", "phonesbook", "annuaire-inverse", "spamcalls",
        "shouldianswer", "unknownphone", "findwhocallsyou", "usphonebook",
    )
    return any(part in host for part in generic_parts)


def _serpapi_domain_label(url: str) -> str:
    host = (urlparse(url).netloc or "").lower().replace("www.", "")
    if "facebook.com" in host or "fb.com" in host:
        return "Facebook"
    if "linkedin.com" in host:
        return "LinkedIn"
    if "scribd.com" in host:
        return "Scribd"
    if "instagram.com" in host:
        return "Instagram"
    if "x.com" in host or "twitter.com" in host:
        return "X/Twitter"
    if "tiktok.com" in host:
        return "TikTok"
    if "youtube.com" in host or "youtu.be" in host:
        return "YouTube"
    if "whatsapp.com" in host:
        return "WhatsApp"
    base = host.split(":", 1)[0]
    if base:
        return base.split(".")[0].capitalize()
    return "Autre"


def _serpapi_candidate_score(item: dict, link: str, number: str, name: str, preferred_domains: tuple[str, ...]) -> tuple[int, str] | None:
    """Score a SerpApi result. Return None when it is probably misleading.

    Correction v5: the anti-faux-positifs filter no longer deletes useful Facebook /
    social results only because a numeric value appears in the URL. It rejects the
    URL-number case only when the visible title/snippet does not also contain a
    credible signal: the exact phone number or at least two words from the name.
    """
    host = (urlparse(link).netloc or "").lower()
    visible_text = _serpapi_item_visible_text(item)
    url_has_number = _url_contains_phone_number(link, number)
    visible_has_number = _visible_text_contains_phone(visible_text, number)
    visible_has_name = _visible_text_supports_name(visible_text, name)
    is_preferred = any(domain in host for domain in preferred_domains)
    is_generic_phone_site = _is_generic_phone_lookup_domain(host)

    # Main fix: reject deceptive links where the phone exists only in the URL.
    if url_has_number and not (visible_has_number or visible_has_name):
        return None

    # Reject generic reverse lookup / spam directory pages unless the metadata also
    # names the correspondent. This removes results like BBB, Mapquest, Nomorobo.
    if is_generic_phone_site and not visible_has_name:
        return None

    # Do not keep very weak non-social results that have neither number nor name in metadata.
    if not is_preferred and not visible_has_number and not visible_has_name:
        return None

    if is_preferred and (visible_has_number or visible_has_name):
        return (0, link)
    if is_preferred:
        return (1, link)
    if visible_has_name:
        return (2, link)
    if visible_has_number:
        return (3, link)
    return (5, link)


def search_internet_for_correspondent(
    number: str,
    name: str = "",
    serpapi_api_key: str | None = None,
    max_results: int = 3,
    cache: dict[tuple[str, str, int], list[WebSearchResult]] | None = None,
) -> list[WebSearchResult]:
    """Search public web results for a correspondent with SerpApi only.

    Priority remains: phone number first, then name. The anti-trompeur filter now
    rejects links where the number appears only in the URL, but preserves social /
    professional results when the title/snippet proves a link to the correspondent.
    """
    api_key = clean_space(serpapi_api_key or os.getenv("SERPAPI_API_KEY", ""))
    if not api_key:
        return []
    n = normalize_phone(number) or re.sub(r"\D", "", number or "")
    name = clean_space(name)
    key = (n, name.lower(), int(max_results))
    if cache is not None and key in cache:
        return cache[key]

    queries: list[tuple[str, str]] = []
    if n:
        queries.append(("phone", f'"{n}" OR "237{n}" OR "+237{n}"'))
    if name and len(name) >= 4 and name.lower() not in {"inconnu", "non identifié", "non identifie"}:
        # The name query is important when the phone query returns mostly reverse-lookup pages.
        queries.append(("name", f'"{name}" Cameroun OR Cameroon OR "+237"'))

    preferred_domains = ("facebook.com", "fb.com", "linkedin.com", "scribd.com", "instagram.com", "twitter.com", "x.com")
    results: list[WebSearchResult] = []
    candidates: list[tuple[int, str]] = []
    seen_urls = set()

    for source, query in queries:
        try:
            query_string = urlencode(
                {
                    "engine": "google",
                    "q": query,
                    "api_key": api_key,
                    "num": max(20, int(max_results) * 8),
                    "hl": "fr",
                    "gl": "cm",
                }
            )
            request = Request(
                f"https://serpapi.com/search.json?{query_string}",
                headers={"User-Agent": "CDR-Insight-Studio/1.0"},
            )
            with urlopen(request, timeout=18) as response:
                payload = json.loads(response.read().decode("utf-8"))
        except Exception:
            continue

        organic = payload.get("organic_results") or []
        for item in organic:
            link = clean_space(item.get("link") or "")
            if not link.startswith(("http://", "https://")) or link in seen_urls:
                continue
            scored = _serpapi_candidate_score(item, link, n, name, preferred_domains)
            if scored is None:
                continue
            score, scored_link = scored
            # Phone-based evidence has priority, but name-based social/professional links
            # should not disappear just because generic phone lookup pages also ranked.
            source_penalty = 0 if source == "phone" else 1
            candidates.append((score + source_penalty, scored_link))
            seen_urls.add(link)

        # Only stop early when we already have enough high-value social/professional links.
        high_value = [link for score, link in candidates if score <= 1]
        if len(high_value) >= int(max_results):
            break

    seen_labels = set()
    for _, link in sorted(candidates, key=lambda x: x[0]):
        if len(results) >= int(max_results):
            break
        label = _serpapi_domain_label(link)
        display_label = label
        if display_label in seen_labels:
            display_label = f"{label} {sum(1 for r in results if r.label.startswith(label)) + 1}"
        results.append(WebSearchResult(display_label, link))
        seen_labels.add(display_label)

    if cache is not None:
        cache[key] = results
    return results


def exchange_counts(rows: list[CallRow], suspect_number: str) -> dict[str, int]:
    counts: dict[str, int] = {}
    for r in rows:
        other = None
        if r.caller == suspect_number and r.called and r.called != suspect_number:
            other = r.called
        elif r.called == suspect_number and r.caller and r.caller != suspect_number:
            other = r.caller
        if other:
            counts[other] = counts.get(other, 0) + 1
    return counts


def parse_certified_exchange_metrics(
    text: str,
) -> tuple[dict[str, int], dict[str, int], dict[str, int]]:
    """Read authoritative call, message and duration metrics from PDF recaps.

    In ``V. Fréquence par correspondant``, ``Total`` means every exchange with
    the correspondent, not calls only. ``VI. Fréquence par durée d'appel``
    supplies the message count and the certified call duration. Consequently:

        certified calls = total exchanges - certified messages

    A metric is returned only when the recap is internally consistent. Detailed
    rows remain available for geography and dates, while these certified values
    are authoritative for report counts and call duration.
    """
    normalized = (text or "").replace("\r", "")
    totals: dict[str, int] = {}
    messages: dict[str, int] = {}
    durations: dict[str, int] = {}

    frequency_match = re.search(
        r"(?im)^\s*V\.\s*Fr[ée]quence\s+par\s+correspondant\s*$",
        normalized,
    )
    duration_match = re.search(
        r"(?im)^\s*VI\.\s*Fr[ée]quence\s+par\s+dur[ée]e\s+d['’]appel\s*$",
        normalized,
    )
    if frequency_match:
        frequency_end = duration_match.start() if duration_match else len(normalized)
        lines = [
            clean_space(line)
            for line in normalized[frequency_match.end():frequency_end].splitlines()
            if clean_space(line)
        ]
        for index in range(len(lines) - 3):
            total_text, incoming_text, outgoing_text, phone_text = lines[index:index + 4]
            if not all(value.isdigit() for value in (total_text, incoming_text, outgoing_text)):
                continue
            phone = normalize_phone(phone_text)
            if not phone or not re.fullmatch(r"(?:237)?6\d{8}", phone_text):
                continue
            total = int(total_text)
            incoming = int(incoming_text)
            outgoing = int(outgoing_text)
            if total != incoming + outgoing:
                continue
            totals[phone] = max(totals.get(phone, 0), total)

    if duration_match:
        lines = [
            clean_space(line)
            for line in normalized[duration_match.end():].splitlines()
            if clean_space(line)
        ]
        for index, phone_text in enumerate(lines):
            phone = normalize_phone(phone_text)
            if not phone or not re.fullmatch(r"(?:237)?6\d{8}", phone_text):
                continue
            # A name can wrap onto several lines. The strict duration + integer
            # pair prevents page references and watermark fragments from matching.
            for cursor in range(index + 1, min(index + 9, len(lines) - 1)):
                if not CALL_DURATION_RE.fullmatch(lines[cursor]):
                    continue
                message_text = lines[cursor + 1]
                if message_text.isdigit():
                    messages[phone] = max(messages.get(phone, 0), int(message_text))
                    durations[phone] = max(
                        durations.get(phone, 0),
                        duration_to_seconds(lines[cursor]),
                    )
                break

    calls: dict[str, int] = {}
    for phone, total in totals.items():
        message_count = messages.get(phone)
        if message_count is None or message_count > total:
            continue
        calls[phone] = total - message_count

    return calls, messages, durations


def parse_certified_exchange_counts(
    text: str,
) -> tuple[dict[str, int], dict[str, int]]:
    """Compatibility wrapper returning corrected call and message counts."""
    calls, messages, _durations = parse_certified_exchange_metrics(text)
    return calls, messages


def report_exchange_counts(
    rows: Iterable[CallRow],
    suspect_number: str,
    correspondent_number: str,
    certified_call_counts: dict[str, int] | None = None,
    certified_sms_counts: dict[str, int] | None = None,
) -> tuple[int, int]:
    """Return certified totals when available, otherwise detailed-row totals."""
    calls, smsdata = count_exchanges_between(
        rows, suspect_number, correspondent_number
    )
    call_metrics = certified_call_counts or {}
    sms_metrics = certified_sms_counts or {}
    if correspondent_number in call_metrics:
        calls = call_metrics[correspondent_number]
    if correspondent_number in sms_metrics:
        smsdata = sms_metrics[correspondent_number]
    return calls, smsdata


def report_call_duration_between(
    rows: Iterable[CallRow],
    suspect_number: str,
    correspondent_number: str,
    certified_call_durations: dict[str, int] | None = None,
) -> int:
    """Return the certified call duration, with detailed rows as a fallback."""
    duration_metrics = certified_call_durations or {}
    if correspondent_number in duration_metrics:
        return duration_metrics[correspondent_number]
    return total_call_duration_between(rows, suspect_number, correspondent_number)


IMEI_SHARED_USER_RE = re.compile(
    r"(?m)^\s*((?:237)?6\d{8})\s+(\d{14,17})\s+(.+?)\s+(\d+)\s+"
    r"(\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}:\d{2})\s+"
    r"(\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}:\d{2})\s*$"
)


def _fold_text(value: str) -> str:
    return "".join(
        character
        for character in unicodedata.normalize("NFKD", value or "")
        if not unicodedata.combining(character)
    ).lower()


def parse_shared_imei_users(text: str) -> list[ImeiUserObservation]:
    """Extract the authoritative 'Imei partagé' chronology table."""
    source = (text or "").split("II. Listing", 1)[0]
    users: list[ImeiUserObservation] = []
    for match in IMEI_SHARED_USER_RE.finditer(source):
        number, imei, name, occurrences, first_seen, last_seen = match.groups()
        users.append(
            ImeiUserObservation(
                number=normalize_phone(number) or number,
                imei=imei,
                name=clean_space(name),
                occurrences=int(occurrences),
                first_seen=datetime.strptime(first_seen, "%d/%m/%Y %H:%M:%S"),
                last_seen=datetime.strptime(last_seen, "%d/%m/%Y %H:%M:%S"),
            )
        )
    return users


def detect_listing_type_from_text(text: str, filename: str = "") -> str:
    """Classify a telecom PDF without trusting the UI selection or filename alone."""
    folded = _fold_text(text)
    structural_folded = _fold_windev(text)
    shared_users = parse_shared_imei_users(text)
    # IMEI evidence always has priority over the generic telecom table headings.
    # This keeps mixed-batch protection effective even when a provider changes
    # the document filename.
    if shared_users and "imei partage" in folded:
        return "imei"
    name = Path(filename or "").stem
    if re.search(r"(?i)^requisition[_-].*[_-]\d{14,17}(?:[_-]|$)", name):
        return "imei"
    if re.search(r"(?i)^num[_-]", name):
        return "phone"
    if "operateur telephonique" in folded and re.search(
        r"(?i)\bimei\s*:\s*(?:n/?a|neant|aucun)", text or ""
    ):
        return "phone"
    # WinDev telephone requisitions can expose an empty ``IMEI :`` value and
    # use a neutral ``Requisition_<phone>.pdf`` filename. Hotfix 23 previously
    # stopped these documents at the type gate although the certified WinDev
    # parser already handled their rows correctly. The identity-section
    # signature distinguishes them from IMEI-sharing requisitions.
    if (
        "identite de l abonne" in structural_folded
        and "imei partage" not in structural_folded
        and is_windev_requisition(text)
    ):
        return "phone"
    return "unknown"


def _imei_operator(text: str, filename: str) -> str:
    candidate = f"{filename} {(text or '')[:12000]}".upper()
    if "ORANGE" in candidate:
        return "ORANGE"
    if "MTN" in candidate:
        return "MTN"
    return "N/D"


def _imei_period(text: str) -> str:
    match = re.search(
        r"P[ée]riode\s*:\s*du\s+(\d{2}/\d{2}/\d{4})\s+au\s+(\d{2}/\d{2}/\d{4})",
        text or "",
        re.I,
    )
    return f"{match.group(1)} au {match.group(2)}" if match else "N/D"


def _imei_requester(text: str) -> str:
    match = re.search(r"Service\s+Demandeur\s*:\s*([^\n\r]+)", text or "", re.I)
    return clean_space(match.group(1)) if match else "N/D"


def analyze_imei_listing(
    pdf_path: str | Path,
    top_n: int = 5,
    theft_date: str = "",
) -> ImeiListingAnalysis:
    pdf_path = Path(pdf_path)
    text = extract_text(pdf_path)
    if detect_listing_type_from_text(text, pdf_path.name) != "imei":
        raise ValueError("Le PDF ne présente pas la structure d'un listing basé sur un IMEI.")
    base = analyze_listing(pdf_path, 22, 7, max(1, top_n), False)
    users = parse_shared_imei_users(text)
    imei_counts: dict[str, int] = {}
    for row in base.rows:
        if row.imei:
            imei_counts[row.imei] = imei_counts.get(row.imei, 0) + 1
    target_imei = (
        max(users, key=lambda item: item.occurrences).imei
        if users
        else (max(imei_counts, key=imei_counts.get) if imei_counts else "")
    )
    if not target_imei:
        raise ValueError("Aucun IMEI exploitable n'a été détecté dans le PDF.")
    rows = [row for row in base.rows if not row.imei or row.imei == target_imei]
    if not users:
        grouped: dict[str, list[CallRow]] = {}
        for row in rows:
            if row.caller:
                grouped.setdefault(row.caller, []).append(row)
        for number, events in grouped.items():
            ident = base.identities.get(number)
            users.append(
                ImeiUserObservation(
                    number=number,
                    imei=target_imei,
                    name=ident.name if ident and ident.name else "Inconnu",
                    occurrences=len(events),
                    first_seen=min(row.dt for row in events),
                    last_seen=max(row.dt for row in events),
                )
            )
    users = [item for item in users if item.imei == target_imei]
    if not users:
        raise ValueError("Aucune ligne téléphonique associée à l'IMEI n'a été détectée.")
    dominant = max(users, key=lambda item: (item.occurrences, item.last_seen))
    identities = dict(base.identities)
    for item in users:
        current = identities.get(item.number)
        identities[item.number] = IdentificationRecord(
            number=item.number,
            name=item.name if item.name.lower() != "inconnu" else (current.name if current else ""),
            cni=current.cni if current else "",
            address=current.address if current else "",
        )
    return ImeiListingAnalysis(
        input_file=str(pdf_path),
        imei=target_imei,
        operator=_imei_operator(text, pdf_path.name),
        period=_imei_period(text),
        requester=_imei_requester(text),
        shared_users=sorted(users, key=lambda item: item.first_seen),
        dominant_user=dominant,
        rows=rows,
        identities=identities,
        parser_desc=f"imei-shared-v1; {base.parser_desc}",
        warnings=list(base.warnings),
        theft_date=clean_space(theft_date),
    )



def analyze_listing(
    pdf_path: str | Path,
    start_hour: int = 22,
    end_hour: int = 7,
    top_n: int = 5,
    add_extra_present: bool = True,
) -> ListingAnalysis:
    pdf_path = Path(pdf_path)
    text = extract_text(pdf_path)
    producer = pdf_producer(pdf_path)
    if is_pdfkit_mtn_requisition(text, producer):
        pdfkit_mtn = parse_pdfkit_mtn_requisition(pdf_path, text)
        coverage = (
            len(pdfkit_mtn.rows) / pdfkit_mtn.expected_events
            if pdfkit_mtn.expected_events else 1.0
        )
        if pdfkit_mtn.expected_events and not 0.90 <= coverage <= 1.10:
            # Controlled compatibility fallback: it is accepted only when the
            # established parser recovers at least 90% of the certified rows.
            generic_info = parse_subscriber_info(text, pdf_path)
            generic_rows, generic_desc = parse_listing_rows(text, pdf_path)
            generic_coverage = len(generic_rows) / pdfkit_mtn.expected_events
            if generic_coverage >= 0.90 and len(generic_rows) > len(pdfkit_mtn.rows):
                info = generic_info
                rows = generic_rows
                identities = parse_identification_map(pdf_path, text)
                parser_desc = (
                    f"{generic_desc}; fallback_from_pdfkit_mtn_coverage="
                    f"{coverage:.3f}"
                )
            else:
                raise ValueError(
                    "Le format MTN/PDFKit a été reconnu, mais seulement "
                    f"{len(pdfkit_mtn.rows)} événement(s) sur "
                    f"{pdfkit_mtn.expected_events} ont été extraits. "
                    "Rapport refusé pour éviter un résultat incomplet."
                )
        else:
            info = pdfkit_mtn.info
            rows = pdfkit_mtn.rows
            identities = pdfkit_mtn.identities
            parser_desc = pdfkit_mtn.parser_desc
    elif is_windev_requisition(text, producer):
        windev = parse_windev_requisition(pdf_path, text)
        listing_text = extract_text_until_tables_end(text)
        expected_events = len(re.findall(
            r"(?m)^\s*\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s*$",
            listing_text,
        ))
        windev_coverage = (
            len(windev.rows) / expected_events if expected_events else 1.0
        )
        if expected_events and windev_coverage < 0.90:
            # Integrity guard for unknown/malformed producers: never emit a
            # severely incomplete report when the established parser recovers
            # more of the source document.
            generic_info = parse_subscriber_info(text, pdf_path)
            generic_rows, generic_desc = parse_listing_rows(text, pdf_path)
            if len(generic_rows) > len(windev.rows):
                info = generic_info
                rows = generic_rows
                identities = parse_identification_map(pdf_path, text)
                parser_desc = (
                    f"{generic_desc}; fallback_from_windev_coverage="
                    f"{windev_coverage:.3f}"
                )
            else:
                info = windev.info
                rows = windev.rows
                identities = windev.identities
                parser_desc = windev.parser_desc
        else:
            info = windev.info
            rows = windev.rows
            identities = windev.identities
            parser_desc = windev.parser_desc
    else:
        # Existing formats intentionally retain their established parser chain.
        info = parse_subscriber_info(text, pdf_path)
        rows, parser_desc = parse_listing_rows(text, pdf_path)
        identities = parse_identification_map(pdf_path, text)
    total_dates = len(re.findall(r"\d{2}/\d{2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}", text))
    warnings: list[str] = []
    if not rows and total_dates:
        raise ValueError(
            f"Aucune ligne n'a été parsée alors que {total_dates} date(s) d'événements existent dans le PDF. "
            "Ce format nécessite probablement un OCR/table parser supplémentaire."
        )
    if not rows:
        warnings.append("Aucune ligne d'appel/SMS détectée dans le listing.")
    if not info.number:
        all_numbers = [n for r in rows for n in [r.caller, r.called] if n]
        info.number = max(set(all_numbers), key=all_numbers.count) if all_numbers else normalize_phone(pdf_path.stem) or ""
    suspect = info.number
    if suspect in identities:
        suspect_ident = identities[suspect]
        if (not info.name or info.name.lower() in {"inconnu", "non identifié", "non identifie"}) and suspect_ident.name:
            info.name = suspect_ident.name
        if not info.cni and suspect_ident.cni:
            info.cni = suspect_ident.cni
        if not info.address and suspect_ident.address:
            info.address = suspect_ident.address
    suspect_events = [r for r in rows if r.caller == suspect]
    suspect_choice = choose_district(suspect_events, start_hour, end_hour)

    counts = exchange_counts(rows, suspect)
    (
        certified_call_counts,
        certified_sms_counts,
        certified_call_durations,
    ) = parse_certified_exchange_metrics(text)
    top = [n for n, _ in sorted(counts.items(), key=lambda kv: kv[1], reverse=True)[:top_n]]
    selected = list(top)
    if add_extra_present and suspect_choice.district != "N/D":
        candidates = []
        for n, c in counts.items():
            if n in selected:
                continue
            events = [r for r in rows if r.caller == n]
            dates = dates_in_suspect_district(events, suspect_choice.district, suspect)
            if dates:
                candidates.append((n, len(set(dates)), c))
        if candidates:
            extra = sorted(candidates, key=lambda x: (x[1], x[2]), reverse=True)[0][0]
            selected.append(extra)

    return ListingAnalysis(
        input_file=str(pdf_path),
        info=info,
        rows=rows,
        parser_desc=parser_desc,
        warnings=warnings,
        suspect=suspect,
        suspect_choice=suspect_choice,
        counts=counts,
        top=top,
        selected=selected,
        identities=identities,
        certified_call_counts=certified_call_counts,
        certified_sms_counts=certified_sms_counts,
        certified_call_durations=certified_call_durations,
    )


def setup_report_document(logo_path: str | Path | None = None) -> Document:
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.left_margin = Cm(0.7)
    section.right_margin = Cm(0.7)
    # Leave enough room for the repeating logo header on every Word page.
    section.top_margin = Cm(2.0)
    section.header_distance = Cm(0.25)
    section.bottom_margin = Cm(0.7)
    active_logo_path = Path(logo_path) if logo_path else DEFAULT_LOGO_PATH
    add_repeating_logo_header(doc, active_logo_path)
    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(7)
    return doc


def add_listing_report_section(
    doc: Document,
    analysis: ListingAnalysis,
    start_hour: int = 22,
    end_hour: int = 7,
    enable_internet_search: bool = False,
    serpapi_api_key: str | None = None,
    max_web_results: int = 3,
    web_search_cache: dict[tuple[str, str, int], list[WebSearchResult]] | None = None,
) -> None:
    info = analysis.info
    rows = analysis.rows
    suspect = analysis.suspect
    suspect_choice = analysis.suspect_choice
    selected = analysis.selected
    top = analysis.top
    identities = analysis.identities
    warnings = analysis.warnings
    suspect_events = [r for r in rows if r.caller == suspect]

    add_title(doc, f"RAPPORT D'ANALYSE DU LISTING {display_phone(suspect)}")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"Plage horaire prioritaire : {start_hour:02d}h-{end_hour:02d}h")
    r.bold = True
    r.font.size = Pt(9)

    add_section(doc, "INFORMATIONS PRINCIPALES DE LA LIGNE ANALYSÉE")
    main_table = doc.add_table(rows=2, cols=10)
    main_table.style = "Table Grid"
    headers = ["NUMÉRO", "NOM LISTING", "OPÉRATEUR", "N° CNI", "ADRESSE", "VILLE", "QUARTIER RETENU", "LOCALISATION", "SOURCE QUARTIER", "PÉRIODE"]
    values = [display_phone(suspect), info.name, info.operator, info.cni, info.address, suspect_choice.city, suspect_choice.display, suspect_choice.localisation, suspect_choice.source, info.period]
    fill_header_row(main_table.rows[0], headers)
    fill_value_row(main_table.rows[1], values)

    add_section(doc, "TOP 3 DES QUARTIERS DE LA LIGNE ANALYSÉE")
    topq = top_districts(suspect_events, 3)
    tq = doc.add_table(rows=1, cols=5)
    tq.style = "Table Grid"
    fill_header_row(tq.rows[0], ["RANG", "QUARTIER", "VILLE", "OCCURRENCES", "JOURS DISTINCTS"])
    for dist, city, count, days, rank in topq:
        row = tq.add_row()
        fill_value_row(row, [str(rank), dist, city, str(count), str(days)])
    if not topq:
        row = tq.add_row()
        fill_value_row(row, ["-", "N/D", "N/D", "0", "0"])

    add_section(doc, "MEILLEURS CORRESPONDANTS ET PRIORITÉ D'ANALYSE")
    headers = [
        "N°", "NUMÉRO TÉLÉPHONE", "NOM LISTING", "NOM CALL APP", "NOM ORANGE MONEY",
        "QUARTIER RETENU", "OCCURENCE DANS LE QUARTIER RETENUE/JOUR DISTINCT", "LOCALISATION",
        "A ÉTÉ DANS LE QUARTIER DE LA LIGNE ANALYSÉE ?",
        "TOTAL DES ÉCHANGES AVEC LA LIGNE ANALYSÉE — APPELS / SMS-DATA",
        "DURÉE TOTAL APPEL", "RECHERCHE SUR INTERNET", "PRIORITÉ D'ANALYSE",
    ]
    search_col_idx = headers.index("RECHERCHE SUR INTERNET")
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    fill_header_row(table.rows[0], headers)
    if selected:
        active_serpapi_key = clean_space(serpapi_api_key or os.getenv("SERPAPI_API_KEY", ""))
        for idx, n in enumerate(selected, start=1):
            events = [r for r in rows if r.caller == n]
            choice = choose_district(events, start_hour, end_hour)
            total_calls, total_sms = report_exchange_counts(
                rows,
                suspect,
                n,
                analysis.certified_call_counts,
                analysis.certified_sms_counts,
            )
            dates = dates_in_suspect_district(events, suspect_choice.district, suspect)
            present = bool(dates)
            in_top = n in top
            if in_top and present:
                priority = "FORTE"
            elif present and len(dates) >= 2:
                priority = "FORTE"
            elif present or in_top:
                priority = "MOYENNE"
            else:
                priority = "FAIBLE"
            ident = identities.get(n)
            nom_listing = ident.name if ident and ident.name else ""
            occurrence_jours = f"{choice.total_count}/{choice.total_days}" if choice.district != "N/D" else "0/0"
            total_duration = format_duration(report_call_duration_between(
                rows,
                suspect,
                n,
                analysis.certified_call_durations,
            ))
            if enable_internet_search and active_serpapi_key:
                web_links = search_internet_for_correspondent(
                    n,
                    nom_listing,
                    serpapi_api_key=active_serpapi_key,
                    max_results=max_web_results,
                    cache=web_search_cache,
                )
                web_status = "Aucun résultat"
            elif enable_internet_search:
                web_links = []
                web_status = "Clé SerpApi manquante"
            else:
                web_links = []
                web_status = "Non activée"
            row = table.add_row()
            fill_value_row(row, [
                str(idx), display_phone(n), nom_listing, "", "", choice.display, occurrence_jours,
                choice.localisation, format_dates(dates), f"{total_calls}, {total_sms}",
                total_duration, web_status, priority
            ])
            if web_links:
                set_cell_hyperlinks(row.cells[search_col_idx], web_links)
    else:
        row = table.add_row()
        fill_value_row(row, [
            "-", "N/D", "", "", "", "N/D", "0/0", "N/D", "NON", "0, 0",
            "00:00:00", "Non activée", "N/D",
        ])
    if warnings:
        note2 = doc.add_paragraph()
        note2.add_run("Avertissement technique : ").bold = True
        note2.add_run(" ".join(warnings))


def summary_from_analysis(analysis: ListingAnalysis, output_path: str | Path) -> ReportSummary:
    return ReportSummary(
        analysis.input_file,
        str(output_path),
        analysis.suspect,
        analysis.info.name,
        len(analysis.rows),
        len(analysis.counts),
        analysis.suspect_choice.display,
        analysis.parser_desc,
        "; ".join(analysis.warnings),
    )


def make_report(
    pdf_path: str | Path,
    out_dir: str | Path,
    start_hour: int = 22,
    end_hour: int = 7,
    top_n: int = 5,
    add_extra_present: bool = True,
    logo_path: str | Path | None = None,
    enable_internet_search: bool = False,
    serpapi_api_key: str | None = None,
    max_web_results: int = 3,
) -> ReportSummary:
    pdf_path = Path(pdf_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    analysis = analyze_listing(pdf_path, start_hour, end_hour, top_n, add_extra_present)
    doc = setup_report_document(logo_path)
    add_listing_report_section(
        doc,
        analysis,
        start_hour,
        end_hour,
        enable_internet_search=enable_internet_search,
        serpapi_api_key=serpapi_api_key,
        max_web_results=max_web_results,
        web_search_cache={},
    )
    safe_number = analysis.suspect or re.sub(r"\W+", "_", pdf_path.stem)
    output_path = out_dir / f"rapport_listing_{safe_number}_{start_hour:02d}h_{end_hour:02d}h.docx"
    doc.save(output_path)
    return summary_from_analysis(analysis, output_path)


def make_combined_report(
    pdf_paths: Iterable[str | Path],
    out_dir: str | Path,
    start_hour: int = 22,
    end_hour: int = 7,
    top_n: int = 5,
    add_extra_present: bool = True,
    logo_path: str | Path | None = None,
    enable_internet_search: bool = False,
    serpapi_api_key: str | None = None,
    max_web_results: int = 3,
) -> tuple[Path, list[ReportSummary]]:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    pdf_paths = [Path(p) for p in pdf_paths]
    if not pdf_paths:
        raise ValueError("Aucun listing PDF fourni.")

    analyses = [analyze_listing(p, start_hour, end_hour, top_n, add_extra_present) for p in pdf_paths]
    doc = setup_report_document(logo_path)
    web_search_cache: dict[tuple[str, str, int], list[WebSearchResult]] = {}
    for idx, analysis in enumerate(analyses):
        if idx > 0:
            doc.add_page_break()
        add_listing_report_section(
            doc,
            analysis,
            start_hour,
            end_hour,
            enable_internet_search=enable_internet_search,
            serpapi_api_key=serpapi_api_key,
            max_web_results=max_web_results,
            web_search_cache=web_search_cache,
        )

    first_number = analyses[0].suspect or "listing"
    output_path = out_dir / f"rapport_global_{first_number}_{len(analyses)}_listings_{start_hour:02d}h_{end_hour:02d}h.docx"
    doc.save(output_path)
    summaries = [summary_from_analysis(a, output_path) for a in analyses]
    return output_path, summaries


def _selected_imei_users(
    users: list[ImeiUserObservation], limit: int = 10
) -> list[ImeiUserObservation]:
    if len(users) <= limit:
        return sorted(users, key=lambda item: item.first_seen)
    candidates = [
        max(users, key=lambda item: (item.occurrences, item.last_seen)),
        min(users, key=lambda item: item.first_seen),
        max(users, key=lambda item: item.last_seen),
    ]
    candidates.extend(sorted(users, key=lambda item: item.occurrences, reverse=True))
    selected: list[ImeiUserObservation] = []
    seen: set[str] = set()
    for item in candidates:
        if item.number in seen:
            continue
        selected.append(item)
        seen.add(item.number)
        if len(selected) >= limit:
            break
    return sorted(selected, key=lambda item: item.first_seen)


def _row_location(row: CallRow | None) -> str:
    if not row:
        return "N/D"
    parts = [value for value in [row.district, row.city] if value and value != "N/D"]
    label = " — ".join(dict.fromkeys(parts)) or clean_space(row.location_clean) or "N/D"
    if row.lat and row.lon:
        label += f" (Lat {row.lat}, Long {row.lon})"
    return label


def add_imei_report_section(
    doc: Document,
    analysis: ImeiListingAnalysis,
    top_n: int = 5,
) -> None:
    dominant = analysis.dominant_user
    dominant_ident = analysis.identities.get(dominant.number)
    dominant_events = [row for row in analysis.rows if row.caller == dominant.number]
    all_events = sorted(analysis.rows, key=lambda row: row.dt)
    principal_events = sorted(dominant_events, key=lambda row: row.dt)

    add_title(doc, f"RAPPORT D'ANALYSE IMEI {analysis.imei}")
    metadata = doc.add_table(rows=2, cols=6)
    metadata.style = "Table Grid"
    fill_header_row(
        metadata.rows[0],
        ["IMEI", "OPÉRATEUR", "PÉRIODE", "SERVICE DEMANDEUR", "DATE DU VOL", "SOURCE"],
    )
    fill_value_row(
        metadata.rows[1],
        [
            analysis.imei,
            analysis.operator,
            analysis.period,
            analysis.requester,
            analysis.theft_date or "Non renseignée",
            Path(analysis.input_file).name,
        ],
    )

    add_section(doc, "LIGNE PRINCIPALE OBSERVÉE")
    principal = doc.add_table(rows=2, cols=7)
    principal.style = "Table Grid"
    fill_header_row(
        principal.rows[0],
        ["LIGNE", "IDENTITÉ", "N° CNI", "ADRESSE", "OCCURRENCES", "PREMIÈRE UTILISATION", "DERNIÈRE UTILISATION"],
    )
    fill_value_row(
        principal.rows[1],
        [
            display_phone(dominant.number),
            dominant.name,
            dominant_ident.cni if dominant_ident else "",
            dominant_ident.address if dominant_ident else "",
            str(dominant.occurrences),
            dominant.first_seen.strftime("%d/%m/%Y %H:%M:%S"),
            dominant.last_seen.strftime("%d/%m/%Y %H:%M:%S"),
        ],
    )

    total_calls = sum(1 for row in analysis.rows if row.is_call)
    total_smsdata = len(analysis.rows) - total_calls
    total_duration = sum(duration_to_seconds(row.duration) for row in analysis.rows if row.is_call)
    add_section(doc, "VUE GLOBALE")
    metrics = doc.add_table(rows=2, cols=6)
    metrics.style = "Table Grid"
    fill_header_row(
        metrics.rows[0],
        ["LIGNES DANS L'IMEI", "ÉVÉNEMENTS", "APPELS", "DURÉE DES APPELS", "SMS / DATA", "LIGNE DOMINANTE"],
    )
    fill_value_row(
        metrics.rows[1],
        [
            str(len(analysis.shared_users)),
            str(len(analysis.rows)),
            str(total_calls),
            format_duration(total_duration),
            str(total_smsdata),
            display_phone(dominant.number),
        ],
    )

    add_section(doc, "CHRONOLOGIE DES LIGNES OBSERVÉES DANS L'IMEI")
    chronology = doc.add_table(rows=1, cols=6)
    chronology.style = "Table Grid"
    fill_header_row(
        chronology.rows[0],
        ["QUALIFICATION", "LIGNE", "IDENTITÉ", "OCCURRENCES", "PREMIÈRE UTILISATION", "DERNIÈRE UTILISATION"],
    )
    for item in _selected_imei_users(analysis.shared_users):
        row = chronology.add_row()
        fill_value_row(
            row,
            [
                "Ligne principale" if item.number == dominant.number else "Ligne observée",
                display_phone(item.number),
                item.name,
                str(item.occurrences),
                item.first_seen.strftime("%d/%m/%Y %H:%M:%S"),
                item.last_seen.strftime("%d/%m/%Y %H:%M:%S"),
            ],
        )

    add_section(doc, "REPÈRES DE LOCALISATION")
    markers = doc.add_table(rows=1, cols=4)
    markers.style = "Table Grid"
    fill_header_row(markers.rows[0], ["REPÈRE", "DATE", "LIGNE", "LOCALISATION"])
    marker_rows = [
        ("Première trace IMEI", all_events[0] if all_events else None),
        ("Dernière trace de la ligne dominante", principal_events[-1] if principal_events else None),
        ("Dernière trace IMEI", all_events[-1] if all_events else None),
    ]
    for label, event in marker_rows:
        row = markers.add_row()
        fill_value_row(
            row,
            [
                label,
                event.dt.strftime("%d/%m/%Y %H:%M:%S") if event else "N/D",
                display_phone(event.caller) if event and event.caller else "N/D",
                _row_location(event),
            ],
        )

    counts = exchange_counts(analysis.rows, dominant.number)
    correspondents = sorted(counts, key=counts.get, reverse=True)[: max(1, top_n)]
    add_section(doc, "CORRESPONDANTS PRINCIPAUX DE LA LIGNE PRINCIPALE")
    contacts = doc.add_table(rows=1, cols=7)
    contacts.style = "Table Grid"
    fill_header_row(
        contacts.rows[0],
        ["RANG", "LIGNE", "IDENTITÉ", "APPELS", "SMS / DATA", "DURÉE APPELS", "TOTAL ÉCHANGES"],
    )
    for rank, number in enumerate(correspondents, start=1):
        calls, smsdata = count_exchanges_between(analysis.rows, dominant.number, number)
        ident = analysis.identities.get(number)
        row = contacts.add_row()
        fill_value_row(
            row,
            [
                str(rank),
                display_phone(number),
                ident.name if ident else "",
                str(calls),
                str(smsdata),
                format_duration(total_call_duration_between(analysis.rows, dominant.number, number)),
                str(calls + smsdata),
            ],
        )
    if not correspondents:
        fill_value_row(contacts.add_row(), ["-", "N/D", "", "0", "0", "00:00:00", "0"])

    add_section(doc, "ZONES PRINCIPALES DE LA LIGNE PRINCIPALE")
    zones = doc.add_table(rows=1, cols=5)
    zones.style = "Table Grid"
    fill_header_row(zones.rows[0], ["RANG", "QUARTIER / SITE", "VILLE", "OCCURRENCES", "JOURS DISTINCTS"])
    ranked_zones = top_districts(dominant_events, 5)
    for district, city, count, days, rank in ranked_zones:
        fill_value_row(zones.add_row(), [str(rank), district, city, str(count), str(days)])
    if not ranked_zones:
        fill_value_row(zones.add_row(), ["-", "N/D", "N/D", "0", "0"])


def imei_summary_from_analysis(
    analysis: ImeiListingAnalysis, output_path: str | Path
) -> ReportSummary:
    dominant_events = [row for row in analysis.rows if row.caller == analysis.dominant_user.number]
    ranked_zones = top_districts(dominant_events, 1)
    retained_district = ranked_zones[0][0] if ranked_zones else "N/D"
    return ReportSummary(
        analysis.input_file,
        str(output_path),
        analysis.dominant_user.number,
        analysis.dominant_user.name,
        len(analysis.rows),
        len(exchange_counts(analysis.rows, analysis.dominant_user.number)),
        retained_district,
        analysis.parser_desc,
        "; ".join(analysis.warnings),
    )


def make_imei_report(
    pdf_path: str | Path,
    out_dir: str | Path,
    top_n: int = 5,
    logo_path: str | Path | None = None,
    theft_date: str = "",
) -> ReportSummary:
    pdf_path = Path(pdf_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    analysis = analyze_imei_listing(pdf_path, top_n=top_n, theft_date=theft_date)
    doc = setup_report_document(logo_path or IMEI_LOGO_PATH)
    add_imei_report_section(doc, analysis, top_n=top_n)
    output_path = out_dir / f"rapport_imei_{analysis.imei}.docx"
    doc.save(output_path)
    return imei_summary_from_analysis(analysis, output_path)


def make_combined_imei_report(
    pdf_paths: Iterable[str | Path],
    out_dir: str | Path,
    top_n: int = 5,
    logo_path: str | Path | None = None,
    theft_date: str = "",
) -> tuple[Path, list[ReportSummary]]:
    paths = [Path(path) for path in pdf_paths]
    if not paths:
        raise ValueError("Aucun listing IMEI fourni.")
    analyses = [
        analyze_imei_listing(path, top_n=top_n, theft_date=theft_date)
        for path in paths
    ]
    imeis = {analysis.imei for analysis in analyses}
    if len(imeis) != 1:
        raise ValueError(
            "Un rapport IMEI combiné ne peut contenir qu'un seul IMEI. "
            "Choisissez un document Word par listing."
        )
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    doc = setup_report_document(logo_path or IMEI_LOGO_PATH)
    for index, analysis in enumerate(analyses):
        if index:
            doc.add_page_break()
        add_imei_report_section(doc, analysis, top_n=top_n)
    imei = next(iter(imeis))
    output_path = out_dir / f"rapport_imei_{imei}_{len(analyses)}_listings.docx"
    doc.save(output_path)
    return output_path, [imei_summary_from_analysis(item, output_path) for item in analyses]


def clear_paragraph(paragraph) -> None:
    for run in list(paragraph.runs):
        paragraph._p.remove(run._r)


def add_repeating_logo_header(doc: Document, logo_path: str | Path | None) -> None:
    if not logo_path:
        return
    logo = Path(logo_path)
    if not logo.exists():
        return
    for section in doc.sections:
        section.different_first_page_header_footer = False
        header = section.header
        paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        clear_paragraph(paragraph)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run()
        run.add_picture(str(logo), width=Cm(2.2))


def add_title(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(14)


def add_section(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(10)


def shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_text(cell, value: str, bold: bool = False, size: int = 6) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(str(value or ""))
    r.bold = bold
    r.font.size = Pt(size)
    r.font.name = "Arial"


def fill_header_row(row, values: list[str]) -> None:
    for cell, value in zip(row.cells, values, strict=True):
        set_cell_text(cell, value, bold=True, size=6)
        shade_cell(cell, "D9EAF7")


def fill_value_row(row, values: list[str]) -> None:
    for cell, value in zip(row.cells, values, strict=True):
        set_cell_text(cell, value, bold=False, size=6)


def add_hyperlink(paragraph, text: str, url: str) -> None:
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    r_pr.append(color)
    r_pr.append(underline)
    new_run.append(r_pr)

    text_element = OxmlElement("w:t")
    text_element.text = text
    new_run.append(text_element)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def set_cell_hyperlinks(cell, links: list[WebSearchResult]) -> None:
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for idx, link in enumerate(links):
        add_hyperlink(p, link.label, link.url)
        if idx < len(links) - 1:
            r = p.add_run(", ")
            r.font.size = Pt(6)
            r.font.name = "Arial"


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Generate Word reports from telecom listing PDFs.")
    parser.add_argument("pdf", nargs="+", help="PDF listing file(s)")
    parser.add_argument("--out", default="reports", help="Output directory")
    parser.add_argument("--start-hour", type=int, default=22)
    parser.add_argument("--end-hour", type=int, default=7)
    parser.add_argument("--enable-internet-search", action="store_true")
    parser.add_argument("--serpapi-api-key", default=None)
    parser.add_argument("--max-web-results", type=int, default=3)
    args = parser.parse_args()
    for pdf in args.pdf:
        summary = make_report(
            pdf,
            args.out,
            args.start_hour,
            args.end_hour,
            enable_internet_search=args.enable_internet_search,
            serpapi_api_key=args.serpapi_api_key,
            max_web_results=args.max_web_results,
        )
        print(summary.output_file)

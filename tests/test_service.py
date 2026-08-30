from io import BytesIO
from zipfile import ZipFile

import pytest

from cdr_insight.demo import build_demo_pdf
from cdr_insight.service import (
    ProcessingError,
    ProcessingOptions,
    Upload,
    process_uploads,
    safe_filename,
)


def test_safe_filename_removes_paths_and_special_characters() -> None:
    assert safe_filename("../../Mon listing (démo).pdf") == "Mon_listing_d_mo_.pdf"


def test_demo_pdf_generates_a_word_report_end_to_end() -> None:
    result = process_uploads(
        [Upload("listing_demo_synthetique.pdf", build_demo_pdf())],
        ProcessingOptions(listing_type="auto"),
    )
    assert result.listing_type == "phone"
    assert result.filename.endswith(".docx")
    assert result.content.startswith(b"PK")
    assert result.summaries[0]["records"] == 3
    with ZipFile(BytesIO(result.content)) as document:
        assert "word/document.xml" in document.namelist()


def test_non_pdf_is_rejected_before_parsing() -> None:
    with pytest.raises(ProcessingError, match="PDF valide"):
        process_uploads([Upload("fake.pdf", b"not a pdf")])


def test_web_enrichment_requires_an_explicit_key() -> None:
    with pytest.raises(ProcessingError, match="clé SerpAPI"):
        process_uploads(
            [Upload("listing_demo_synthetique.pdf", build_demo_pdf())],
            ProcessingOptions(enable_internet_search=True),
        )

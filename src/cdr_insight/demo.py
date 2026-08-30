"""Generate a fully synthetic PDF for local demonstrations and tests."""

from __future__ import annotations

import io

import pymupdf as fitz

DEMO_TEXT = """Service Demandeur : UNITÉ DE DÉMONSTRATION
I. Identité de l'abonné
Téléphone : 600000001
Nom et Prénom : PERSONNE ALPHA
Opérateur téléphonique : DEMO TELCO
Numéro CNI : N/A
IMEI : N/A
Adresse : ADRESSE DE DÉMONSTRATION
II. Listing
600000001
BONAPRISO-DOUALA (Cell: DEMO-01 Long: 9.704 Lat: 4.043)
490154203237518
12/08/2026 22:30:00
00:01:05
600000002
600000001
BONAPRISO-DOUALA (Cell: DEMO-01 Long: 9.704 Lat: 4.043)
490154203237518
13/08/2026 23:15:00
SMS/DATA
600000002
600000002
AKWA-DOUALA (Cell: DEMO-02 Long: 9.707 Lat: 4.051)
490154203237518
14/08/2026 08:10:00
00:00:42
600000001
"""


def build_demo_pdf() -> bytes:
    """Return a small in-memory PDF containing fictional telecom records."""
    document = fitz.open()
    page = document.new_page(width=595, height=842)
    page.insert_textbox(
        fitz.Rect(45, 45, 550, 800),
        DEMO_TEXT,
        fontsize=9,
        fontname="helv",
        lineheight=1.25,
    )
    buffer = io.BytesIO()
    document.save(buffer)
    document.close()
    return buffer.getvalue()

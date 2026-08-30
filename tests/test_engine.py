from datetime import datetime

from cdr_insight.engine import (
    CallRow,
    choose_district,
    count_exchanges_between,
    detect_listing_type_from_text,
    duration_to_seconds,
    normalize_phone,
    total_call_duration_between,
)


def _row(
    caller: str,
    called: str,
    hour: int,
    district: str = "ZONE DEMO",
    duration: str = "00:01:00",
) -> CallRow:
    return CallRow(
        caller_raw=caller,
        caller=caller,
        called_raw=called,
        called=called,
        location_raw=f"{district}-VILLE DEMO (Cell: DEMO)",
        location_clean=f"{district}-VILLE DEMO (Cell: DEMO)",
        district=district,
        city="VILLE DEMO",
        lat="4.000",
        lon="9.000",
        cell="DEMO",
        imei="490154203237518",
        dt=datetime(2026, 8, 12, hour, 30),
        duration=duration,
    )


def test_phone_normalization_and_listing_detection() -> None:
    assert normalize_phone("+237 600 000 001") == "600000001"
    text = "Opérateur téléphonique : DEMO\nIMEI : N/A\nII. Listing"
    assert detect_listing_type_from_text(text, "document.pdf") == "phone"


def test_exchange_metrics_are_bidirectional() -> None:
    subject = "600000001"
    contact = "600000002"
    rows = [
        _row(subject, contact, 22, duration="00:01:05"),
        _row(contact, subject, 23, duration="SMS/DATA"),
        _row(contact, subject, 1, duration="00:00:35"),
    ]
    assert count_exchanges_between(rows, subject, contact) == (2, 1)
    assert total_call_duration_between(rows, subject, contact) == 100
    assert duration_to_seconds("01:02:03") == 3723


def test_district_choice_prefers_repeated_night_activity() -> None:
    rows = [
        _row("600000001", "600000002", 22, "ZONE NUIT"),
        _row("600000001", "600000003", 23, "ZONE NUIT"),
        _row("600000001", "600000004", 10, "ZONE JOUR"),
    ]
    choice = choose_district(rows, 22, 7)
    assert choice.district == "ZONE NUIT"
    assert choice.window_count == 2

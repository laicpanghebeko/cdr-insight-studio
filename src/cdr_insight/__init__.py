"""CDR Insight Studio — standalone telecom listing analysis."""

from .engine import (
    CallRow,
    DistrictChoice,
    ImeiListingAnalysis,
    ListingAnalysis,
    ReportSummary,
    analyze_imei_listing,
    analyze_listing,
    detect_listing_type_from_text,
)
from .service import GeneratedReport, ProcessingOptions, process_uploads

__all__ = [
    "CallRow",
    "DistrictChoice",
    "GeneratedReport",
    "ImeiListingAnalysis",
    "ListingAnalysis",
    "ProcessingOptions",
    "ReportSummary",
    "analyze_imei_listing",
    "analyze_listing",
    "detect_listing_type_from_text",
    "process_uploads",
]

__version__ = "1.0.0"

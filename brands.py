"""
brands.py
Central brand definitions for the quote tool.

Used by:
  - app.py            → injects a public copy into the page so the UI can build
                        the brand picker, and serves logos via /brand-logo/<key>
  - pdf_generator.py  → header banner colours / title / logo / footer
  - excel_generator.py→ title row colours / text / logo

Colours are stored as 6-digit hex WITHOUT a leading '#'. Helpers add it where a
given library needs it (ReportLab wants '#RRGGBB', openpyxl wants 'RRGGBB').

Logo binaries live in Blob storage under brand/… (see upload_brand.py) and are
fetched at render time via storage.get_bytes(). This keeps the app filesystem
ephemeral-safe on Azure App Service — same model as the JSON data in storage.py.
"""
from typing import Any, Dict

DEFAULT_BRAND = "marlu"

BRANDS: Dict[str, Dict[str, Any]] = {
    "marlu": {
        "key":         "marlu",
        "ui_name":     "Marlu Group",            # sidebar label
        "doc_name":    "MARLU GROUP",             # PDF banner / Excel title
        "tagline":     "Labour Hire — Quotation",
        "excel_title": "MARLU GROUP — Labour Hire Price Calculator & Quote",
        "footer_note": "Marlu Group standard terms & conditions.",
        "footer_strip":"Marlu Group  -  Labour Hire Quotation",
        "banner_bg":   "FFFFFF",   # white banner
        "banner_fg":   "1A1A1A",   # dark text
        "accent":      "2E8B8B",   # UI accent
        "xl_title_fg": "1A6060",   # Excel title text (dark teal)
        "xl_title_bg": "FFFFFF",   # Excel title fill (white)
        "logo_blob":   "brand/marlu-logo.jpg",
        # white banner → logo sits directly, no chip needed
        "logo_on_chip": False,
    },
    "westlink": {
        "key":         "westlink",
        "ui_name":     "WestLink Workforce",
        "doc_name":    "WESTLINK WORKFORCE",
        "tagline":     "Labour Hire — Quotation",
        "excel_title": "WESTLINK WORKFORCE — Labour Hire Price Calculator & Quote",
        "footer_note": "WestLink Workforce standard terms & conditions.",
        "footer_strip":"WestLink Workforce  -  Labour Hire Quotation",
        "banner_bg":   "FFFFFF",   # white banner
        "banner_fg":   "1A1A1A",   # dark text
        "accent":      "6FBF1B",   # green
        "xl_title_fg": "2E6B0F",   # Excel title text (dark green)
        "xl_title_bg": "FFFFFF",   # Excel title fill (white)
        "logo_blob":   "brand/westlink-logo.png",
        # transparent logo on a white banner → no chip needed
        "logo_on_chip": False,
    },
}


def get_brand(key: Any) -> Dict[str, Any]:
    """Return the brand config for `key`, falling back to the default brand."""
    k = str(key or "").strip().lower()
    return BRANDS.get(k, BRANDS[DEFAULT_BRAND])


def brand_key(payload: Dict[str, Any]) -> str:
    """Pull the selected brand key out of an export payload (default-safe)."""
    k = str((payload or {}).get("brand") or "").strip().lower()
    return k if k in BRANDS else DEFAULT_BRAND


def public_config() -> Dict[str, Any]:
    """The subset of brand info safe to embed in the page for the UI picker."""
    return {
        "default": DEFAULT_BRAND,
        "brands": [
            {
                "key":     b["key"],
                "ui_name": b["ui_name"],
                "accent":  "#" + b["accent"],
                "logo_url": f"/brand-logo/{b['key']}",
            }
            for b in BRANDS.values()
        ],
    }

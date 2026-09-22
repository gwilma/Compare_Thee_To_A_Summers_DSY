"""Read CIBSE weather-file metadata (location, file type, period, scenario) from file names.

CIBSE names its files along the lines of ``London_LHR_DSY1_2050High50.epw`` or
``Manchester_TRY.csv``; naming varies between releases, so this is lenient and
every field can be corrected in the app.
"""

from __future__ import annotations

import re

#: CIBSE 2016 weather locations and the tokens that identify them in file names.
CIBSE_LOCATIONS = {
    "Belfast": ["belfast"],
    "Birmingham": ["birmingham"],
    "Cardiff": ["cardiff"],
    "Edinburgh": ["edinburgh"],
    "Glasgow": ["glasgow"],
    "Leeds": ["leeds"],
    "London (Heathrow)": ["lhr", "heathrow"],
    "London (Weather Centre)": ["lwc", "weathercentre", "weather_centre"],
    "London (Gatwick)": ["lgw", "gatwick"],
    "Manchester": ["manchester"],
    "Newcastle": ["newcastle"],
    "Norwich": ["norwich"],
    "Nottingham": ["nottingham"],
    "Plymouth": ["plymouth"],
    "Southampton": ["southampton"],
    "Swindon": ["swindon"],
}

KINDS = ["TRY", "DSY1", "DSY2", "DSY3"]
PERIODS = ["Baseline", "2020s", "2050s", "2080s"]
EMISSIONS = ["Low", "Medium", "High"]
PERCENTILES = [10, 50, 90]


def parse_reference_name(filename: str) -> dict:
    """Best-effort metadata from a CIBSE-style file name. Unknown fields are ``None``."""
    name = filename.rsplit("/", 1)[-1]
    low = name.lower()
    flat = re.sub(r"[^a-z0-9]", "", low)

    location = None
    for loc, tokens in CIBSE_LOCATIONS.items():
        if any(tok.replace("_", "") in flat for tok in tokens):
            location = loc
            break
    if location is None and "london" in flat:
        location = "London (Heathrow)"

    kind = None
    m = re.search(r"dsy[\s_\-]*([123])(?!\d)", low)
    if m:
        kind = f"DSY{m.group(1)}"
    elif re.search(r"dsy", low):
        kind = "DSY1"
    elif re.search(r"(?<![a-z])try(?![a-z])|_try|try_", low):
        kind = "TRY"

    period = None
    m = re.search(r"(20[2-8]0)s?", low)
    if m:
        period = f"{m.group(1)}s"
    elif kind is not None:
        period = "Baseline"

    emissions = None
    m = re.search(r"(low|medium|med|high)", low)
    if m and period not in (None, "Baseline"):
        emissions = {"low": "Low", "med": "Medium", "medium": "Medium", "high": "High"}[m.group(1)]

    percentile = None
    m = re.search(r"(?:low|medium|med|high)[\s_\-]*(10|50|90)(?!\d)", low) or re.search(
        r"(?<!\d)(10|50|90)(?:th|pc|p|%|pct|percentile)", low
    )
    if m and period not in (None, "Baseline"):
        percentile = int(m.group(1))

    return {
        "location": location,
        "kind": kind,
        "period": period,
        "emissions": emissions,
        "percentile": percentile,
    }


def reference_label(meta: dict) -> str:
    """Human label such as 'DSY1 · 2050s · High emissions · 50th percentile'."""
    parts = [meta.get("kind") or "Weather file"]
    period = meta.get("period")
    if period:
        parts.append("current climate" if period == "Baseline" else period)
    if meta.get("emissions"):
        parts.append(f"{meta['emissions']} emissions")
    if meta.get("percentile"):
        parts.append(f"{_ordinal(int(meta['percentile']))} percentile")
    return " · ".join(parts)


def _ordinal(n: int) -> str:
    suffix = "th" if 10 <= n % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"

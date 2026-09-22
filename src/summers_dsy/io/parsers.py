"""Parsers for weather files: EPW, CIBSE/generic CSV and Met Office MIDAS Open (BADC-CSV)."""

from __future__ import annotations

import io
import re
from pathlib import Path

import pandas as pd

from ..model import WeatherSeries, nominal_index, to_hourly
from .naming import parse_reference_name

EPW_MISSING_DRY_BULB = 99.9


class WeatherFileError(ValueError):
    """Raised when a file cannot be interpreted as hourly weather."""


def _text(source) -> str:
    if isinstance(source, Path):
        raw = source.read_bytes()
    elif isinstance(source, str):
        if "\n" not in source and len(source) < 1024 and Path(source).is_file():
            raw = Path(source).read_bytes()
        else:
            return source
    elif isinstance(source, bytes):
        raw = source
    elif hasattr(source, "read"):
        raw = source.read()
        if isinstance(raw, str):
            return raw
    else:
        raise WeatherFileError("Unsupported input: pass a path, bytes or a file-like object")
    for enc in ("utf-8-sig", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise WeatherFileError("Could not decode file")


# --------------------------------------------------------------------------- EPW


def parse_epw(source, name: str = "EPW file") -> WeatherSeries:
    """EnergyPlus weather file. CIBSE distributes its TRY/DSY sets in this format."""
    text = _text(source)
    lines = text.splitlines()
    if not lines or not lines[0].upper().startswith("LOCATION"):
        raise WeatherFileError("Not an EPW file (first line should start with LOCATION)")
    loc = lines[0].split(",")
    meta = {
        "site": loc[1].strip() if len(loc) > 1 else "",
        "wmo": loc[5].strip() if len(loc) > 5 else "",
        "latitude": _float(loc[6]) if len(loc) > 6 else None,
        "longitude": _float(loc[7]) if len(loc) > 7 else None,
        "format": "epw",
    }
    df = pd.read_csv(io.StringIO("\n".join(lines[8:])), header=None, usecols=range(9))
    df.columns = ["year", "month", "day", "hour", "minute", "flags", "dry_bulb", "dew_point", "rel_humidity"]
    df["dry_bulb"] = df["dry_bulb"].where(df["dry_bulb"] < EPW_MISSING_DRY_BULB)
    return _from_columns(df, name, "epw", meta)


def _float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _from_columns(df: pd.DataFrame, name: str, source: str, meta: dict) -> WeatherSeries:
    """Build a series from year?/month/day/hour-ending columns, choosing typical vs actual year."""
    ref = parse_reference_name(name)
    years = df["year"].dropna().unique() if "year" in df else []
    single_year = len(years) == 1
    typical = ref.get("kind") is not None or not single_year
    meta = {**meta, **{k: v for k, v in ref.items() if v is not None}}
    if typical:
        index, keep = nominal_index(df["month"], df["day"], df["hour"])
        data = df.loc[keep, [c for c in ("dry_bulb", "rel_humidity") if c in df]].set_axis(index)
        if single_year:
            meta["source_year"] = int(years[0])
        meta["typical_year"] = True
    else:
        ts = pd.to_datetime(
            {"year": df["year"].astype(int), "month": df["month"], "day": df["day"], "hour": df["hour"] - 1}
        )
        data = df[[c for c in ("dry_bulb", "rel_humidity") if c in df]].set_axis(pd.DatetimeIndex(ts))
        meta["typical_year"] = False
    return WeatherSeries(name, data, source, meta)


# --------------------------------------------------------------------------- CSV

_TEMP_NAMES = [
    "dry_bulb",
    "drybulb",
    "dry_bulb_temperature",
    "dry_bulb_temperature_c",
    "dbt",
    "db",
    "air_temperature",
    "temperature_2m",
    "temp",
    "temperature",
    "t",
    "tdb",
]
_DATETIME_NAMES = ["datetime", "date_time", "timestamp", "ob_time", "time", "date"]


def _norm(col: str) -> str:
    col = re.sub(r"\(.*?\)|\[.*?\]", "", str(col).strip().lower())
    return re.sub(r"[^a-z0-9]+", "_", col).strip("_")


def _find_temp_column(cols: list[str]) -> str | None:
    normed = {_norm(c): c for c in cols}
    for cand in _TEMP_NAMES:
        if cand in normed:
            return normed[cand]
    for n, c in normed.items():
        if "dry" in n and ("bulb" in n or "temp" in n):
            return c
    for n, c in normed.items():
        if "temp" in n and "dew" not in n and "wet" not in n and "soil" not in n:
            return c
    return None


def _header_row(lines: list[str]) -> int:
    for i, line in enumerate(lines[:80]):
        cells = [_norm(c) for c in re.split(r"[,;\t]", line)]
        has_temp = _find_temp_column(cells) is not None
        has_time = {"month", "day", "hour"} <= set(cells) or any(c in _DATETIME_NAMES for c in cells)
        if has_temp and has_time:
            return i
    raise WeatherFileError(
        "Could not find a header row with a temperature column and date/time columns. "
        "Expected e.g. 'Month,Day,Hour,Dry bulb' or 'datetime,temperature'."
    )


def parse_csv(source, name: str = "CSV file") -> WeatherSeries:
    """CIBSE-style or generic CSV with a header row.

    Accepts either ``[Year,] Month, Day, Hour`` columns (hours 1-24 are taken as
    hour-ending, 0-23 as hour-beginning) or a single date-time column, plus a
    dry-bulb/temperature column in degC.
    """
    text = _text(source)
    lines = text.splitlines()
    start = _header_row(lines)
    sep = max([",", ";", "\t"], key=lines[start].count)
    df = pd.read_csv(io.StringIO("\n".join(lines[start:])), sep=sep, skipinitialspace=True)
    df = df.dropna(how="all")
    normed = {_norm(c): c for c in df.columns}
    temp_col = _find_temp_column(list(df.columns))
    temps = pd.to_numeric(df[temp_col], errors="coerce")
    meta = {"format": "csv", "temperature_column": str(temp_col)}

    if {"month", "day", "hour"} <= set(normed):
        hours = pd.to_numeric(df[normed["hour"]], errors="coerce")
        cols = pd.DataFrame(
            {
                "month": pd.to_numeric(df[normed["month"]], errors="coerce"),
                "day": pd.to_numeric(df[normed["day"]], errors="coerce"),
                # Normalise to hour-ending for _from_columns.
                "hour": hours if hours.min() >= 1 else hours + 1,
                "dry_bulb": temps,
            }
        )
        if "year" in normed:
            cols["year"] = pd.to_numeric(df[normed["year"]], errors="coerce")
        cols = cols.dropna(subset=["month", "day", "hour"])
        return _from_columns(cols, name, "csv", meta)

    dt_col = next(normed[c] for c in _DATETIME_NAMES if c in normed)
    stamps = pd.to_datetime(df[dt_col], errors="coerce", utc=True, dayfirst=_looks_dayfirst(df[dt_col]))
    series = pd.Series(temps.to_numpy(), index=stamps.dt.tz_localize(None)).dropna()
    series = series[series.index.notna()]
    hourly = to_hourly(series)
    meta["typical_year"] = False
    return WeatherSeries(name, hourly.to_frame("dry_bulb"), "csv", meta)


def _looks_dayfirst(values: pd.Series) -> bool:
    sample = str(values.dropna().iloc[0]) if values.notna().any() else ""
    return bool(re.match(r"^\d{1,2}/\d{1,2}/\d{4}", sample))


# --------------------------------------------------------------------------- MIDAS


def parse_midas(source, name: str = "MIDAS Open") -> WeatherSeries:
    """Met Office MIDAS Open hourly weather observations (BADC-CSV).

    One file per station-year from CEDA's ``uk-hourly-weather-obs`` dataset.
    Uses ``ob_time`` and ``air_temperature``; duplicate reports for an hour are averaged.
    """
    text = _text(source)
    lines = text.splitlines()
    meta: dict = {"format": "midas"}
    data_start = None
    for i, line in enumerate(lines):
        stripped = line.strip().lower()
        if stripped == "data":
            data_start = i + 1
            break
        parts = [p.strip() for p in line.split(",")]
        if parts[0] == "observation_station" and len(parts) > 2:
            meta["site"] = parts[2]
        elif parts[0] == "location" and len(parts) > 3:
            meta["latitude"], meta["longitude"] = _float(parts[2]), _float(parts[3])
    if data_start is None:
        raise WeatherFileError("Not a BADC-CSV file (no 'data' line found)")
    end = next((i for i in range(data_start, len(lines)) if lines[i].strip().lower() == "end data"), len(lines))
    df = pd.read_csv(io.StringIO("\n".join(lines[data_start:end])), skipinitialspace=True)
    if "ob_time" not in df or "air_temperature" not in df:
        raise WeatherFileError("MIDAS file lacks ob_time/air_temperature columns")
    t = pd.Series(
        pd.to_numeric(df["air_temperature"], errors="coerce").to_numpy(),
        index=pd.to_datetime(df["ob_time"], errors="coerce"),
    )
    t = t[t.index.notna()]
    meta["typical_year"] = False
    return WeatherSeries(meta.get("site", name) or name, to_hourly(t).to_frame("dry_bulb"), "midas", meta)


# --------------------------------------------------------------------------- dispatch


def load_weather_file(source, filename: str) -> WeatherSeries:
    """Pick the right parser from the file's name and content."""
    text = _text(source)
    head = text[:4000]
    stem = Path(filename).stem
    if head.upper().startswith("LOCATION"):
        return parse_epw(text, stem)
    if re.search(r"^\s*Conventions\s*,\s*G\s*,\s*BADC-CSV", head, re.M | re.I) or re.search(r"^data\s*$", text, re.M):
        return parse_midas(text, stem)
    return parse_csv(text, stem)


def series_to_csv(series: WeatherSeries) -> bytes:
    """Export hourly dry-bulb as a simple CSV."""
    out = series.data[["dry_bulb"]].copy()
    out.index.name = "datetime"
    return out.to_csv(float_format="%.2f").encode()


__all__ = [
    "WeatherFileError",
    "parse_epw",
    "parse_csv",
    "parse_midas",
    "load_weather_file",
    "series_to_csv",
]

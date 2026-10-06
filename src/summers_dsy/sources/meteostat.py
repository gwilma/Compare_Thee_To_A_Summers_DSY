"""Meteostat: weather stations and multi-year hourly temperature downloads.

Hourly data come from Meteostat's annual files, one gzipped CSV per station and year::

    https://data.meteostat.net/hourly/{year}/{station}.csv.gz

The files have a header row: four time columns (year, month, day, hour, UTC), then one column per
parameter (``temp`` is air temperature in degC), each with a ``<parameter>_source`` column naming
the provider of each value. Gaps in the observations are filled with model forecasts
(``dwd_mosmix``, ``metno_forecast``); those values are dropped unless ``include_model`` is set.
No API key is needed. The retired ``bulk.meteostat.net/v2`` files (no header) are still read as a
fallback.

The station list is Meteostat's open station directory (stations.db, CC BY 4.0), reduced to the
stations with hourly temperature observations. A UK snapshot ships with the package; it can be
refreshed from Meteostat.
"""

from __future__ import annotations

import datetime as dt
import gzip
import io
import json
import sqlite3
from functools import lru_cache
from pathlib import Path

import numpy as np
import pandas as pd

from ..model import WeatherSeries, to_hourly
from .http import SourceError, data_dir, get_bytes

HOURLY_URL = "https://data.meteostat.net/hourly/{year}/{station}.csv.gz"
LEGACY_HOURLY_URL = "https://bulk.meteostat.net/v2/hourly/{year}/{station}.csv.gz"
STATIONS_DB_URLS = [
    "https://data.meteostat.net/stations.db",
    "https://raw.githubusercontent.com/meteostat/weather-stations/master/stations.db",
]
POSTCODE_URL = "https://api.postcodes.io/postcodes/{postcode}"

#: Providers whose values are model forecasts rather than observations.
MODEL_SOURCES = {"dwd_mosmix", "metno_forecast"}
#: Providers of hourly observations (used to describe each station's hourly coverage).
HOURLY_OBS_PROVIDERS = ("isd_lite", "metar", "dwd_hourly", "dwd_poi", "eccc_hourly", "gsa_hourly", "gsa_synop")

LEGACY_COLUMNS = ["date", "hour", "temp", "dwpt", "rhum", "prcp", "snow", "wdir", "wspd", "wpgt", "pres", "tsun", "coco"]
STATION_COLUMNS = ["id", "name", "region", "latitude", "longitude", "elevation", "wmo", "icao", "hourly_start", "hourly_end"]
BUNDLED_STATIONS = Path(__file__).resolve().parent / "data" / "meteostat_stations_uk.csv"
ATTRIBUTION = "Weather station data © Meteostat (meteostat.net), CC BY 4.0."

ONE_DAY = 86400


# ------------------------------------------------------------------ stations


def stations_from_db(path: str | Path, countries: list[str] | None = None) -> list[dict]:
    """Stations with hourly temperature observations, from a Meteostat stations.db."""
    countries = countries or ["GB"]
    con = sqlite3.connect(f"file:{Path(path).resolve()}?mode=ro", uri=True)
    try:
        providers = ",".join("?" * len(HOURLY_OBS_PROVIDERS))
        marks = ",".join("?" * len(countries))
        rows = con.execute(
            f"""
            SELECT s.id, COALESCE(n.name, s.id), s.region, s.latitude, s.longitude, s.elevation,
                   (SELECT value FROM identifiers WHERE station = s.id AND key = 'wmo'),
                   (SELECT value FROM identifiers WHERE station = s.id AND key = 'icao'),
                   MIN(i.start), MAX(i."end")
            FROM stations s
            JOIN inventory i ON i.station = s.id AND i.parameter = 'temp' AND i.provider IN ({providers})
            LEFT JOIN names n ON n.station = s.id AND n.language = 'en'
            WHERE s.country IN ({marks})
            GROUP BY s.id
            ORDER BY COALESCE(n.name, s.id)
            """,
            [*HOURLY_OBS_PROVIDERS, *countries],
        ).fetchall()
    finally:
        con.close()
    return [dict(zip(STATION_COLUMNS, r)) for r in rows]


def _frame(records: list[dict] | pd.DataFrame) -> pd.DataFrame:
    df = pd.DataFrame(records, columns=STATION_COLUMNS) if not isinstance(records, pd.DataFrame) else records.copy()
    df["id"] = df["id"].astype(str)
    df["start_year"] = pd.to_datetime(df["hourly_start"], errors="coerce").dt.year
    df["end_year"] = pd.to_datetime(df["hourly_end"], errors="coerce").dt.year
    df["label"] = df["name"].astype(str) + " (" + df["id"] + ")"
    return df.reset_index(drop=True)


@lru_cache(maxsize=1)
def bundled_stations() -> pd.DataFrame:
    """The UK station snapshot shipped with the package."""
    return _frame(pd.read_csv(BUNDLED_STATIONS, dtype={"id": str, "wmo": str, "icao": str}))


def load_stations() -> pd.DataFrame:
    """The refreshed station list if one has been downloaded, else the bundled UK snapshot."""
    cached = data_dir() / "meteostat_stations.csv"
    if cached.exists():
        return _frame(pd.read_csv(cached, dtype={"id": str, "wmo": str, "icao": str}))
    return bundled_stations()


def refresh_stations(countries: list[str] | None = None) -> pd.DataFrame:
    """Download Meteostat's current station database and rebuild the local station list."""
    db = data_dir() / "meteostat_stations.db"
    last_error = None
    for url in STATIONS_DB_URLS:
        try:
            db.write_bytes(get_bytes(url, max_age_s=7 * ONE_DAY))
            break
        except SourceError as exc:
            last_error = exc
    else:
        raise SourceError(f"Could not download the Meteostat station list: {last_error}")
    df = _frame(stations_from_db(db, countries))
    df[STATION_COLUMNS].to_csv(data_dir() / "meteostat_stations.csv", index=False)
    return df


def search_stations(stations: pd.DataFrame, text: str) -> pd.DataFrame:
    """Stations whose name, id, WMO or ICAO code contains ``text`` (case-insensitive)."""
    text = text.strip().lower()
    if not text:
        return stations
    hay = (stations["name"].astype(str) + " " + stations["id"] + " " + stations["wmo"].fillna("").astype(str)
           + " " + stations["icao"].fillna("").astype(str)).str.lower()
    return stations[hay.str.contains(text, regex=False)]


def haversine_km(lat1, lon1, lat2, lon2):
    lat1, lon1, lat2, lon2 = map(np.radians, (lat1, lon1, lat2, lon2))
    a = np.sin((lat2 - lat1) / 2) ** 2 + np.cos(lat1) * np.cos(lat2) * np.sin((lon2 - lon1) / 2) ** 2
    return 6371.0 * 2 * np.arcsin(np.sqrt(a))


def nearest_stations(stations: pd.DataFrame, lat: float, lon: float, n: int = 10) -> pd.DataFrame:
    """The ``n`` closest stations, with a ``distance_km`` column."""
    out = stations.copy()
    out["distance_km"] = haversine_km(lat, lon, out["latitude"].to_numpy(), out["longitude"].to_numpy())
    return out.nsmallest(n, "distance_km")


def geocode_postcode(postcode: str) -> tuple[float, float, str]:
    """Latitude, longitude and a description for a UK postcode (postcodes.io, free, no key)."""
    pc = "".join(postcode.split()).upper()
    if not pc:
        raise ValueError("Enter a postcode")
    raw = get_bytes(POSTCODE_URL.format(postcode=pc), max_age_s=30 * ONE_DAY, not_found_ok=True)
    if raw is None:
        raise SourceError(f"Postcode {postcode} was not found")
    res = json.loads(raw).get("result") or {}
    if res.get("latitude") is None:
        raise SourceError(f"Postcode {postcode} has no location")
    where = ", ".join(x for x in (res.get("admin_district"), res.get("region") or res.get("country")) if x)
    return float(res["latitude"]), float(res["longitude"]), f"{res.get('postcode', pc)} ({where})"


# ------------------------------------------------------------------ hourly data


def parse_hourly(raw: bytes) -> pd.DataFrame:
    """Parse a Meteostat hourly file (gzipped or plain) into ``temp`` and ``source`` columns.

    Handles the current format (header row; year, month, day, hour, then parameters and
    ``*_source`` columns) and the retired v2 format (no header; date, hour, temp, ...).
    """
    if raw[:2] == b"\x1f\x8b":
        raw = gzip.decompress(raw)
    text = raw.decode("utf-8-sig")
    first = text.split("\n", 1)[0].strip().lower()
    if "temp" in [c.strip().strip('"') for c in first.split(",")]:
        df = pd.read_csv(io.StringIO(text))
        df.columns = [c.strip().lower() for c in df.columns]
        if {"year", "month", "day", "hour"} <= set(df.columns):
            idx = pd.to_datetime(df[["year", "month", "day", "hour"]])
        elif "time" in df.columns:
            idx = pd.to_datetime(df["time"])
        else:  # four leading time columns, whatever they are called
            t = df.iloc[:, :4].set_axis(["year", "month", "day", "hour"], axis=1)
            idx = pd.to_datetime(t)
        source = df["temp_source"].astype(str) if "temp_source" in df.columns else pd.Series("", index=df.index)
        temps = pd.to_numeric(df["temp"], errors="coerce")
    else:
        df = pd.read_csv(io.StringIO(text), header=None, names=LEGACY_COLUMNS)
        idx = pd.to_datetime(df["date"]) + pd.to_timedelta(df["hour"], unit="h")
        temps = pd.to_numeric(df["temp"], errors="coerce")
        source = pd.Series("", index=df.index)
    out = pd.DataFrame({"temp": temps.to_numpy(), "source": source.fillna("").to_numpy()}, index=pd.DatetimeIndex(idx))
    return out[out["temp"].notna()].sort_index()


def is_model(source: pd.Series) -> pd.Series:
    """True where the value came from a model forecast rather than an observation."""
    s = source.astype(str).str.lower()
    return s.apply(lambda v: any(m in v for m in MODEL_SOURCES))


def _max_age(year: int) -> float | None:
    return ONE_DAY if year >= dt.date.today().year - 1 else None


def get_year(station: str, year: int) -> pd.DataFrame | None:
    """One station-year from Meteostat (current endpoint, then the legacy one), or None if absent."""
    for url in (HOURLY_URL, LEGACY_HOURLY_URL):
        try:
            raw = get_bytes(url.format(year=year, station=station), max_age_s=_max_age(year), not_found_ok=True)
        except SourceError:
            raw = None
        if raw:
            return parse_hourly(raw)
    return None


def fetch_years(station: str, start_year: int, end_year: int, include_model: bool = False,
                name: str | None = None, progress=None) -> WeatherSeries:
    """Hourly temperature for whole calendar years, plus the preceding December for spin-up.

    ``progress(done, total, year)`` is called after each file. The result's ``meta`` records, per
    year, the share of hours that were model-filled (and dropped unless ``include_model``).
    """
    if end_year < start_year:
        raise ValueError("end year precedes start year")
    years = list(range(start_year - 1, end_year + 1))
    parts, model_share, missing = [], {}, []
    for k, year in enumerate(years, start=1):
        df = get_year(station, year)
        if df is not None and year == start_year - 1:
            df = df[df.index.month == 12]  # only the December, to warm up the running mean
        if df is None or df.empty:
            if year >= start_year:
                missing.append(year)
        else:
            model = is_model(df["source"])
            if year >= start_year:
                model_share[year] = float(model.mean())
            parts.append(df["temp"] if include_model else df.loc[~model, "temp"])
        if progress:
            progress(k, len(years), year)
    if not parts:
        raise SourceError(f"Meteostat has no hourly data for station {station} in {start_year}–{end_year}")
    temps = pd.concat(parts).sort_index()
    temps = temps[~temps.index.duplicated(keep="first")]
    hourly = to_hourly(temps)
    meta = {"station": station, "typical_year": False, "years": list(range(start_year, end_year + 1)),
            "model_share": model_share, "missing_years": missing, "include_model": include_model,
            "attribution": "Hourly data © Meteostat (meteostat.net)"}
    return WeatherSeries(name or f"Meteostat {station}", hourly.to_frame("dry_bulb"), "meteostat", meta)


def station_location_guess(station_name: str) -> str | None:
    """A CIBSE location the station name suggests, e.g. 'Manchester Airport' -> 'Manchester'."""
    from ..io.naming import parse_reference_name

    return parse_reference_name(station_name.replace(" ", "_")).get("location")


__all__ = [
    "ATTRIBUTION",
    "HOURLY_URL",
    "MODEL_SOURCES",
    "bundled_stations",
    "fetch_years",
    "geocode_postcode",
    "get_year",
    "haversine_km",
    "is_model",
    "load_stations",
    "nearest_stations",
    "parse_hourly",
    "refresh_stations",
    "search_stations",
    "station_location_guess",
    "stations_from_db",
]

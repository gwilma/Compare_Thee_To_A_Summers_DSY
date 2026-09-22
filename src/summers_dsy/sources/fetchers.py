"""Fetch hourly observed temperature from public sources.

All fetchers return a :class:`WeatherSeries` on a complete hourly UTC index
covering whole calendar years, so the running mean has the spring spin-up it
needs before the 1 May season start.

Sources
-------
* **Open-Meteo historical API** (ERA5 / ERA5-Land reanalysis): free, no key,
  gridded, 1940 onwards. Reanalysis smooths peaks and urban heat islands.
* **Meteostat bulk data**: free, no key; station observations assembled from
  NOAA ISD and DWD, with gaps filled from models.
* **NOAA Integrated Surface Database** (global-hourly CSV): free, no key; raw
  synoptic/METAR reports for WMO stations.
* **Met Office MIDAS Open** (CEDA): the authoritative UK station record. Needs
  a free CEDA account; set ``CEDA_TOKEN`` (an access token) to download, or
  upload the BADC-CSV files instead.
"""

from __future__ import annotations

import datetime as dt
import gzip
import io
import json
import os

import numpy as np
import pandas as pd

from ..io.parsers import parse_midas
from ..model import WeatherSeries, to_hourly
from .http import SourceError, get_bytes

OPEN_METEO_URL = "https://archive-api.open-meteo.com/v1/archive"
METEOSTAT_URL = "https://bulk.meteostat.net/v2/hourly/{year}/{station}.csv.gz"
NOAA_ISD_URL = "https://www.ncei.noaa.gov/data/global-hourly/access/{year}/{station}.csv"
MIDAS_URL = (
    "https://dap.ceda.ac.uk/badc/ukmo-midas-open/data/uk-hourly-weather-obs/"
    "dataset-version-{version}/{county}/{src:05d}_{site}/qc-version-1/"
    "midas-open_uk-hourly-weather-obs_dv-{version}_{county}_{src:05d}_{site}_qcv-1_{year}.csv"
)
MIDAS_DEFAULT_VERSION = "202407"

METEOSTAT_COLUMNS = ["date", "hour", "temp", "dwpt", "rhum", "prcp", "snow", "wdir", "wspd", "wpgt", "pres", "tsun", "coco"]

#: ISD quality codes treated as unusable (suspect or erroneous).
ISD_BAD_QUALITY = set("2367")

ONE_DAY = 86400


def _max_age(year: int) -> float | None:
    """Past years are immutable enough to cache forever; refresh the current year daily."""
    return ONE_DAY if year >= dt.date.today().year else None


def _years(start: int, end: int) -> list[int]:
    if end < start:
        raise ValueError("end year precedes start year")
    return list(range(start, end + 1))


def _finish(name: str, source: str, temps: pd.Series, meta: dict, years: list[int]) -> WeatherSeries:
    if temps.dropna().empty:
        raise SourceError(f"{source} returned no temperature data for {name} {years[0]}-{years[-1]}")
    hourly = to_hourly(temps.sort_index())
    hourly = hourly[(hourly.index.year >= years[0]) & (hourly.index.year <= years[-1])]
    meta = {**meta, "typical_year": False, "years": years}
    return WeatherSeries(name, hourly.to_frame("dry_bulb"), source, meta)


# ----------------------------------------------------------------- Open-Meteo


def fetch_open_meteo(lat: float, lon: float, start_year: int, end_year: int, model: str = "era5", name: str | None = None) -> WeatherSeries:
    """Hourly 2 m temperature from the Open-Meteo historical weather API."""
    years = _years(start_year, end_year)
    end = min(dt.date(end_year, 12, 31), dt.date.today() - dt.timedelta(days=6))
    params = {
        "latitude": round(lat, 4),
        "longitude": round(lon, 4),
        "start_date": f"{start_year}-01-01",
        "end_date": end.isoformat(),
        "hourly": "temperature_2m,relative_humidity_2m",
        "timezone": "GMT",
        "models": model,
    }
    payload = json.loads(get_bytes(OPEN_METEO_URL, params=params, max_age_s=_max_age(end_year)))
    if payload.get("error"):
        raise SourceError(f"Open-Meteo: {payload.get('reason', 'unknown error')}")
    hourly = payload["hourly"]
    idx = pd.to_datetime(hourly["time"])
    temps = pd.Series(np.asarray(hourly["temperature_2m"], dtype=float), index=idx)
    meta = {"latitude": lat, "longitude": lon, "model": model, "grid_elevation": payload.get("elevation")}
    return _finish(name or f"ERA5 {lat:.2f},{lon:.2f}", "open-meteo", temps, meta, years)


# ------------------------------------------------------------------ Meteostat


def parse_meteostat(raw: bytes) -> pd.Series:
    """Parse one Meteostat bulk hourly file (gzipped CSV, no header)."""
    text = gzip.decompress(raw).decode()
    df = pd.read_csv(io.StringIO(text), header=None, names=METEOSTAT_COLUMNS)
    idx = pd.to_datetime(df["date"]) + pd.to_timedelta(df["hour"], unit="h")
    return pd.Series(pd.to_numeric(df["temp"], errors="coerce").to_numpy(), index=idx)


def fetch_meteostat(station: str, start_year: int, end_year: int, name: str | None = None) -> WeatherSeries:
    """Hourly temperature from Meteostat's bulk endpoint (one gzipped CSV per station-year)."""
    years = _years(start_year, end_year)
    parts = []
    for year in years:
        raw = get_bytes(METEOSTAT_URL.format(year=year, station=station), max_age_s=_max_age(year), not_found_ok=True)
        if raw:
            parts.append(parse_meteostat(raw))
    if not parts:
        raise SourceError(f"Meteostat has no hourly data for station {station} in {start_year}-{end_year}")
    return _finish(name or f"Meteostat {station}", "meteostat", pd.concat(parts), {"station": station}, years)


# ------------------------------------------------------------------- NOAA ISD


def parse_isd(raw: bytes) -> pd.Series:
    """Parse an ISD global-hourly CSV: ``TMP`` is tenths of degC with a quality flag, e.g. ``+0215,1``."""
    df = pd.read_csv(io.BytesIO(raw), usecols=["DATE", "TMP"], dtype={"TMP": str})
    parts = df["TMP"].str.split(",", expand=True)
    value = pd.to_numeric(parts[0], errors="coerce")
    quality = parts[1].fillna("") if parts.shape[1] > 1 else pd.Series("", index=df.index)
    ok = (value.abs() < 9999) & ~quality.isin(ISD_BAD_QUALITY)
    return pd.Series((value / 10.0).where(ok).to_numpy(), index=pd.to_datetime(df["DATE"])).dropna()


def fetch_noaa_isd(station: str, start_year: int, end_year: int, name: str | None = None) -> WeatherSeries:
    """Hourly temperature from NOAA ISD global-hourly files (``station`` = USAF+WBAN, e.g. 03772099999)."""
    years = _years(start_year, end_year)
    parts = []
    for year in years:
        raw = get_bytes(NOAA_ISD_URL.format(year=year, station=station), max_age_s=_max_age(year), not_found_ok=True)
        if raw:
            parts.append(parse_isd(raw))
    if not parts:
        raise SourceError(f"NOAA ISD has no data for station {station} in {start_year}-{end_year}")
    return _finish(name or f"ISD {station}", "noaa-isd", pd.concat(parts), {"station": station}, years)


# ---------------------------------------------------------------- MIDAS Open


def fetch_midas(
    county: str,
    src_id: int,
    site: str,
    start_year: int,
    end_year: int,
    token: str | None = None,
    version: str = MIDAS_DEFAULT_VERSION,
    name: str | None = None,
) -> WeatherSeries:
    """Hourly air temperature from Met Office MIDAS Open on CEDA (needs a CEDA access token)."""
    token = token or os.environ.get("CEDA_TOKEN")
    if not token:
        raise SourceError(
            "MIDAS Open downloads need a CEDA access token (set CEDA_TOKEN or enter one). "
            "Alternatively download the yearly BADC-CSV files from CEDA and upload them."
        )
    years = _years(start_year, end_year)
    headers = {"Authorization": f"Bearer {token}"}
    parts = []
    for year in years:
        url = MIDAS_URL.format(version=version, county=county, src=src_id, site=site, year=year)
        raw = get_bytes(url, headers=headers, max_age_s=_max_age(year), not_found_ok=True)
        if raw:
            parts.append(parse_midas(raw).dry_bulb)
    if not parts:
        raise SourceError(f"MIDAS Open has no files for {site} ({src_id}) in {start_year}-{end_year}, dataset version {version}")
    meta = {"station": f"{src_id:05d}_{site}", "county": county, "dataset_version": version}
    return _finish(name or f"MIDAS {site}", "midas", pd.concat(parts), meta, years)


def combine_uploads(series_list: list[WeatherSeries], name: str) -> WeatherSeries:
    """Join several uploaded observed files (e.g. one MIDAS file per year) into one record."""
    temps = pd.concat([s.dry_bulb for s in series_list]).sort_index()
    temps = temps[~temps.index.duplicated(keep="first")]
    years = sorted(set(temps.index.year))
    return _finish(name, series_list[0].source, temps, dict(series_list[0].meta), years)

"""Overheating metrics for a weather series.

Weighted cooling degree hours follow Eames (2016) and Liu, Kershaw, Eames &
Coley (2016), the basis of the CIBSE 2016 probabilistic Design Summer Years::

    WCDH  = sum over season hours of max(0, T - Tcomf)^2                (K^2 h)
    TWCDH = sum over season hours of max(0, T - (Tcomf + dT_region))^2
    SWCDH = sum over season hours of max(0, T - T_static)^2

* ``Tcomf`` is the TM52 adaptive comfort temperature (see :mod:`summers_dsy.comfort`).
* ``T_static`` is a fixed regional threshold, the 93rd centile of the regional
  hourly temperature in Eames' method.
* ``dT_region`` is a regional adjustment to the comfort temperature.

The regional values are location-specific, so they are inputs here. The
helpers :func:`static_threshold` and :func:`threshold_offset` derive them from
a baseline series using the 93rd-centile convention (the ``offset`` is the
93rd centile of ``T - Tcomf``, i.e. the TWCDH analogue of the SWCDH threshold).

Warm events distinguish the DSY types: DSY2 has a short, intense warm spell
and DSY3 a long, sustained one. An event is a run of consecutive days with a
positive daily WCDH (optionally above a minimum), characterised by its
duration, severity (event WCDH total) and intensity (peak daily WCDH).
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field

import numpy as np
import pandas as pd

from .comfort import daily_comfort, hourly_from_daily
from .model import WeatherSeries

#: CIBSE TM59 (and TM52) assessment period: 1 May - 30 September.
TM59_SEASON = ((5, 1), (9, 30))
#: Original TM49 DSY ranking period: April - September.
TM49_SEASON = ((4, 1), (9, 30))

STATIC_CENTILE = 93.0


@dataclass
class AnalysisConfig:
    """Settings for the running mean, comfort temperature and metric thresholds."""

    season_start: tuple[int, int] = TM59_SEASON[0]
    season_end: tuple[int, int] = TM59_SEASON[1]
    alpha: float = 0.8
    trm_method: str = "exponential"  # or "seven_day"
    category: str = "II"
    clamp_trm: bool = False
    #: SWCDH fixed threshold (degC). None -> derived from the series itself.
    static_threshold: float | None = None
    #: TWCDH offset added to Tcomf (K). None -> derived from the series itself.
    twcdh_offset: float | None = None
    #: Minimum daily WCDH (K^2 h) for a day to count as part of a warm event.
    event_min_daily_wcdh: float = 0.0
    #: Daily maximum threshold for counting hot days (degC).
    hot_day_threshold: float = 28.0
    #: Daily minimum threshold for counting warm nights (degC).
    warm_night_threshold: float = 16.0
    #: Night-time window (hour-beginning, file time). TM59 treats 22:00-07:00 as bedroom night-time.
    night_start_hour: int = 22
    night_end_hour: int = 7
    #: Threshold for the nightly mean (degC). TM59:2026 limits the mean bedroom temperature at night to
    #: below 27 degC, with no more than four exceedance nights between May and September.
    night_threshold: float = 27.0
    #: Nights a night needs at least this many valid hours to have a mean.
    night_min_hours: int = 7

    def to_dict(self) -> dict:
        return asdict(self)


def season_mask(index: pd.DatetimeIndex, start=TM59_SEASON[0], end=TM59_SEASON[1]) -> np.ndarray:
    """Boolean mask for timestamps between (month, day) ``start`` and ``end`` inclusive."""
    md = index.month * 100 + index.day
    lo, hi = start[0] * 100 + start[1], end[0] * 100 + end[1]
    if lo <= hi:
        return np.asarray((md >= lo) & (md <= hi))
    return np.asarray((md >= lo) | (md <= hi))


def weighted_degree_hours(temps: pd.Series, threshold) -> pd.Series:
    """Hourly squared exceedance ``max(0, T - threshold)^2``."""
    excess = (temps - threshold).clip(lower=0)
    return excess**2


def static_threshold(series: WeatherSeries, config: AnalysisConfig | None = None, centile: float = STATIC_CENTILE) -> float:
    """93rd centile of in-season hourly dry-bulb: the SWCDH threshold."""
    config = config or AnalysisConfig()
    t = series.dry_bulb[season_mask(series.data.index, config.season_start, config.season_end)]
    return float(np.nanpercentile(t, centile))


def threshold_offset(series: WeatherSeries, config: AnalysisConfig | None = None, centile: float = STATIC_CENTILE) -> float:
    """93rd centile of in-season (T - Tcomf): the TWCDH adjustment to Tcomf."""
    config = config or AnalysisConfig()
    return _offset_centile(hourly_comfort(series, config), config, centile)


def _offset_centile(hourly: pd.DataFrame, config: AnalysisConfig, centile: float = STATIC_CENTILE) -> float:
    season = hourly[season_mask(hourly.index, config.season_start, config.season_end)]
    return float(np.nanpercentile(season["dry_bulb"] - season["t_comf"], centile))


def daily_table(series: WeatherSeries, config: AnalysisConfig | None = None) -> pd.DataFrame:
    """Daily min/max/mean, Trm, Tcomf and upper limit for the whole series."""
    config = config or AnalysisConfig()
    return daily_comfort(
        series.dry_bulb,
        alpha=config.alpha,
        method=config.trm_method,
        cyclic=series.is_typical_year,
        category=config.category,
        clamp=config.clamp_trm,
    )


def hourly_comfort(series: WeatherSeries, config: AnalysisConfig | None = None, daily: pd.DataFrame | None = None) -> pd.DataFrame:
    """Hourly dry-bulb alongside the day's Tcomf and upper limit."""
    config = config or AnalysisConfig()
    if daily is None:
        daily = daily_table(series, config)
    idx = series.data.index
    return pd.DataFrame(
        {
            "dry_bulb": series.dry_bulb,
            "t_comf": hourly_from_daily(daily["t_comf"], idx),
            "t_upper": hourly_from_daily(daily["t_upper"], idx),
        },
        index=idx,
    )


def nightly_means(temps: pd.Series, start_hour: int = 22, end_hour: int = 7, min_hours: int = 7) -> pd.DataFrame:
    """Mean temperature of each night, indexed by the date of the evening the night starts.

    A night runs from ``start_hour`` on day d to ``end_hour`` on day d+1 (hours are hour-beginning,
    so 22:00-07:00 is the nine hours 22, 23, 0, ..., 6). Nights with fewer than ``min_hours`` valid
    hours have no mean.
    """
    hours = temps.index.hour
    if start_hour > end_hour:
        evening, morning = hours >= start_hour, hours < end_hour
        night_date = temps.index.normalize() - pd.to_timedelta(np.where(morning, 1, 0), unit="D")
        in_night = evening | morning
    else:  # a window that does not cross midnight
        in_night = (hours >= start_hour) & (hours < end_hour)
        night_date = temps.index.normalize()
    t = temps[in_night]
    grouped = t.groupby(night_date[in_night])
    out = pd.DataFrame({"night_hours": grouped.count(), "night_mean": grouped.mean()})
    out.loc[out["night_hours"] < min_hours, "night_mean"] = np.nan
    out.index.name = None
    return out


@dataclass
class WarmEvent:
    start: pd.Timestamp
    end: pd.Timestamp
    duration_days: int
    severity: float  # total WCDH over the event, K^2 h
    intensity: float  # peak daily WCDH, K^2 h
    peak_temperature: float


def warm_events(daily_wcdh: pd.Series, daily_tmax: pd.Series, min_daily_wcdh: float = 0.0) -> list[WarmEvent]:
    """Runs of consecutive days whose daily WCDH exceeds ``min_daily_wcdh``."""
    hot = (daily_wcdh > min_daily_wcdh).to_numpy()
    events: list[WarmEvent] = []
    i, n = 0, len(hot)
    while i < n:
        if not hot[i]:
            i += 1
            continue
        j = i
        while j + 1 < n and hot[j + 1]:
            j += 1
        chunk = daily_wcdh.iloc[i : j + 1]
        events.append(
            WarmEvent(
                start=chunk.index[0],
                end=chunk.index[-1],
                duration_days=j - i + 1,
                severity=float(chunk.sum()),
                intensity=float(chunk.max()),
                peak_temperature=float(daily_tmax.iloc[i : j + 1].max()),
            )
        )
        i = j + 1
    return events


@dataclass
class SeasonAnalysis:
    """Everything computed for one series over one season."""

    name: str
    config: AnalysisConfig
    static_threshold: float
    twcdh_offset: float
    daily: pd.DataFrame  # whole-series daily table with daily WCDH columns
    hourly: pd.DataFrame  # in-season hourly table with exceedance columns
    events: list[WarmEvent] = field(default_factory=list)
    metrics: dict = field(default_factory=dict)

    def season_daily(self) -> pd.DataFrame:
        c = self.config
        return self.daily[season_mask(self.daily.index, c.season_start, c.season_end)]

    def events_frame(self) -> pd.DataFrame:
        return pd.DataFrame([asdict(e) for e in self.events])


#: Metric keys with display label, unit and a short description.
METRIC_INFO: dict[str, tuple[str, str, str]] = {
    "wcdh": ("WCDH", "K²h", "Weighted cooling degree hours above the adaptive comfort temperature (Eames 2016)"),
    "twcdh": ("TWCDH", "K²h", "Threshold WCDH above Tcomf + regional offset"),
    "swcdh": ("SWCDH", "K²h", "Static WCDH above the regional 93rd-centile temperature; CIBSE ranks DSY1 (1-in-7 year) by this"),
    "peak_daily_wcdh": ("Peak daily WCDH", "K²h", "Intensity of the most intense day"),
    "max_event_severity": ("Max event severity", "K²h", "WCDH total of the most severe warm event"),
    "max_event_duration": ("Longest warm event", "days", "Longest run of consecutive days with WCDH > 0"),
    "n_events": ("Warm events", "count", "Number of warm events in the season"),
    "hours_above_upper": ("Hours > Tmax", "h", "Hours of outdoor air above Tcomf + category offset (TM52 criterion 1 analogue)"),
    "max_delta_t": ("Max ΔT above Tmax", "K", "Largest exceedance of Tmax (TM52 criterion 3 analogue)"),
    "max_daily_weighted_exceedance": ("Max daily We", "K·h", "Largest daily sum of rounded ΔT above Tmax (TM52 criterion 2 analogue)"),
    "t_max": ("Peak temperature", "°C", "Highest hourly dry-bulb in the season"),
    "mean_daily_max": ("Mean daily max", "°C", "Season average of daily maxima"),
    "mean_daily_min": ("Mean daily min", "°C", "Season average of daily minima"),
    "season_mean": ("Season mean", "°C", "Mean dry-bulb over the season"),
    "apr_sep_mean": ("Apr–Sep mean", "°C", "Mean April–September dry-bulb (original TM49 DSY ranking)"),
    "hot_days": ("Hot days", "days", "Days with maximum at or above the hot-day threshold"),
    "warm_nights": ("Warm nights", "nights", "Days with minimum at or above the warm-night threshold"),
    "cdh_22": ("CDH (base 22 °C)", "K·h", "Unweighted cooling degree hours above 22 °C"),
    "night_max_mean": ("Warmest night (mean)", "°C",
                       "Highest night-time (22:00–07:00) mean outdoor temperature in the season"),
    "nights_above": ("Nights ≥ threshold", "nights",
                     "Nights whose mean outdoor temperature (22:00–07:00) is at or above the night threshold, 27 °C by "
                     "default (TM59:2026 bedroom criterion analogue: no more than 4 such nights)"),
    "mean_t_djf": ("Winter mean (DJF)", "°C",
                   "Mean air temperature, December–February of the analysis year (energy demand: heating)"),
    "mean_t_mam": ("Spring mean (MAM)", "°C", "Mean air temperature, March–May (energy demand)"),
    "mean_t_jja": ("Summer mean (JJA)", "°C", "Mean air temperature, June–August (energy demand: cooling)"),
    "mean_t_son": ("Autumn mean (SON)", "°C", "Mean air temperature, September–November (energy demand)"),
    "jja_mean_daily_max": ("Summer mean daily max (JJA)", "°C",
                           "Mean of daily maximum temperatures, June–August (overheating risk)"),
}

#: Meteorological seasons used for the seasonal means (months; DJF uses the analysis year's own December).
MET_SEASONS = {"djf": (12, 1, 2), "mam": (3, 4, 5), "jja": (6, 7, 8), "son": (9, 10, 11)}

DEFAULT_COMPARISON_METRICS = [
    "wcdh",
    "twcdh",
    "swcdh",
    "peak_daily_wcdh",
    "max_event_severity",
    "max_event_duration",
    "t_max",
    "season_mean",
]


def analyse(series: WeatherSeries, config: AnalysisConfig | None = None) -> SeasonAnalysis:
    """Compute the daily table, hourly exceedances, warm events and summary metrics."""
    config = config or AnalysisConfig()
    daily = daily_table(series, config)
    hourly_all = hourly_comfort(series, config, daily)

    s_thr = config.static_threshold if config.static_threshold is not None else static_threshold(series, config)
    offset = config.twcdh_offset if config.twcdh_offset is not None else _offset_centile(hourly_all, config)

    t = hourly_all["dry_bulb"]
    hourly_all["wcdh"] = weighted_degree_hours(t, hourly_all["t_comf"])
    hourly_all["twcdh"] = weighted_degree_hours(t, hourly_all["t_comf"] + offset)
    hourly_all["swcdh"] = weighted_degree_hours(t, s_thr)
    hourly_all["delta_t"] = t - hourly_all["t_upper"]

    for col in ("wcdh", "twcdh", "swcdh"):
        daily[f"{col}_daily"] = hourly_all[col].resample("D").sum(min_count=1)
    nights = nightly_means(t, config.night_start_hour, config.night_end_hour, config.night_min_hours)
    daily = daily.join(nights)

    in_season = season_mask(hourly_all.index, config.season_start, config.season_end)
    hourly = hourly_all[in_season]
    d_mask = season_mask(daily.index, config.season_start, config.season_end)
    sd = daily[d_mask]

    events = warm_events(sd["wcdh_daily"].fillna(0), sd["t_max"], config.event_min_daily_wcdh)

    # TM52 rounds dT to the nearest whole degree for criteria 1-3 (hours with dT >= 1 K count).
    dt_rounded = np.round(hourly["delta_t"])
    exceed = dt_rounded.where(dt_rounded >= 1, 0)
    daily_we = exceed.resample("D").sum()

    apr_sep = season_mask(hourly_all.index, *TM49_SEASON)
    # Seasonal means use the analysis year only (an observed slice may start with the previous December).
    year = hourly_all.index.max().year
    this_year = hourly_all[hourly_all.index.year == year]
    season_means = {f"mean_t_{k}": float(this_year.loc[this_year.index.month.isin(m), "dry_bulb"].mean())
                    for k, m in MET_SEASONS.items()}
    days_this_year = daily[daily.index.year == year]
    jja_days = days_this_year[days_this_year.index.month.isin(MET_SEASONS["jja"])]
    metrics = {
        "wcdh": float(hourly["wcdh"].sum()),
        "twcdh": float(hourly["twcdh"].sum()),
        "swcdh": float(hourly["swcdh"].sum()),
        "peak_daily_wcdh": float(sd["wcdh_daily"].max()),
        "max_event_severity": max((e.severity for e in events), default=0.0),
        "max_event_duration": max((e.duration_days for e in events), default=0),
        "n_events": len(events),
        "hours_above_upper": int((dt_rounded >= 1).sum()),
        "max_delta_t": float(max(dt_rounded.max(), 0)),
        "max_daily_weighted_exceedance": float(daily_we.max()) if len(daily_we) else 0.0,
        "t_max": float(hourly["dry_bulb"].max()),
        "mean_daily_max": float(sd["t_max"].mean()),
        "mean_daily_min": float(sd["t_min"].mean()),
        "season_mean": float(hourly["dry_bulb"].mean()),
        "apr_sep_mean": float(hourly_all.loc[apr_sep, "dry_bulb"].mean()),
        "hot_days": int((sd["t_max"] >= config.hot_day_threshold).sum()),
        "warm_nights": int((sd["t_min"] >= config.warm_night_threshold).sum()),
        "cdh_22": float((hourly["dry_bulb"] - 22).clip(lower=0).sum()),
        "night_max_mean": float(sd["night_mean"].max()),
        "nights_above": int((sd["night_mean"] >= config.night_threshold).sum()),
        **season_means,
        "jja_mean_daily_max": float(jja_days["t_max"].mean()),
        "season_coverage": float(hourly["dry_bulb"].notna().mean()),
    }
    return SeasonAnalysis(
        name=series.name,
        config=config,
        static_threshold=s_thr,
        twcdh_offset=offset,
        daily=daily,
        hourly=hourly,
        events=events,
        metrics=metrics,
    )


def metrics_table(analyses: list[SeasonAnalysis]) -> pd.DataFrame:
    """One row per analysis, one column per metric."""
    return pd.DataFrame({a.name: a.metrics for a in analyses}).T

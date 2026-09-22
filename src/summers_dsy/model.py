"""Core data container for an hourly weather series."""

from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np
import pandas as pd

#: Nominal (non-leap) year used to place typical/design-year files on a calendar.
NOMINAL_YEAR = 2001


@dataclass
class WeatherSeries:
    """Hourly weather at one site.

    ``data`` is indexed by a tz-naive, hourly ``DatetimeIndex`` (hour-beginning,
    UTC/GMT) and must contain a ``dry_bulb`` column in degC. Other columns
    (e.g. ``rel_humidity``) are optional and carried along untouched.
    """

    name: str
    data: pd.DataFrame
    source: str = ""
    meta: dict = field(default_factory=dict)

    def __post_init__(self) -> None:
        if "dry_bulb" not in self.data.columns:
            raise ValueError("WeatherSeries data needs a 'dry_bulb' column")
        if not isinstance(self.data.index, pd.DatetimeIndex):
            raise TypeError("WeatherSeries data must have a DatetimeIndex")
        self.data = self.data.sort_index()
        self.data = self.data[~self.data.index.duplicated(keep="first")]

    @property
    def dry_bulb(self) -> pd.Series:
        return self.data["dry_bulb"]

    @property
    def is_typical_year(self) -> bool:
        """True for design/typical-year files (TRY/DSY/EPW), which wrap around."""
        return bool(self.meta.get("typical_year", False))

    @property
    def years(self) -> list[int]:
        return sorted(set(self.data.index.year))

    def year(self, year: int) -> "WeatherSeries":
        """The calendar-year slice of a multi-year observed record."""
        sub = self.data[self.data.index.year == year]
        if sub.empty:
            raise KeyError(f"No data for {year} in {self.name}")
        meta = {**self.meta, "year": year}
        return WeatherSeries(f"{self.name} {year}", sub, self.source, meta)

    def completeness(self) -> float:
        """Fraction of expected hours that have a valid dry-bulb value."""
        idx = self.data.index
        expected = pd.date_range(idx.min().floor("D"), idx.max().floor("D") + pd.Timedelta(hours=23), freq="h")
        return float(self.dry_bulb.reindex(expected).notna().mean())


def to_hourly(values: pd.Series, max_gap_hours: int = 6) -> pd.Series:
    """Regularise a timestamped temperature series to a complete hourly index.

    For sub-hourly or irregular observations each hour takes the report(s)
    nearest to it (e.g. an on-the-hour SYNOP in preference to a :50 METAR),
    averaging exact ties; gaps of up to ``max_gap_hours`` are linearly interpolated.
    """
    values = values.dropna()
    if values.empty:
        return values
    hour = values.index.round("h")
    offset = pd.Series(np.abs((values.index - hour).total_seconds()), index=values.index)
    nearest = offset.groupby(hour).transform("min")
    keep = (offset == nearest).to_numpy()
    rounded = values[keep].groupby(hour[keep]).mean()
    full = pd.date_range(rounded.index.min(), rounded.index.max(), freq="h")
    hourly = rounded.reindex(full)
    return hourly.interpolate(limit=max_gap_hours, limit_area="inside")


def nominal_index(months, days, hours_ending) -> tuple[pd.DatetimeIndex, np.ndarray]:
    """Hour-beginning timestamps in the nominal year from month/day/hour-ending columns.

    Returns the index and a boolean mask of the input rows kept (29 Feb is dropped).
    """
    months = np.asarray(months, dtype=int)
    days = np.asarray(days, dtype=int)
    hours = np.asarray(hours_ending, dtype=int) - 1
    # Drop 29 February if a typical-year file carries one.
    keep = ~((months == 2) & (days == 29))
    ts = pd.to_datetime(
        {"year": NOMINAL_YEAR, "month": months[keep], "day": days[keep], "hour": hours[keep]}
    )
    return pd.DatetimeIndex(ts), keep

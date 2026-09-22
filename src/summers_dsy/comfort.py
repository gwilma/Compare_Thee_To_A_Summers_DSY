"""Adaptive thermal comfort: running mean and comfort temperature (CIBSE TM52 / BS EN 15251).

Definitions (CIBSE TM52:2013 section 5, BS EN 15251:2007 Annex A / BS EN 16798-1):

* Exponentially weighted running mean of the daily mean outdoor temperature::

      Trm = (1 - alpha) * (Tod-1 + alpha*Tod-2 + alpha^2*Tod-3 + ...)

  evaluated recursively as ``Trm(n) = (1 - alpha) * Tod(n-1) + alpha * Trm(n-1)``,
  with alpha = 0.8 recommended.

* Where a recursion cannot be seeded, BS EN 15251 gives the 7-day approximation::

      Trm = (Tod-1 + 0.8 Tod-2 + 0.6 Tod-3 + 0.5 Tod-4 + 0.4 Tod-5 + 0.3 Tod-6 + 0.2 Tod-7) / 3.8

* Comfort temperature ``Tcomf = 0.33 Trm + 18.8``.
* Maximum acceptable temperature ``Tmax = Tcomf + offset`` where the offset is
  2 K (Category I), 3 K (Category II, the TM52/TM59 default) or 4 K (Category III).
"""

from __future__ import annotations

import numpy as np
import pandas as pd

SEVEN_DAY_WEIGHTS = np.array([1.0, 0.8, 0.6, 0.5, 0.4, 0.3, 0.2])

CATEGORY_OFFSETS = {"I": 2.0, "II": 3.0, "III": 4.0}

#: Days prepended from the end of a typical-year file to spin up the recursion.
#: alpha**30 ~ 0.001, so the seed has no practical influence after this long.
CYCLIC_SPINUP_DAYS = 30


def daily_mean(hourly: pd.Series, min_hours: int = 18) -> pd.Series:
    """Daily mean of hourly values; days with fewer than ``min_hours`` valid hours are NaN."""
    grouped = hourly.resample("D")
    means = grouped.mean()
    counts = grouped.count()
    return means.where(counts >= min_hours)


def seven_day_approximation(daily: pd.Series) -> pd.Series:
    """BS EN 15251 7-day approximation of the running mean, for every day."""
    values = daily.to_numpy(dtype=float)
    out = np.full(len(values), np.nan)
    for i in range(7, len(values)):
        window = values[i - 7 : i][::-1]  # Tod-1 ... Tod-7
        if not np.isnan(window).any():
            out[i] = float(window @ SEVEN_DAY_WEIGHTS / SEVEN_DAY_WEIGHTS.sum())
    return pd.Series(out, index=daily.index, name="trm")


def running_mean(
    daily: pd.Series,
    alpha: float = 0.8,
    method: str = "exponential",
    cyclic: bool = False,
    max_fill_days: int = 3,
) -> pd.Series:
    """Running mean outdoor temperature Trm for each day in ``daily``.

    ``method="exponential"`` (default) seeds with the 7-day approximation and
    then applies the exact recursion; ``method="seven_day"`` uses the
    approximation every day. With ``cyclic=True`` (typical/design years) the
    end of the year is wrapped round to spin up the start of the year.
    Short gaps in the daily means (``max_fill_days``) are interpolated; after a
    longer gap the recursion is reseeded once seven valid days are available.
    """
    if not 0 <= alpha < 1:
        raise ValueError("alpha must be in [0, 1)")
    daily = daily.astype(float).interpolate(limit=max_fill_days, limit_area="inside")

    if cyclic:
        n = min(CYCLIC_SPINUP_DAYS, len(daily))
        spin = daily.iloc[-n:]
        spin.index = daily.index[0] - pd.to_timedelta(np.arange(n, 0, -1), unit="D")
        extended = pd.concat([spin, daily])
        return running_mean(extended, alpha, method, cyclic=False, max_fill_days=max_fill_days).iloc[n:]

    if method == "seven_day":
        return seven_day_approximation(daily)
    if method != "exponential":
        raise ValueError(f"Unknown running-mean method {method!r}")

    values = daily.to_numpy()
    seeds = seven_day_approximation(daily).to_numpy()
    out = np.full(len(values), np.nan)
    prev = np.nan
    for i in range(len(values)):
        if not np.isnan(prev) and i > 0 and not np.isnan(values[i - 1]):
            prev = (1 - alpha) * values[i - 1] + alpha * prev
        else:
            prev = seeds[i]
        out[i] = prev
    return pd.Series(out, index=daily.index, name="trm")


def comfort_temperature(trm: pd.Series | float, clamp: bool = False):
    """TM52 comfort temperature ``0.33 Trm + 18.8``.

    BS EN 15251 states the upper limit applies for 10 < Trm < 30 degC; with
    ``clamp=True`` Trm is held within that range before evaluating.
    """
    if clamp:
        trm = np.clip(trm, 10.0, 30.0)
    return 0.33 * trm + 18.8


def daily_comfort(
    hourly: pd.Series,
    alpha: float = 0.8,
    method: str = "exponential",
    cyclic: bool = False,
    category: str = "II",
    clamp: bool = False,
) -> pd.DataFrame:
    """Daily table of mean/min/max temperature, Trm, Tcomf and Tmax."""
    grouped = hourly.resample("D")
    df = pd.DataFrame(
        {
            "t_mean": daily_mean(hourly),
            "t_min": grouped.min(),
            "t_max": grouped.max(),
        }
    )
    df["t_range"] = df["t_max"] - df["t_min"]
    df["trm"] = running_mean(df["t_mean"], alpha=alpha, method=method, cyclic=cyclic)
    df["t_comf"] = comfort_temperature(df["trm"], clamp=clamp)
    df["t_upper"] = df["t_comf"] + CATEGORY_OFFSETS[category]
    return df


def hourly_from_daily(daily_values: pd.Series, hourly_index: pd.DatetimeIndex) -> pd.Series:
    """Broadcast a daily series onto an hourly index."""
    return daily_values.reindex(hourly_index.floor("D")).set_axis(hourly_index)

"""Compare an observed summer with a set of reference (DSY/TRY) weather files."""

from __future__ import annotations

from dataclasses import dataclass, replace

import numpy as np
import pandas as pd

from .io.naming import reference_label
from .metrics import (
    DEFAULT_COMPARISON_METRICS,
    METRIC_INFO,
    AnalysisConfig,
    SeasonAnalysis,
    analyse,
    static_threshold,
    threshold_offset,
)
from .model import WeatherSeries

#: Metrics built from squared or summed exceedances: heavily skewed, so compared on a log scale.
LOG_METRICS = {"wcdh", "twcdh", "swcdh", "peak_daily_wcdh", "max_event_severity", "cdh_22", "max_daily_weighted_exceedance", "hours_above_upper"}


@dataclass
class Thresholds:
    static_threshold: float
    twcdh_offset: float
    provenance: str


def resolve_thresholds(config: AnalysisConfig, baseline: WeatherSeries | None, fallback: WeatherSeries) -> Thresholds:
    """Fix the regional SWCDH threshold and TWCDH offset once, for every series in a comparison.

    User-set values win. Otherwise both are derived (93rd centile) from the
    ``baseline`` series, normally the location's current-climate TRY, or from
    ``fallback`` (the observed record) when no baseline is available.
    """
    source = baseline if baseline is not None else fallback
    where = f"derived from {source.name}"
    s = config.static_threshold if config.static_threshold is not None else static_threshold(source, config)
    o = config.twcdh_offset if config.twcdh_offset is not None else threshold_offset(source, config)
    if config.static_threshold is not None and config.twcdh_offset is not None:
        where = "set manually"
    return Thresholds(s, o, where)


def with_thresholds(config: AnalysisConfig, t: Thresholds) -> AnalysisConfig:
    return replace(config, static_threshold=t.static_threshold, twcdh_offset=t.twcdh_offset)


def pick_baseline(entries: pd.DataFrame) -> int | None:
    """Library id of the current-climate TRY (else DSY1) for a location's entries."""
    if entries.empty:
        return None
    base = entries[entries["period"] == "Baseline"]
    for kind in ("TRY", "DSY1"):
        hit = base[base["kind"] == kind]
        if not hit.empty:
            return int(hit.iloc[0]["id"])
    return None


def _transform(values: pd.DataFrame | pd.Series, metric: str):
    return np.log1p(values.clip(lower=0)) if metric in LOG_METRICS else values


def similarity(
    target: dict,
    references: pd.DataFrame,
    metrics: list[str] | None = None,
    weights: dict[str, float] | None = None,
) -> pd.DataFrame:
    """Rank references by closeness to ``target`` across several metrics.

    Each metric is (log-)transformed where skewed, then scaled by its spread
    across the references plus the target, so no metric dominates through its
    units. Distance is the weighted RMS of the scaled differences and the score is
    ``100 * exp(-distance)``: 100 means identical, about 37 means one spread away on average.
    """
    metrics = metrics or DEFAULT_COMPARISON_METRICS
    weights = weights or {}
    w = np.array([weights.get(m, 1.0) for m in metrics], dtype=float)
    if w.sum() <= 0:
        raise ValueError("At least one metric needs a positive weight")

    ref = references[metrics].astype(float)
    tgt = pd.Series({m: float(target[m]) for m in metrics})
    z = pd.DataFrame(index=ref.index)
    for m in metrics:
        col = _transform(ref[m], m)
        t = float(_transform(pd.Series([tgt[m]]), m).iloc[0])
        spread = float(pd.concat([col, pd.Series([t])]).std(ddof=0))
        z[m] = (col - t) / (spread if spread > 1e-9 else 1.0)

    distance = np.sqrt((z.pow(2) * w).sum(axis=1) / w.sum())
    out = pd.DataFrame({"distance": distance, "score": 100 * np.exp(-distance)})
    out = out.join(z.add_prefix("z_"))
    return out.sort_values("distance")


def bracket(target_value: float, reference_values: pd.Series) -> tuple[str | None, str | None]:
    """Names of the references immediately below and above ``target_value``."""
    below = reference_values[reference_values <= target_value]
    above = reference_values[reference_values > target_value]
    return (below.idxmax() if not below.empty else None, above.idxmin() if not above.empty else None)


def return_periods(values: pd.Series) -> pd.DataFrame:
    """Rank and empirical return period (Weibull plotting position, (n+1)/rank) of annual values."""
    values = values.dropna()
    rank = values.rank(ascending=False, method="min")
    return pd.DataFrame({"value": values, "rank": rank.astype(int), "return_period": (len(values) + 1) / rank})


@dataclass
class Comparison:
    target: SeasonAnalysis
    references: dict[int, SeasonAnalysis]
    labels: dict[int, str]
    meta: dict[int, dict]
    ranking: pd.DataFrame
    thresholds: Thresholds
    metrics: list[str]

    @property
    def best_id(self) -> int:
        return int(self.ranking.index[0])

    def reference_metrics(self) -> pd.DataFrame:
        return pd.DataFrame({i: a.metrics for i, a in self.references.items()}).T

    def headline(self, place: str) -> str:
        best = self.meta[self.best_id]
        score = self.ranking.iloc[0]["score"]
        return f"{self.target.name} in {place} was most similar to {reference_label(best)} (similarity {score:.0f}/100)."


def compare(
    target: WeatherSeries,
    references: dict[int, WeatherSeries],
    config: AnalysisConfig,
    thresholds: Thresholds,
    metrics: list[str] | None = None,
    weights: dict[str, float] | None = None,
) -> Comparison:
    """Analyse the target and every reference with shared thresholds, then rank."""
    metrics = metrics or DEFAULT_COMPARISON_METRICS
    unknown = [m for m in metrics if m not in METRIC_INFO]
    if unknown:
        raise ValueError(f"Unknown metrics: {unknown}")
    cfg = with_thresholds(config, thresholds)
    tgt = analyse(target, cfg)
    refs = {i: analyse(s, cfg) for i, s in references.items()}
    table = pd.DataFrame({i: a.metrics for i, a in refs.items()}).T
    ranking = similarity(tgt.metrics, table, metrics, weights)
    labels = {i: reference_label(s.meta) for i, s in references.items()}
    ranking.insert(0, "label", ranking.index.map(labels))
    return Comparison(tgt, refs, labels, {i: s.meta for i, s in references.items()}, ranking, thresholds, metrics)


def years_exceeding(
    year_metrics: pd.DataFrame,
    references: pd.DataFrame,
    labels: dict,
    metric: str,
    location: str,
    latest_complete_year: int | None = None,
) -> pd.DataFrame:
    """How many observed years exceeded each reference file of one location on ``metric``.

    ``year_metrics`` has one row per analysed year (index = year); ``references`` one row per
    reference file of the same location (index = file id). A year exceeds a file when its metric
    is strictly greater. The statement reads e.g. "5 of the last 10 years (2016–2025) exceeded
    DSY1 · 2050s · High emissions · 50th percentile for Zone 7 on SWCDH"; "last" is used when the
    range runs to ``latest_complete_year`` (default: last year).
    """
    if metric not in METRIC_INFO:
        raise ValueError(f"Unknown metric {metric!r}")
    values = year_metrics[metric].dropna()
    if values.empty:
        raise ValueError("No analysed years to compare")
    years = sorted(int(y) for y in values.index)
    n, y0, y1 = len(years), years[0], years[-1]
    latest = latest_complete_year if latest_complete_year is not None else pd.Timestamp.today().year - 1
    span = f"{'the last ' if y1 >= latest else 'the '}{n} year{'s' if n != 1 else ''} ({y0}–{y1})" if n > 1 else f"{y0}"
    name, unit, _ = METRIC_INFO[metric]
    rows = []
    for ref_id, ref_value in references[metric].dropna().items():
        k = int((values > ref_value).sum())
        exceeded = [int(y) for y, v in values.items() if v > ref_value]
        lead = f"{k} of {span}" if n > 1 else (f"{y0}" if k else f"{y0} did not")
        verb = "exceeded" if (n > 1 or k) else "exceed"
        rows.append({
            "reference": ref_id,
            "label": labels.get(ref_id, str(ref_id)),
            "reference_value": float(ref_value),
            "years_exceeded": k,
            "years": n,
            "share": k / n,
            "exceeding_years": ", ".join(map(str, exceeded)),
            "statement": f"{lead} {verb} {labels.get(ref_id, ref_id)} for {location} on {name}",
        })
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    return out.sort_values(["reference_value"], ascending=False).reset_index(drop=True)

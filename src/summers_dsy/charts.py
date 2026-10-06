"""Plotly figures for the app.

Colour roles (validated for CVD separation and contrast in both modes):
slot 1 blue = daily minimum / the observed year, slot 2 orange = daily
maximum / the best-matching reference, slot 3 aqua = comfort temperature.
Ranges and context series are neutral grey so the data lines stay loud.
"""

from __future__ import annotations

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from .compare import LOG_METRICS, Comparison
from .metrics import METRIC_INFO, SeasonAnalysis
from .model import NOMINAL_YEAR

THEMES = {
    "light": {
        "surface": "#fcfcfb",
        "ink": "#0b0b0b",
        "ink2": "#52514e",
        "muted": "#898781",
        "grid": "#e1e0d9",
        "axis": "#c3c2b7",
        "series": ["#2a78d6", "#eb6834", "#1baf7a"],
        "context": "rgba(82,81,78,0.28)",
        "band": "rgba(82,81,78,0.10)",
    },
    "dark": {
        "surface": "#1a1a19",
        "ink": "#ffffff",
        "ink2": "#c3c2b7",
        "muted": "#898781",
        "grid": "#2c2c2a",
        "axis": "#383835",
        "series": ["#3987e5", "#d95926", "#199e70"],
        "context": "rgba(195,194,183,0.28)",
        "band": "rgba(195,194,183,0.10)",
    },
}

FONT = "system-ui, -apple-system, 'Segoe UI', sans-serif"


def short_label(label: str) -> str:
    """'DSY1 · 2050s · High emissions · 50th percentile' -> 'DSY1 · 2050s · High · 50th'."""
    return label.replace(" emissions", "").replace(" percentile", "")


def _log_ticks(values) -> list[float]:
    """1-2-5 ticks spanning the positive values, for uncluttered log axes."""
    v = np.asarray([x for x in values if x and x > 0], dtype=float)
    if v.size == 0:
        return []
    lo, hi = np.floor(np.log10(v.min())), np.ceil(np.log10(v.max()))
    ticks = [m * 10**e for e in range(int(lo), int(hi) + 1) for m in (1, 2, 5)]
    inside = [x for x in ticks if v.min() / 1.6 <= x <= v.max() * 1.6]
    return inside or ticks


def _theme(name: str) -> dict:
    return THEMES.get(name, THEMES["light"])


def _layout(fig: go.Figure, t: dict, height: int, title: str | None = None, legend: bool = True) -> go.Figure:
    fig.update_layout(
        height=height,
        margin=dict(l=56, r=120, t=56 if title else 24, b=40),
        paper_bgcolor=t["surface"],
        plot_bgcolor=t["surface"],
        font=dict(family=FONT, color=t["ink2"], size=13),
        title=dict(text=title, font=dict(color=t["ink"], size=16), x=0, xanchor="left", y=0.98) if title else None,
        hoverlabel=dict(bgcolor=t["surface"], bordercolor=t["axis"], font=dict(color=t["ink"], family=FONT)),
        showlegend=legend,
        legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="right", x=1, font=dict(color=t["ink2"]), bgcolor="rgba(0,0,0,0)"),
        hovermode="x unified",
    )
    fig.update_xaxes(showgrid=False, linecolor=t["axis"], ticks="outside", tickcolor=t["axis"], tickfont=dict(color=t["muted"]), zeroline=False)
    fig.update_yaxes(gridcolor=t["grid"], gridwidth=1, linecolor=t["axis"], tickfont=dict(color=t["muted"]), zeroline=False)
    return fig


def _end_label(fig, x, y, text, t, row=None, col=None, dy=0):
    kw = dict(row=row, col=col) if row else {}
    fig.add_annotation(
        x=x, y=y, text=text, showarrow=False, xanchor="left", xshift=6, yshift=dy,
        font=dict(color=t["ink2"], size=12), **kw,
    )


# ------------------------------------------------------------------ profile


def _profile_traces(d: pd.DataFrame, t: dict, show_legend: bool, upper_label: str) -> list[go.Scatter]:
    blue, orange, aqua = t["series"]
    x = d.index
    return [
        go.Scatter(x=x, y=d["t_min"], name="Daily min", line=dict(color=blue, width=2), legendgroup="min",
                   showlegend=show_legend, hovertemplate="%{y:.1f} °C"),
        go.Scatter(x=x, y=d["t_max"], name="Daily max", line=dict(color=orange, width=2), legendgroup="max",
                   fill="tonexty", fillcolor=t["band"], showlegend=show_legend, hovertemplate="%{y:.1f} °C"),
        go.Scatter(x=x, y=d["t_comf"], name="Tcomf", line=dict(color=aqua, width=2), legendgroup="comf",
                   showlegend=show_legend, hovertemplate="%{y:.1f} °C"),
        go.Scatter(x=x, y=d["t_upper"], name=upper_label, line=dict(color=t["muted"], width=1, dash="dot"),
                   legendgroup="upper", showlegend=show_legend, hovertemplate="%{y:.1f} °C"),
        go.Scatter(x=x[d["t_max"] > d["t_upper"]], y=d.loc[d["t_max"] > d["t_upper"], "t_max"], mode="markers",
                   name="Max above Tmax", legendgroup="exceed", showlegend=show_legend,
                   marker=dict(color=orange, size=8, line=dict(color=t["surface"], width=2)), hoverinfo="skip"),
    ]


def daily_profile(analysis: SeasonAnalysis, theme: str = "light", title: str | None = None,
                  whole_year: bool = False, y_range: tuple[float, float] | None = None) -> go.Figure:
    """Daily max/min with the range between them, comfort temperature and upper limit, plus daily range columns."""
    t = _theme(theme)
    d = analysis.daily if whole_year else analysis.season_daily()
    upper = f"Tmax (Cat {analysis.config.category})"
    fig = make_subplots(rows=2, cols=1, shared_xaxes=True, row_heights=[0.74, 0.26], vertical_spacing=0.05)
    for tr in _profile_traces(d, t, True, upper):
        fig.add_trace(tr, row=1, col=1)
    fig.add_trace(
        go.Bar(x=d.index, y=d["t_range"], name="Diurnal range", marker=dict(color=t["context"], cornerradius=2),
               showlegend=False, hovertemplate="range %{y:.1f} K"),
        row=2, col=1,
    )
    last = d.dropna(subset=["t_max"]).index[-1]
    for col, label, dy in (("t_max", "Daily max", 8), ("t_comf", "Tcomf", 0), ("t_min", "Daily min", -8)):
        _end_label(fig, last, d.loc[last, col], label, t, row=1, col=1, dy=dy)
    fig.update_yaxes(title_text="°C", row=1, col=1, range=y_range)
    fig.update_yaxes(title_text="Range K", row=2, col=1)
    fig.update_layout(bargap=0.35)
    fig.update_xaxes(tickformat="%-d %b", row=2, col=1)
    return _layout(fig, t, 520, title)


def profile_small_multiples(analyses: list[SeasonAnalysis], titles: list[str], theme: str = "light") -> go.Figure:
    """Side-by-side season profiles on a shared temperature scale, aligned by calendar date."""
    t = _theme(theme)
    fig = make_subplots(rows=1, cols=len(analyses), shared_yaxes=True, subplot_titles=[short_label(x) for x in titles],
                        horizontal_spacing=0.03)
    for i, a in enumerate(analyses, start=1):
        d = a.season_daily().copy()
        d.index = d.index.map(lambda ts: ts.replace(year=NOMINAL_YEAR))
        for tr in _profile_traces(d, t, i == 1, f"Tmax (Cat {a.config.category})"):
            fig.add_trace(tr, row=1, col=i)
    fig.update_xaxes(tickformat="%b")
    fig.update_yaxes(title_text="°C", col=1)
    for ann in fig.layout.annotations:
        ann.font = dict(color=t["ink"], size=13)
    fig = _layout(fig, t, 420)
    fig.update_layout(margin=dict(r=24, t=40, b=90), legend=dict(yanchor="top", y=-0.12, xanchor="left", x=0))
    return fig


# --------------------------------------------------------------- comparison


def _season_dates(a: SeasonAnalysis) -> pd.DataFrame:
    d = a.season_daily()[["wcdh_daily"]].fillna(0).copy()
    d.index = d.index.map(lambda ts: ts.replace(year=NOMINAL_YEAR))
    return d


def cumulative_wcdh(comparison: Comparison, theme: str = "light", n_context: int = 12) -> go.Figure:
    """Season accumulation of WCDH: the observed year against its closest references.

    Steep steps mark intense warm spells (DSY2-like); long steady climbs mark
    prolonged ones (DSY3-like).
    """
    t = _theme(theme)
    blue, orange, _ = t["series"]
    fig = go.Figure()
    ids = list(comparison.ranking.index[:n_context])
    best = comparison.best_id
    for i in reversed(ids[1:]):
        d = _season_dates(comparison.references[i])
        fig.add_trace(go.Scatter(x=d.index, y=d["wcdh_daily"].cumsum(), name=comparison.labels[i], showlegend=False,
                                 line=dict(color=t["context"], width=1.5), hovertemplate="%{y:,.0f} K²h"))
    d_best = _season_dates(comparison.references[best])
    d_obs = _season_dates(comparison.target)
    fig.add_trace(go.Scatter(x=d_best.index, y=d_best["wcdh_daily"].cumsum(), name=f"Best match: {comparison.labels[best]}",
                             line=dict(color=orange, width=2), hovertemplate="%{y:,.0f} K²h"))
    fig.add_trace(go.Scatter(x=d_obs.index, y=d_obs["wcdh_daily"].cumsum(), name=comparison.target.name,
                             line=dict(color=blue, width=2), hovertemplate="%{y:,.0f} K²h"))
    if len(ids) > 1:
        fig.add_trace(go.Scatter(x=[None], y=[None], name=f"Next {len(ids) - 1} closest", line=dict(color=t["context"], width=1.5)))
    for d, label in ((d_obs, comparison.target.name), (d_best, "Best match")):
        _end_label(fig, d.index[-1], d["wcdh_daily"].sum(), label, t)
    fig.update_yaxes(title_text="Cumulative WCDH (K²h)", rangemode="tozero")
    fig.update_xaxes(tickformat="%-d %b", range=[d_obs.index.min(), d_obs.index.max()])
    return _layout(fig, t, 420)


def similarity_bars(comparison: Comparison, theme: str = "light", top: int = 12) -> go.Figure:
    """Horizontal bars of similarity score, closest first; the best match is emphasised."""
    t = _theme(theme)
    blue, orange, _ = t["series"]
    r = comparison.ranking.head(top).iloc[::-1]
    colors = [orange if i == comparison.best_id else blue for i in r.index]
    fig = go.Figure(go.Bar(
        x=r["score"], y=[short_label(x) for x in r["label"]], orientation="h", marker=dict(color=colors, cornerradius=4),
        text=[f"{s:.0f}" for s in r["score"]], textposition="outside", textfont=dict(color=t["ink2"]),
        cliponaxis=False, hovertemplate="%{y}<br>similarity %{x:.1f}<extra></extra>",
    ))
    fig.update_xaxes(range=[0, 105], title_text="Similarity score", showgrid=True, gridcolor=t["grid"])
    fig.update_yaxes(showgrid=False, tickfont=dict(color=t["ink2"]))
    fig = _layout(fig, t, max(260, 34 * len(r) + 90), legend=False)
    fig.update_layout(hovermode="closest", bargap=0.35, margin=dict(l=24, r=40))
    return fig


def metric_strips(comparison: Comparison, theme: str = "light", metrics: list[str] | None = None) -> go.Figure:
    """One dot strip per metric: every reference as a grey dot, the best match and observed year highlighted."""
    t = _theme(theme)
    blue, orange, _ = t["series"]
    metrics = metrics or comparison.metrics
    table = comparison.reference_metrics()
    row_px, gap_px = 70, 84
    height = len(metrics) * (row_px + gap_px) + 70
    fig = make_subplots(rows=len(metrics), cols=1, vertical_spacing=gap_px / height,
                        subplot_titles=[f"{METRIC_INFO[m][0]} ({METRIC_INFO[m][1]})" for m in metrics])
    best = comparison.best_id
    for row, m in enumerate(metrics, start=1):
        others = table.drop(index=best)
        labels = [comparison.labels[i] for i in others.index]
        fig.add_trace(go.Scatter(x=others[m], y=[0] * len(others), mode="markers", name="Reference files",
                                 marker=dict(color=t["context"], size=10, line=dict(color=t["surface"], width=2)),
                                 text=labels, hovertemplate="%{text}<br>%{x:,.1f}<extra></extra>", showlegend=row == 1),
                      row=row, col=1)
        fig.add_trace(go.Scatter(x=[table.loc[best, m]], y=[0], mode="markers", name=f"Best match: {comparison.labels[best]}",
                                 marker=dict(color=orange, size=13, line=dict(color=t["surface"], width=2)),
                                 hovertemplate=f"{comparison.labels[best]}<br>%{{x:,.1f}}<extra></extra>", showlegend=row == 1),
                      row=row, col=1)
        fig.add_trace(go.Scatter(x=[comparison.target.metrics[m]], y=[0], mode="markers", name=comparison.target.name,
                                 marker=dict(color=blue, size=13, symbol="diamond", line=dict(color=t["surface"], width=2)),
                                 hovertemplate=f"{comparison.target.name}<br>%{{x:,.1f}}<extra></extra>", showlegend=row == 1),
                      row=row, col=1)
        if m in LOG_METRICS and (table[m] > 0).all() and comparison.target.metrics[m] > 0:
            ticks = _log_ticks([*table[m], comparison.target.metrics[m]])
            fig.update_xaxes(type="log", tickvals=ticks, ticktext=[f"{x:,.0f}" for x in ticks], minor=dict(showgrid=False),
                             row=row, col=1)
        fig.update_yaxes(visible=False, range=[-1, 1], row=row, col=1)
    for ann in fig.layout.annotations:
        ann.update(font=dict(color=t["ink"], size=13), x=0, xanchor="left", yshift=4)
    fig = _layout(fig, t, height)
    fig.update_layout(hovermode="closest", margin=dict(r=24, t=70))
    fig.update_xaxes(showgrid=True, gridcolor=t["grid"])
    return fig


def annual_bars(values: pd.Series, metric: str, theme: str = "light", selected: int | None = None,
                reference_lines: dict[str, float] | None = None) -> go.Figure:
    """One column per observed year, with reference-file values as labelled horizontal rules."""
    t = _theme(theme)
    blue, orange, _ = t["series"]
    colors = [orange if y == selected else blue for y in values.index]
    name, unit, _ = METRIC_INFO[metric]
    width = float(np.clip(26 * len(values) / 800, 0.12, 0.8))  # ~24px columns at typical widths
    fig = go.Figure(go.Bar(x=values.index.astype(str), y=values.values, width=width, marker=dict(color=colors, cornerradius=4),
                           hovertemplate=f"%{{x}}<br>{name} %{{y:,.1f}} {unit}<extra></extra>"))
    for label, v in (reference_lines or {}).items():
        fig.add_hline(y=v, line=dict(color=t["muted"], width=1, dash="dot"),
                      annotation=dict(text=label, font=dict(color=t["ink2"], size=11)),
                      annotation_position="right")
    fig.update_yaxes(title_text=f"{name} ({unit})", rangemode="tozero")
    fig = _layout(fig, t, 360, legend=False)
    fig.update_layout(hovermode="closest", margin=dict(r=200))
    return fig


def hourly_heatmap(analysis: SeasonAnalysis, theme: str = "light") -> go.Figure:
    """Hour-of-day by date: exceedance of outdoor air above Tcomf (K), a one-hue sequential scale."""
    t = _theme(theme)
    h = analysis.hourly
    excess = (h["dry_bulb"] - h["t_comf"]).clip(lower=0)
    grid = excess.groupby([excess.index.hour, excess.index.normalize()]).first().unstack()
    ramp = ["#cde2fb", "#86b6ef", "#3987e5", "#1c5cab", "#0d366b"] if theme != "dark" else ["#0d366b", "#184f95", "#256abf", "#5598e7", "#b7d3f6"]
    scale = [[0, t["surface"]], [0.001, ramp[0]], [0.25, ramp[1]], [0.5, ramp[2]], [0.75, ramp[3]], [1, ramp[4]]]
    fig = go.Figure(go.Heatmap(
        z=grid.values, x=grid.columns, y=grid.index, colorscale=scale, zmin=0,
        colorbar=dict(title=dict(text="K above Tcomf", font=dict(color=t["ink2"])), tickfont=dict(color=t["muted"]), outlinewidth=0),
        hovertemplate="%{x|%-d %b} %{y}:00<br>%{z:.1f} K above Tcomf<extra></extra>",
    ))
    fig.update_yaxes(title_text="Hour (UTC)", dtick=6, showgrid=False)
    fig.update_xaxes(tickformat="%-d %b")
    fig = _layout(fig, t, 320, legend=False)
    fig.update_layout(hovermode="closest", margin=dict(r=24))
    return fig


__all__ = [
    "THEMES",
    "annual_bars",
    "cumulative_wcdh",
    "daily_profile",
    "hourly_heatmap",
    "metric_strips",
    "profile_small_multiples",
    "similarity_bars",
]


def station_map(candidates: pd.DataFrame, selected_id: str, origin: tuple[float, float] | None = None,
                theme: str = "light") -> go.Figure:
    """Candidate stations (grey), the selected station (slot 1) and the searched point (slot 2)."""
    t = _theme(theme)
    blue, orange, _ = t["series"]
    sel = candidates[candidates["id"] == selected_id]
    others = candidates[candidates["id"] != selected_id]
    fig = go.Figure()
    fig.add_trace(go.Scattermap(lat=others["latitude"], lon=others["longitude"], mode="markers", name="Stations",
                                marker=dict(size=9, color="#898781"), text=others["label"], hoverinfo="text"))
    fig.add_trace(go.Scattermap(lat=sel["latitude"], lon=sel["longitude"], mode="markers", name="Selected station",
                                marker=dict(size=14, color=blue), text=sel["label"], hoverinfo="text"))
    if origin:
        fig.add_trace(go.Scattermap(lat=[origin[0]], lon=[origin[1]], mode="markers", name="Your site",
                                    marker=dict(size=12, color=orange), hoverinfo="name"))
    pts = pd.concat([candidates[["latitude", "longitude"]],
                     pd.DataFrame([origin], columns=["latitude", "longitude"]) if origin else None])
    lat_span = float(pts["latitude"].max() - pts["latitude"].min())
    lon_span = float(pts["longitude"].max() - pts["longitude"].min())
    span = max(lat_span, lon_span * 0.6, 0.05)
    zoom = float(np.clip(np.log2(360 / span) - 1.2, 3.5, 11))
    centre = dict(lat=float(pts["latitude"].mean()), lon=float(pts["longitude"].mean()))
    if not sel.empty and len(candidates) > 20:  # a long name-search list: centre on the chosen station
        centre, zoom = dict(lat=float(sel["latitude"].iloc[0]), lon=float(sel["longitude"].iloc[0])), 5.0
    fig.update_layout(map=dict(style="open-street-map", center=centre, zoom=zoom), height=260,
                      margin=dict(l=0, r=0, t=0, b=0), paper_bgcolor=t["surface"],
                      font=dict(family=FONT, color=t["ink2"], size=12),
                      legend=dict(orientation="h", yanchor="top", y=0.99, xanchor="left", x=0.01,
                                  bgcolor="rgba(255,255,255,0.8)" if theme != "dark" else "rgba(26,26,25,0.8)"))
    return fig

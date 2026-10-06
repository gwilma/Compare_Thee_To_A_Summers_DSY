"""Compare thee to a summer's DSY - overheating weather analysis.

Run with:  streamlit run app/streamlit_app.py
"""

from __future__ import annotations

import datetime as dt
from dataclasses import asdict

import pandas as pd
import streamlit as st

from summers_dsy import charts
from summers_dsy.compare import (
    Thresholds,
    bracket,
    years_exceeding,
    compare,
    pick_baseline,
    resolve_thresholds,
    return_periods,
    with_thresholds,
)
from summers_dsy.io import WeatherFileError, load_weather_file, parse_reference_name, reference_label
from summers_dsy.io.naming import CIBSE_LOCATIONS, EMISSIONS, KINDS, PERCENTILES, PERIODS
from summers_dsy.library import Library
from summers_dsy.metrics import (
    DEFAULT_COMPARISON_METRICS,
    METRIC_INFO,
    TM49_SEASON,
    TM59_SEASON,
    AnalysisConfig,
    analyse,
)
from summers_dsy.model import WeatherSeries
from summers_dsy.sources import (
    STATIONS,
    SourceError,
    data_dir,
    meteostat,
    combine_uploads,
    fetch_meteostat,
    fetch_midas,
    fetch_noaa_isd,
    fetch_open_meteo,
)
from summers_dsy.synthetic import demo_library

st.set_page_config(page_title="Compare thee to a summer's DSY", page_icon="☀️", layout="wide")

THEME = "dark" if getattr(getattr(st.context, "theme", None), "type", "light") == "dark" else "light"
TODAY = dt.date.today()
#: Latest year whose May–September season is complete.
LATEST_COMPLETE = TODAY.year if TODAY >= dt.date(TODAY.year, 10, 1) else TODAY.year - 1
PLOTLY = {"displaylogo": False, "modeBarButtonsToRemove": ["lasso2d", "select2d"]}

st.markdown(
    """
    <style>
      .headline {font-size: 1.45rem; line-height: 1.35; font-weight: 600; margin: .25rem 0 .75rem;}
      .subtle {color: #898781; font-size: .9rem;}
      div[data-testid="stMetricValue"] {font-variant-numeric: normal;}
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_resource
def get_library() -> Library:
    return Library()


LIB = get_library()


def fmt(metric: str, value: float) -> str:
    unit = METRIC_INFO[metric][1]
    if unit in ("K²h", "K·h"):
        return f"{value:,.0f} {unit}"
    if unit in ("days", "count", "h", "nights"):
        return f"{value:,.0f} {unit}"
    return f"{value:,.1f} {unit}"


# ------------------------------------------------------------------ sidebar

with st.sidebar:
    st.header("Analysis settings")
    season_choice = st.radio(
        "Season",
        ["TM59: 1 May – 30 Sep", "TM49: 1 Apr – 30 Sep", "Custom"],
        help="The period over which degree hours, events and statistics are summed.",
    )
    if season_choice.startswith("TM59"):
        season = TM59_SEASON
    elif season_choice.startswith("TM49"):
        season = TM49_SEASON
    else:
        c1, c2 = st.columns(2)
        s = c1.date_input("From", dt.date(2001, 5, 1), format="DD/MM")
        e = c2.date_input("To", dt.date(2001, 9, 30), format="DD/MM")
        season = ((s.month, s.day), (e.month, e.day))

    with st.expander("Running mean & comfort", expanded=False):
        alpha = st.slider("α (running-mean weight)", 0.5, 0.95, 0.8, 0.05, help="TM52 / BS EN 15251 recommend 0.8.")
        trm_method = st.selectbox(
            "Running mean method",
            ["exponential", "seven_day"],
            format_func=lambda m: {"exponential": "Exponential recursion (TM52)", "seven_day": "BS EN 15251 7-day approximation"}[m],
        )
        category = st.selectbox("Comfort category", ["I", "II", "III"], index=1, help="Tmax = Tcomf + 2 / 3 / 4 K. TM52/TM59 use Category II.")
        clamp = st.checkbox("Clamp Trm to 10–30 °C", help="The BS EN 15251 range of validity for the adaptive upper limit.")

    with st.expander("Degree-hour thresholds", expanded=False):
        st.caption(
            "SWCDH uses a fixed regional threshold, and TWCDH uses Tcomf plus a regional offset. Left on auto, both are "
            "the 93rd centile of the location's current-climate TRY (or DSY1) from the library, falling back to the observed record."
        )
        manual_static = st.checkbox("Set SWCDH threshold")
        static_val = st.number_input("SWCDH threshold (°C)", value=25.0, step=0.5, disabled=not manual_static)
        manual_offset = st.checkbox("Set TWCDH offset")
        offset_val = st.number_input("TWCDH offset (K above Tcomf)", value=0.0, step=0.5, disabled=not manual_offset)

    with st.expander("Day counts & events", expanded=False):
        hot_day = st.number_input("Hot day: max ≥ (°C)", value=28.0, step=0.5)
        warm_night = st.number_input("Warm night: min ≥ (°C)", value=16.0, step=0.5)
        event_min = st.number_input("Warm-event day: daily WCDH > (K²h)", value=0.0, step=5.0)

    with st.expander("Night-time (TM59:2026)", expanded=False):
        st.caption("TM59:2026 limits the mean bedroom temperature at night to below 27 °C, with no more than 4 exceedance "
                   "nights May–September. The same test applied to outdoor air shows how hard a season's nights are.")
        night_thr = st.number_input("Night mean threshold (°C)", value=27.0, step=0.5)
        c1, c2 = st.columns(2)
        night_start = c1.number_input("Night starts (hour)", 0, 23, 22)
        night_end = c2.number_input("Night ends (hour)", 0, 23, 7)

CONFIG = AnalysisConfig(
    season_start=season[0],
    season_end=season[1],
    alpha=alpha,
    trm_method=trm_method,
    category=category,
    clamp_trm=clamp,
    static_threshold=static_val if manual_static else None,
    twcdh_offset=offset_val if manual_offset else None,
    event_min_daily_wcdh=event_min,
    hot_day_threshold=hot_day,
    warm_night_threshold=warm_night,
    night_threshold=night_thr,
    night_start_hour=int(night_start),
    night_end_hour=int(night_end),
)
CONFIG_KEY = tuple(sorted(asdict(CONFIG).items()))


# ----------------------------------------------------------------- caching


@st.cache_data(show_spinner="Downloading weather data…", max_entries=16)
def cached_fetch(source: str, params: tuple) -> WeatherSeries:
    p = dict(params)
    if source == "open-meteo":
        return fetch_open_meteo(p["lat"], p["lon"], p["start"], p["end"], model=p["model"], name=p["name"])
    if source == "meteostat":
        return fetch_meteostat(p["station"], p["start"], p["end"], name=p["name"])
    if source == "noaa-isd":
        return fetch_noaa_isd(p["station"], p["start"], p["end"], name=p["name"])
    if source == "midas":
        return fetch_midas(p["county"], p["src_id"], p["site"], p["start"], p["end"], token=p["token"] or None,
                           version=p["version"], name=p["name"])
    raise ValueError(source)


@st.cache_data(show_spinner=False, max_entries=256)
def library_series(entry_id: int, _stamp: str) -> WeatherSeries:
    return LIB.load(entry_id)


def library_stamp() -> str:
    e = LIB.entries()
    return "" if e.empty else f"{len(e)}-{e['added_at'].max()}-{e['id'].sum()}"


def thresholds_for(location: str | None, observed: WeatherSeries | None) -> Thresholds | None:
    entries = LIB.entries(location) if location else pd.DataFrame()
    base_id = pick_baseline(entries) if location else None
    baseline = library_series(base_id, library_stamp()) if base_id is not None else None
    if baseline is None and observed is None:
        return None
    return resolve_thresholds(CONFIG, baseline, observed if observed is not None else baseline)


@st.cache_data(show_spinner="Analysing years…", max_entries=32)
def per_year_analyses(obs_key: str, _observed: WeatherSeries, config_key, thr: tuple[float, float]):
    cfg = with_thresholds(CONFIG, Thresholds(thr[0], thr[1], ""))
    out = {}
    for y in _observed.years:
        try:
            sub = _observed.year(y)
        except KeyError:
            continue
        a = analyse(sub, cfg)
        if a.metrics["season_coverage"] >= 0.8:
            out[y] = a
    return out


@st.cache_data(show_spinner=False)
def station_list(_stamp: str) -> pd.DataFrame:
    return meteostat.load_stations()


def station_stamp() -> str:
    p = data_dir() / "meteostat_stations.csv"
    return str(p.stat().st_mtime) if p.exists() else "bundled"


def describe_station(row: pd.Series) -> str:
    years = f"{int(row['start_year'])}–{int(row['end_year'])}" if pd.notna(row["start_year"]) else "years unknown"
    if pd.notna(row["end_year"]) and int(row["end_year"]) < LATEST_COMPLETE - 1:
        years += ", closed"
    dist = f" · {row['distance_km']:.0f} km" if "distance_km" in row and pd.notna(row.get("distance_km")) else ""
    return f"{row['name']} ({row['id']}) · {years}{dist}"


def meteostat_picker() -> None:
    """Station search, year range and multi-year download from Meteostat."""
    stations = station_list(station_stamp())
    how = st.radio("Find a station", ["Name", "Postcode", "Lat/long"], horizontal=True,
                   help="Search by name or code (WMO, ICAO), or list the stations nearest a UK postcode or a point.")
    candidates, origin = stations, None
    if how == "Name":
        q = st.text_input("Search", placeholder="e.g. Heathrow, EGLL or 03772")
        candidates = meteostat.search_stations(stations, q)
    elif how == "Postcode":
        pc = st.text_input("UK postcode", placeholder="e.g. SW1A 1AA")
        if pc.strip():
            try:
                lat, lon, where = meteostat.geocode_postcode(pc)
                origin = (lat, lon)
                st.caption(f"{where}: {lat:.4f}, {lon:.4f}")
            except (SourceError, ValueError) as exc:
                st.error(str(exc))
    else:
        c1, c2 = st.columns(2)
        origin = (c1.number_input("Latitude", value=51.5072, format="%.4f"),
                  c2.number_input("Longitude", value=-0.1276, format="%.4f"))
    if origin:
        candidates = meteostat.nearest_stations(stations, *origin, n=10)
    if candidates.empty:
        st.warning("No station matches. Try another name or a postcode.")
        return
    sid = st.selectbox("Station", candidates.index.tolist(), format_func=lambda i: describe_station(candidates.loc[i]))
    row = candidates.loc[sid]
    st.plotly_chart(charts.station_map(candidates, row["id"], origin, THEME), width="stretch", config=PLOTLY)

    first = int(row["start_year"]) if pd.notna(row["start_year"]) else 1973
    last = min(int(row["end_year"]) if pd.notna(row["end_year"]) else LATEST_COMPLETE, LATEST_COMPLETE)
    if last < first:
        st.warning("This station has no complete summer of hourly data.")
        return
    y0, y1 = st.slider("Years to download", first, last, (max(first, last - 9), last),
                       help="Whole calendar years. The December before the first year is downloaded too, to start the running mean.")
    include_model = st.checkbox("Fill gaps with Meteostat model data", value=False,
                                help="Meteostat fills missing hours with weather-model forecasts. Off: observations only "
                                     "(gaps of up to 6 hours are interpolated; years with under 80% of the season are skipped).")
    locs = LIB.locations()
    guess = meteostat.station_location_guess(str(row["name"]))
    options = ["(choose later)"] + locs
    loc = st.selectbox("Library location to compare with", options,
                       index=options.index(guess) if guess in options else 0,
                       help="Years are compared only with reference files of this location.")
    n = y1 - y0 + 1
    if st.button(f"Download {n} year{'s' if n > 1 else ''} from Meteostat", type="primary", width="stretch"):
        bar = st.progress(0.0, text="Starting download…")
        try:
            series = meteostat.fetch_years(
                str(row["id"]), y0, y1, include_model=include_model, name=str(row["name"]),
                progress=lambda k, total, year: bar.progress(k / total, text=f"Downloaded {year} ({k} of {total} files)"))
            bar.empty()
            st.session_state["observed"] = series
            st.session_state["observed_key"] = f"meteostat:{row['id']}:{y0}-{y1}:{include_model}"
            st.session_state["observed_location"] = None if loc == "(choose later)" else loc
            st.session_state.pop("selected_year", None)
        except (SourceError, ValueError) as exc:
            bar.empty()
            st.error(str(exc))
    with st.expander("Station list"):
        st.caption(f"{len(stations)} UK stations with hourly temperature observations. {meteostat.ATTRIBUTION}")
        if st.button("Refresh from Meteostat"):
            try:
                meteostat.refresh_stations()
                station_list.clear()
                st.rerun()
            except SourceError as exc:
                st.error(str(exc))


# ------------------------------------------------------------------ header

st.title("Compare thee to a summer's DSY")
st.markdown(
    "<span class='subtle'>Overheating metrics for real UK summers (running mean, TM52 comfort temperature, "
    "weighted cooling degree hours), compared against CIBSE Design Summer Years.</span>",
    unsafe_allow_html=True,
)

tab_obs, tab_lib, tab_cmp, tab_about = st.tabs(["① Observed weather", "② DSY library", "③ Compare", "About the metrics"])


# ------------------------------------------------------------ observed tab

with tab_obs:
    left, right = st.columns([1, 2], gap="large")
    with left:
        source = st.selectbox(
            "Data source",
            ["meteostat", "open-meteo", "noaa-isd", "midas", "upload"],
            format_func=lambda s: {
                "open-meteo": "Open-Meteo: ERA5 reanalysis",
                "meteostat": "Meteostat: station observations, many years",
                "noaa-isd": "NOAA ISD: station observations",
                "midas": "Met Office MIDAS Open (CEDA)",
                "upload": "Upload files (EPW, CSV, MIDAS)",
            }[s],
        )
        station_names = [s.name for s in STATIONS] + ["Custom…"]
        if source == "meteostat":
            meteostat_picker()
        elif source != "upload":
            st_name = st.selectbox("Station", station_names)
            station = next((s for s in STATIONS if s.name == st_name), None)
            c1, c2 = st.columns(2)
            start = c1.number_input("From year", 1940, TODAY.year, max(TODAY.year - 10, 1940))
            end = c2.number_input("To year", 1940, TODAY.year, TODAY.year - 1)
            params: dict = {"start": int(start), "end": int(end)}
            if source == "open-meteo":
                c1, c2 = st.columns(2)
                params["lat"] = c1.number_input("Latitude", value=station.latitude if station else 51.5, format="%.3f")
                params["lon"] = c2.number_input("Longitude", value=station.longitude if station else -0.12, format="%.3f")
                params["model"] = st.selectbox("Reanalysis", ["era5", "era5_land", "best_match"],
                                               format_func={"era5": "ERA5 (~25 km)", "era5_land": "ERA5-Land (~9 km)", "best_match": "Open-Meteo best match"}.get,
                                               help="ERA5 is ~25 km, ERA5-Land ~9 km. Both smooth urban heat islands and peaks.")
            elif source == "noaa-isd":
                params["station"] = st.text_input("ISD station (USAF+WBAN)", (station.usaf_wban or "") if station else "")
            elif source == "midas":
                county, src_id, site = station.midas if station and station.midas else ("", 0, "")
                params["county"] = st.text_input("CEDA county folder", county, help="e.g. greater-london")
                params["src_id"] = int(st.number_input("MIDAS src_id", value=src_id, step=1))
                params["site"] = st.text_input("Site folder name", site, help="e.g. heathrow")
                params["version"] = st.text_input("Dataset version", "202407")
                params["token"] = st.text_input("CEDA access token", type="password", help="Or set CEDA_TOKEN on the server.")
            params["name"] = station.name if station else "Custom site"
            if st.button("Fetch data", type="primary", width="stretch"):
                try:
                    series = cached_fetch(source, tuple(sorted(params.items())))
                    st.session_state["observed"] = series
                    st.session_state["observed_key"] = f"{source}:{sorted((k, v) for k, v in params.items() if k != 'token')}"
                    st.session_state["observed_location"] = station.cibse_location if station else None
                except (SourceError, ValueError) as exc:
                    st.error(str(exc))
        else:
            files = st.file_uploader("Observed weather files", type=["epw", "csv", "txt"], accept_multiple_files=True,
                                     help="One or more EPW, CSV (datetime + temperature) or MIDAS Open BADC-CSV files.")
            site_name = st.text_input("Site name", "Uploaded site")
            loc_guess = st.selectbox("Nearest CIBSE location", ["(none)"] + sorted(set(CIBSE_LOCATIONS) | set(LIB.locations())))
            if files and st.button("Load files", type="primary", width="stretch"):
                try:
                    parsed = [load_weather_file(f.getvalue(), f.name) for f in files]
                    st.session_state["observed"] = combine_uploads(parsed, site_name)
                    st.session_state["observed_key"] = "upload:" + ",".join(f"{f.name}:{f.size}" for f in files)
                    st.session_state["observed_location"] = None if loc_guess == "(none)" else loc_guess
                except (WeatherFileError, SourceError, ValueError) as exc:
                    st.error(f"Could not read the files: {exc}")

    observed: WeatherSeries | None = st.session_state.get("observed")
    analyses, a = {}, None
    with right:
        if observed is None:
            st.info("Choose a source and station, then fetch data. Uploaded files and fetched data stay on this machine.")
        else:
            loc = st.session_state.get("observed_location")
            thr = thresholds_for(loc if loc in LIB.locations() else None, observed)
            analyses = per_year_analyses(st.session_state["observed_key"], observed, CONFIG_KEY, (thr.static_threshold, thr.twcdh_offset))
            st.session_state["observed_analyses"] = analyses
            st.subheader(observed.name)
            if observed.source == "meteostat":
                miss = observed.meta.get("missing_years") or []
                share = observed.meta.get("model_share") or {}
                model_note = (f"model-filled hours {'included' if observed.meta.get('include_model') else 'dropped'} "
                              f"(up to {100 * max(share.values()):.1f}% of a year)" if share else "")
                st.caption(" · ".join(x for x in (
                    f"Meteostat station {observed.meta.get('station')}",
                    f"no file for {', '.join(map(str, miss))}" if miss else "", model_note,
                    observed.meta.get("attribution", "")) if x))
            st.caption(
                f"{observed.source} · {observed.years[0]}–{observed.years[-1]} · {observed.completeness():.1%} hourly coverage · "
                f"SWCDH threshold {thr.static_threshold:.1f} °C, TWCDH offset {thr.twcdh_offset:+.1f} K ({thr.provenance})"
            )
            if not analyses:
                st.warning("No year has at least 80% of the season's hours, so there's nothing to analyse.")
            else:
                years = sorted(analyses)
                table = pd.DataFrame({y: a.metrics for y, a in analyses.items()}).T
                sel = st.select_slider("Year", options=years, value=years[-1])
                st.session_state["selected_year"] = sel
                a = analyses[sel]
                m = a.metrics
                cols = st.columns(5)
                rp = return_periods(table["wcdh"])
                cols[0].metric("WCDH", fmt("wcdh", m["wcdh"]),
                               help=f"Rank {rp.loc[sel, 'rank']} of {len(rp)} years (≈ 1-in-{rp.loc[sel, 'return_period']:.1f})")
                cols[1].metric("Peak temperature", fmt("t_max", m["t_max"]))
                cols[2].metric("Longest warm event", fmt("max_event_duration", m["max_event_duration"]))
                cols[3].metric(f"Hours > Tmax (Cat {CONFIG.category})", fmt("hours_above_upper", m["hours_above_upper"]))
                cols[4].metric(f"Nights ≥ {CONFIG.night_threshold:g} °C", fmt("nights_above", m["nights_above"]),
                               help=f"Nights ({CONFIG.night_start_hour:02d}:00–{CONFIG.night_end_hour:02d}:00) with a mean outdoor "
                                    f"temperature at or above the threshold. Warmest night: {m['night_max_mean']:.1f} °C. "
                                    "TM59:2026 allows bedrooms no more than 4 nights with a mean of 27 °C or more.")
    if a is not None:
        st.divider()
        st.markdown(f"##### {sel}: daily maximum and minimum, diurnal range and comfort temperature")
        st.plotly_chart(charts.daily_profile(a, THEME), width="stretch", config=PLOTLY)
        st.markdown("##### Hours above the comfort temperature")
        st.plotly_chart(charts.hourly_heatmap(a, THEME), width="stretch", config=PLOTLY)

        st.markdown("#### Year by year")
        metric = st.selectbox("Metric", list(METRIC_INFO), format_func=lambda k: f"{METRIC_INFO[k][0]} ({METRIC_INFO[k][1]})")
        ref_lines = {}
        if loc and loc in LIB.locations():
            ents = LIB.entries(loc)
            base = ents[ents["period"] == "Baseline"]
            cfg = with_thresholds(CONFIG, thr)
            for _, row in base.iterrows():
                ref_lines[row["kind"]] = analyse(library_series(int(row["id"]), library_stamp()), cfg).metrics[metric]
        st.plotly_chart(charts.annual_bars(table[metric], metric, THEME, selected=sel, reference_lines=ref_lines),
                        width="stretch", config=PLOTLY)
        if ref_lines:
            st.caption(f"Dotted lines: current-climate CIBSE files for {loc}.")
        with st.expander("Table: metrics for every year"):
            shown = table.drop(columns=["season_coverage"]).rename(columns={k: METRIC_INFO[k][0] for k in METRIC_INFO})
            st.dataframe(shown.style.format(precision=1), width="stretch")
        with st.expander(f"Warm events in {sel}"):
            ev = a.events_frame()
            if ev.empty:
                st.write("No warm events in the season.")
            else:
                ev["start"] = ev["start"].dt.strftime("%d %b")
                ev["end"] = ev["end"].dt.strftime("%d %b")
                st.dataframe(ev.sort_values("severity", ascending=False).round(1), width="stretch", hide_index=True)
        st.download_button(f"Download {sel} daily table (CSV)", a.daily.round(3).to_csv().encode(),
                           file_name=f"{observed.name}_{sel}_daily.csv".replace(" ", "_"))


# ------------------------------------------------------------- library tab

with tab_lib:
    st.markdown(
        "Upload your licensed **CIBSE TRY/DSY files** (EPW or CSV) once. The library keeps each file's metadata and hourly dry-bulb "
        "in a local database (`data/library.sqlite`, git-ignored), so metrics are recomputed whenever settings change."
    )
    up = st.file_uploader("CIBSE weather files", type=["epw", "csv"], accept_multiple_files=True, key="lib_upload")
    if up:
        rows, parsed = [], {}
        for f in up:
            try:
                s = load_weather_file(f.getvalue(), f.name)
            except (WeatherFileError, ValueError) as exc:
                st.error(f"{f.name}: {exc}")
                continue
            parsed[f.name] = s
            meta = parse_reference_name(f.name)
            rows.append({"file": f.name, "location": meta["location"], "kind": meta["kind"], "period": meta["period"],
                         "emissions": meta["emissions"], "percentile": meta["percentile"],
                         "hours": len(s.data), "source year": s.meta.get("source_year")})
        if rows:
            st.caption("Check the metadata read from the file names and correct anything wrong before adding.")
            edited = st.data_editor(
                pd.DataFrame(rows),
                hide_index=True,
                width="stretch",
                disabled=["file", "hours", "source year"],
                column_config={
                    "location": st.column_config.SelectboxColumn(options=list(CIBSE_LOCATIONS) + ["Other"], required=True),
                    "kind": st.column_config.SelectboxColumn(options=KINDS, required=True),
                    "period": st.column_config.SelectboxColumn(options=PERIODS, required=True),
                    "emissions": st.column_config.SelectboxColumn(options=EMISSIONS),
                    "percentile": st.column_config.SelectboxColumn(options=PERCENTILES),
                },
                key="lib_editor",
            )
            if st.button(f"Add {len(edited)} file(s) to library", type="primary"):
                added = 0
                for _, r in edited.iterrows():
                    try:
                        LIB.add(parsed[r["file"]], r["file"], r.to_dict())
                        added += 1
                    except ValueError as exc:
                        st.error(f"{r['file']}: {exc}")
                st.success(f"Added {added} file(s).")

    entries = LIB.entries()
    st.markdown(f"#### Library: {len(entries)} file(s)")
    if entries.empty:
        st.info("The library is empty. Upload CIBSE files above, or load the synthetic demo set to try the app.")
    else:
        view = entries[["id", "location", "kind", "period", "emissions", "percentile", "filename", "source_year"]].copy()
        view.insert(0, "delete", False)
        ed = st.data_editor(
            view, hide_index=True, width="stretch", disabled=["id", "filename", "source_year"],
            column_config={
                "location": st.column_config.SelectboxColumn(options=sorted(set(list(CIBSE_LOCATIONS) + ["Other"] + view["location"].tolist()))),
                "kind": st.column_config.SelectboxColumn(options=KINDS),
                "period": st.column_config.SelectboxColumn(options=PERIODS),
                "emissions": st.column_config.SelectboxColumn(options=EMISSIONS),
                "percentile": st.column_config.SelectboxColumn(options=PERCENTILES),
            },
            key="lib_table",
        )
        if st.button("Save changes"):
            for _, r in ed.iterrows():
                if r["delete"]:
                    LIB.delete([int(r["id"])])
                else:
                    LIB.update(int(r["id"]), **{k: r[k] for k in ("location", "kind", "period", "emissions", "percentile")})
            st.rerun()
    if st.button("Load synthetic demo set (not CIBSE data)"):
        for fname, (s, meta) in demo_library().items():
            LIB.add(s, fname, meta)
        st.rerun()


# ------------------------------------------------------------- compare tab

with tab_cmp:
    observed = st.session_state.get("observed")
    analyses = st.session_state.get("observed_analyses") or {}
    locations = LIB.locations()
    if observed is None or not analyses:
        st.info("First load observed weather in tab ①.")
    elif not locations:
        st.info("Add reference weather files to the library in tab ②.")
    else:
        c1, c2, c3 = st.columns([1, 2, 2])
        years = sorted(analyses)
        year = c1.selectbox("Observed year", years, index=years.index(st.session_state.get("selected_year", years[-1])))
        default_loc = st.session_state.get("observed_location")
        location = c2.selectbox("Compare with library location", locations,
                                index=locations.index(default_loc) if default_loc in locations else 0)
        ents = LIB.entries(location)
        with c3:
            kinds = st.multiselect("File types", sorted(ents["kind"].unique()), default=sorted(ents["kind"].unique()))
        c4, c5, c6 = st.columns(3)
        periods = c4.multiselect("Periods", [p for p in PERIODS if p in set(ents["period"])], default=[p for p in PERIODS if p in set(ents["period"])])
        em_opts = sorted(e for e in ents["emissions"].dropna().unique())
        emissions = c5.multiselect("Emissions", em_opts, default=em_opts)
        pc_opts = sorted(int(p) for p in ents["percentile"].dropna().unique())
        pcts = c6.multiselect("Percentiles", pc_opts, default=pc_opts)

        with st.expander("Metrics and weights used for matching"):
            metrics = st.multiselect("Metrics", list(METRIC_INFO), default=DEFAULT_COMPARISON_METRICS,
                                     format_func=lambda k: METRIC_INFO[k][0])
            weights = {}
            wcols = st.columns(4)
            for i, mkey in enumerate(metrics):
                weights[mkey] = wcols[i % 4].slider(METRIC_INFO[mkey][0], 0.0, 3.0, 1.0, 0.25, key=f"w_{mkey}")

        chosen = ents[
            ents["kind"].isin(kinds)
            & ents["period"].isin(periods)
            & (ents["emissions"].isin(emissions) | ents["emissions"].isna())
            & (ents["percentile"].isin(pcts) | ents["percentile"].isna())
        ]
        if chosen.empty or not metrics:
            st.warning("Nothing to compare with: widen the filters or pick at least one metric.")
        else:
            thr = thresholds_for(location, observed)
            target = observed.year(year, name=str(year))
            refs = {int(i): library_series(int(i), library_stamp()) for i in chosen["id"]}
            try:
                result = compare(target, refs, CONFIG, thr, metrics, weights)
            except ValueError as exc:
                st.error(str(exc))
                st.stop()

            best = result.meta[result.best_id]
            st.markdown(
                f"<div class='headline'>{year} at {observed.name} was most similar to "
                f"<u>{reference_label(best)}</u> for {location}.</div>",
                unsafe_allow_html=True,
            )
            below, above = bracket(result.target.metrics["wcdh"], result.reference_metrics()["wcdh"])
            parts = [f"Its WCDH of {fmt('wcdh', result.target.metrics['wcdh'])}"]
            if below is not None and above is not None:
                parts.append(f"lies between {result.labels[below]} and {result.labels[above]}")
            elif above is not None:
                parts.append(f"is below every selected file (lowest: {result.labels[above]})")
            elif below is not None:
                parts.append(f"exceeds every selected file (highest: {result.labels[below]})")
            st.markdown(" ".join(parts) + f". Thresholds: SWCDH {thr.static_threshold:.1f} °C, TWCDH Tcomf {thr.twcdh_offset:+.1f} K ({thr.provenance}).")

            top = result.ranking.head(3)
            cols = st.columns(3)
            for col, (rid, row) in zip(cols, top.iterrows()):
                col.metric(row["label"], f"{row['score']:.0f} / 100", help="Similarity: 100 means identical on every chosen metric")

            st.markdown(f"##### How many years exceeded each {location} reference file")
            yr_an = per_year_analyses(st.session_state["observed_key"], observed, CONFIG_KEY,
                                      (thr.static_threshold, thr.twcdh_offset))
            year_table = pd.DataFrame({y: a.metrics for y, a in yr_an.items()}).T
            mc1, mc2 = st.columns([1, 2])
            mkey = mc1.selectbox("Metric", list(METRIC_INFO), index=list(METRIC_INFO).index("swcdh"),
                                 format_func=lambda k: f"{METRIC_INFO[k][0]} ({METRIC_INFO[k][1]})", key="years_metric",
                                 help="CIBSE ranks DSY1 by SWCDH. A year exceeds a file when its value is higher.")
            ref_table = result.reference_metrics()
            ex = years_exceeding(year_table, ref_table, result.labels, mkey, location, latest_complete_year=LATEST_COMPLETE)
            mc2.caption(f"{len(year_table)} analysed year(s) against the {len(ref_table)} selected {location} files, all with the "
                        f"same thresholds. Only files for this location are used.")
            key_rows = ex[[result.meta[i].get("kind") == "DSY1" for i in ex["reference"]]]
            if not key_rows.empty:
                st.markdown("\n".join(f"- {s_}" for s_ in key_rows["statement"]))
            st.dataframe(
                ex[["statement", "years_exceeded", "years", "exceeding_years", "reference_value"]].rename(columns={
                    "statement": "Statement", "years_exceeded": "Years exceeded", "years": "Years",
                    "exceeding_years": "Which years", "reference_value": f"File {METRIC_INFO[mkey][0]}"}),
                hide_index=True, width="stretch",
                column_config={f"File {METRIC_INFO[mkey][0]}": st.column_config.NumberColumn(format="%.1f"),
                               "Statement": st.column_config.TextColumn(width="large")})
            default_refs = [i for i in ex["reference"] if result.meta[i].get("kind") == "DSY1"
                            and result.meta[i].get("percentile") in (50, None)][:4] or list(ex["reference"][:3])
            shown = st.multiselect("Reference lines on the chart", list(ex["reference"]), default=default_refs,
                                   format_func=lambda i: result.labels[i])
            st.plotly_chart(charts.annual_bars(year_table[mkey], mkey, THEME, selected=year,
                                               reference_lines={charts.short_label(result.labels[i]): ref_table.loc[i, mkey]
                                                                for i in shown}),
                            width="stretch", config=PLOTLY)
            st.download_button("Download the statements (CSV)", ex.to_csv(index=False).encode(),
                               file_name=f"years_vs_{location}_{mkey}.csv".replace(" ", "_"))

            g1, g2 = st.columns([1, 1], gap="large")
            with g1:
                st.markdown("##### Closest reference files")
                st.plotly_chart(charts.similarity_bars(result, THEME), width="stretch", config=PLOTLY)
            with g2:
                st.markdown("##### How the season's heat accumulated")
                st.plotly_chart(charts.cumulative_wcdh(result, THEME), width="stretch", config=PLOTLY)
                st.caption("Steep steps are short, intense spells (DSY2-like). Long steady climbs are prolonged spells (DSY3-like).")

            st.markdown("##### Temperature profiles: observed and best matches")
            second = int(result.ranking.index[1]) if len(result.ranking) > 1 else None
            panels = [result.target, result.references[result.best_id]] + ([result.references[second]] if second is not None else [])
            titles = [str(year), result.labels[result.best_id]] + ([result.labels[second]] if second is not None else [])
            st.plotly_chart(charts.profile_small_multiples(panels, titles, THEME), width="stretch", config=PLOTLY)

            st.markdown("##### Where the observed year sits on each metric")
            st.plotly_chart(charts.metric_strips(result, THEME), width="stretch", config=PLOTLY)

            with st.expander("Table: metrics for the observed year and every reference file"):
                tbl = result.reference_metrics()
                tbl.index = [result.labels[i] for i in tbl.index]
                tbl.loc[f"Observed {year}"] = pd.Series(result.target.metrics)
                tbl = tbl.loc[[f"Observed {year}"] + [result.labels[i] for i in result.ranking.index]]
                tbl.insert(0, "similarity", [None] + result.ranking["score"].round(1).tolist())
                shown = tbl.drop(columns=["season_coverage"]).rename(columns={k: METRIC_INFO[k][0] for k in METRIC_INFO})
                st.dataframe(shown.style.format(precision=1, na_rep="–"), width="stretch")
                st.download_button("Download comparison (CSV)", shown.to_csv().encode(), file_name=f"comparison_{year}_{location}.csv")


# --------------------------------------------------------------- about tab

with tab_about:
    st.markdown(
        r"""
### Comfort temperature (CIBSE TM52 / BS EN 15251)
* **Running mean of the daily mean outdoor temperature**:
  $T_{rm} = (1-\alpha)\,(T_{od-1} + \alpha T_{od-2} + \alpha^2 T_{od-3} + \dots)$, evaluated recursively as
  $T_{rm,n} = (1-\alpha)T_{od,n-1} + \alpha T_{rm,n-1}$ with $\alpha = 0.8$. The recursion is seeded with the
  BS EN 15251 approximation $(T_{od-1} + 0.8T_{od-2} + 0.6T_{od-3} + 0.5T_{od-4} + 0.4T_{od-5} + 0.3T_{od-6} + 0.2T_{od-7})/3.8$.
  Design-year files wrap round from December to January; observed years use the preceding December.
* **Comfort temperature**: $T_{comf} = 0.33\,T_{rm} + 18.8$. **Upper limit**: $T_{max} = T_{comf} + 3$ K (Category II; I = +2, III = +4).

### Weighted cooling degree hours (Eames 2016; Liu *et al.* 2016)
Summed over the hours of the season (TM59: 1 May to 30 September):
* **WCDH** $= \sum \max(0,\,T - T_{comf})^2$, the adaptive-comfort form of the metric (Eames 2016).
* **TWCDH** $= \sum \max(0,\,T - (T_{comf} + \Delta T_{region}))^2$, where the comfort temperature is adjusted by a regional offset.
* **SWCDH** $= \sum \max(0,\,T - T_{static})^2$, where $T_{static}$ is a regional 93rd-centile temperature. CIBSE selects DSY1 as the year with a 1-in-7 return period ranked by SWCDH; DSY2 is the year with the most intense heat event and DSY3 the year with the longest.

When not set manually, the regional values are derived from the location's current-climate TRY (or DSY1) in the library.
$T_{static}$ is the 93rd centile of in-season hourly temperature, and $\Delta T_{region}$ is the 93rd centile of $T - T_{comf}$.
If the library has no file for that location, the observed record is used instead. The same values are applied to every
file in a comparison.

### Warm events: what separates DSY1, DSY2 and DSY3
An event is a run of consecutive days with daily WCDH above the event threshold (default 0).
**Severity** is the event's total WCDH, **intensity** its peak daily WCDH, and **duration** its length in days.
DSY1 is a moderately warm summer, DSY2 a year with a short, intense warm spell, and DSY3 a year with a long, sustained one.

### TM59:2026 night-time indicator on outdoor air
TM59:2026 requires the mean bedroom temperature at night to stay below 27 °C, with no more than four exceedance nights
between May and September. Here the same test is applied to outdoor air. Each night runs from 22:00 to 07:00 (file time) and
is labelled by the evening it starts. A night needs at least 7 valid hours to count. **Warmest night** is the highest nightly
mean in the season, and **Nights ≥ threshold** counts nights with a mean at or above the threshold (27 °C by default; set it in
the sidebar). Outdoor nights in the UK rarely reach 27 °C, so a lower threshold is often more telling. Neither is a building
compliance check.

### Seasonal climate (energy demand and overheating)
**Seasonal mean air temperature** for each meteorological season, winter (December–February), spring, summer and
autumn, is a guide to heating and cooling energy demand. **Summer mean daily maximum** (June–August) is a guide to
overheating risk. Both use the analysis year only: winter takes that year's January, February and December.

### TM52-style indicators on outdoor air
Hours with rounded $\Delta T = T - T_{max} \ge 1$ K, the largest daily weighted exceedance $\sum h_e \Delta T$, and the maximum $\Delta T$.
These are computed on external air as indicators of climate severity. They are *not* building compliance checks.

### Years against reference files
For a location's reference files, the app counts how many of the analysed years had a higher value of the chosen
metric (default SWCDH) than each file, e.g. "5 of the last 10 years (2016–2025) exceeded DSY1 · 2050s · High
emissions · 50th percentile". It says "the last" when the range runs to the latest complete summer. Only files of
the chosen location are used, and every year and file is analysed with that location's thresholds.

### Similarity
Each metric is put on a log scale where it is a squared or summed exceedance, then divided by its spread across the files
being compared. The distance is the weighted RMS difference, and the score is $100\,e^{-d}$.

### Data sources and caveats
* **Open-Meteo ERA5 / ERA5-Land**: gridded reanalysis (~25 km / ~9 km). It under-represents peaks and urban heat islands.
* **Meteostat**: one file per station and year from data.meteostat.net (UTC). Meteostat fills gaps with model forecasts;
  those values are dropped unless you tick *Fill gaps with Meteostat model data*. Gaps of up to 6 hours are interpolated,
  and years with under 80% of the May–September hours are not counted. Station list © Meteostat, CC BY 4.0. Check
  meteostat.net's terms for the hourly data before commercial use.
* **NOAA ISD**: station reports. Where several reports fall near an hour, the one nearest the hour is used.
* **Met Office MIDAS Open**: the authoritative UK record, which needs a CEDA account. You can upload its yearly BADC-CSV files.
* All times are UTC/GMT, matching CIBSE files.
        """
    )

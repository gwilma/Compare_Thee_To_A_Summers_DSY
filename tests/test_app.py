"""Headless smoke test of the Streamlit app (no network: observed data is injected)."""

import pandas as pd
import pytest

from summers_dsy.library import Library
from summers_dsy.model import WeatherSeries
from summers_dsy.synthetic import WarmSpell, demo_library, synthetic_year

AppTest = pytest.importorskip("streamlit.testing.v1").AppTest


def observed_record() -> WeatherSeries:
    parts = []
    for i, year in enumerate(range(2016, 2021)):
        spells = [WarmSpell(200 + 3 * i, 4 + 4 * i, 5 + i)]
        s = synthetic_year(spells=spells, warming=0.5, seed=40 + i, year=year)
        parts.append(s)
    temps = pd.concat(parts)
    temps = temps.reindex(pd.date_range(temps.index[0], temps.index[-1], freq="h")).interpolate()
    return WeatherSeries("Test site", temps.to_frame("dry_bulb"), "test", {"typical_year": False})


def test_app_runs_end_to_end(tmp_path, monkeypatch):
    monkeypatch.setenv("SUMMERS_DSY_DATA", str(tmp_path))
    lib = Library(tmp_path / "library.sqlite")
    for fname, (s, meta) in demo_library().items():
        lib.add(s, fname, meta)

    at = AppTest.from_file("../app/streamlit_app.py", default_timeout=60)
    at.session_state["observed"] = observed_record()
    at.session_state["observed_key"] = "test"
    at.session_state["observed_location"] = "Demo (synthetic)"
    at.run()
    assert not at.exception, at.exception
    headline = [m.value for m in at.markdown if "most similar to" in m.value]
    assert headline, "comparison headline missing"
    assert "Demo (synthetic)" in headline[0]


def test_app_meteostat_multi_year(tmp_path, monkeypatch):
    """Find a station by name, download several years (simulated Meteostat files), compare years with the library."""
    import gzip

    from summers_dsy.sources import meteostat as ms

    monkeypatch.setenv("SUMMERS_DSY_DATA", str(tmp_path))
    lib = Library(tmp_path / "library.sqlite")
    for fname, (s, meta) in demo_library().items():
        lib.add(s, fname, meta)

    def year_file(year):
        temps = synthetic_year(spells=[WarmSpell(200, 4 + (year % 5), 4 + (year % 4))], warming=0.3 * (year - 2018),
                               seed=year, year=year)
        rows = ["year,month,day,hour,temp,temp_source"]
        rows += [f"{t.year},{t.month},{t.day},{t.hour},{v:.1f},isd_lite" for t, v in temps.items()]
        return gzip.compress("\n".join(rows).encode())

    monkeypatch.setattr(ms, "get_bytes", lambda url, **kw: year_file(int(url.split("/")[-2])) if "data.meteostat" in url else None)

    at = AppTest.from_file("../app/streamlit_app.py", default_timeout=120)
    at.run()
    assert not at.exception, at.exception
    at.text_input[0].set_value("Heathrow").run()
    at.slider[0].set_value((2019, 2023)).run()
    loc_box = next(sb for sb in at.selectbox if sb.label == "Library location to compare with")
    loc_box.set_value("Demo (synthetic)").run()
    next(b for b in at.button if b.label.startswith("Download 5 years")).click().run()
    assert not at.exception, at.exception
    assert at.session_state["observed"].meta["station"] == "03772"
    frames = [df.value for df in at.dataframe]
    statements = [s for f in frames if "Statement" in f.columns for s in f["Statement"]]
    assert statements, "years-exceeding statements missing"
    assert any(" of the 5 years (2019–2023) exceeded DSY" in s and "for Demo (synthetic) on SWCDH" in s for s in statements)

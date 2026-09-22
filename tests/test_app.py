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

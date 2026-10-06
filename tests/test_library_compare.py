import pandas as pd
import pytest

from summers_dsy.compare import bracket, compare, pick_baseline, resolve_thresholds, return_periods, similarity
from summers_dsy.library import Library
from summers_dsy.metrics import AnalysisConfig
from summers_dsy.synthetic import WarmSpell, demo_library, synthetic_series


@pytest.fixture
def library(tmp_path):
    lib = Library(tmp_path / "lib.sqlite")
    for filename, (series, meta) in demo_library().items():
        lib.add(series, filename, meta)
    return lib


def test_library_round_trip(library):
    entries = library.entries()
    assert len(entries) == len(demo_library())
    assert entries.iloc[0]["period"] == "Baseline"
    s = library.load(int(entries.iloc[0]["id"]))
    assert len(s.data) == 8760 and s.is_typical_year


def test_library_replaces_same_scenario(library):
    n = len(library)
    series, meta = next(iter(demo_library().values()))
    library.add(series, "again.epw", meta)
    assert len(library) == n


def test_library_update_and_delete(library):
    first = int(library.entries().iloc[0]["id"])
    library.update(first, emissions="Low", percentile=10)
    row = library.entries().set_index("id").loc[first]
    assert row["emissions"] == "Low" and row["percentile"] == 10
    library.delete([first])
    assert first not in library.entries()["id"].tolist()


def test_library_rejects_partial_year(tmp_path):
    lib = Library(tmp_path / "l.sqlite")
    s = synthetic_series("x")
    s.data = s.data.iloc[:4000]
    with pytest.raises(ValueError):
        lib.add(s, "x.epw", {"location": "X", "kind": "DSY1", "period": "Baseline"})


def test_similarity_identical_scores_100():
    refs = pd.DataFrame({"wcdh": [100.0, 400.0, 900.0], "t_max": [30.0, 32.0, 35.0]}, index=["a", "b", "c"])
    ranked = similarity({"wcdh": 400.0, "t_max": 32.0}, refs, ["wcdh", "t_max"])
    assert ranked.index[0] == "b"
    assert ranked.iloc[0]["score"] == pytest.approx(100.0)


def test_similarity_weights():
    refs = pd.DataFrame({"wcdh": [100.0, 1000.0], "t_max": [35.0, 30.0]}, index=["a", "b"])
    target = {"wcdh": 110.0, "t_max": 30.0}
    assert similarity(target, refs, ["wcdh", "t_max"], {"wcdh": 1, "t_max": 0}).index[0] == "a"
    assert similarity(target, refs, ["wcdh", "t_max"], {"wcdh": 0, "t_max": 1}).index[0] == "b"


def test_bracket_and_return_periods():
    assert bracket(5, pd.Series({"a": 1, "b": 4, "c": 9})) == ("b", "c")
    rp = return_periods(pd.Series({2018: 10.0, 2019: 30.0, 2020: 20.0}))
    assert rp.loc[2019, "rank"] == 1 and rp.loc[2019, "return_period"] == pytest.approx(4.0)


def test_compare_end_to_end(library):
    entries = library.entries()
    base_id = pick_baseline(entries)
    assert entries.set_index("id").loc[base_id, "kind"] == "TRY"
    refs = {int(i): library.load(int(i)) for i in entries["id"]}
    observed = synthetic_series("2019", typical=False, year=2019, spells=[WarmSpell(215, 5, 9.5)], warming=0.9, seed=12)
    cfg = AnalysisConfig()
    th = resolve_thresholds(cfg, refs[base_id], observed)
    assert "TRY" in th.provenance
    result = compare(observed, refs, cfg, th)
    assert result.meta[result.best_id]["kind"] == "DSY2"
    assert result.headline("Demo").startswith("2019 in Demo was most similar to DSY2")
    # Shared thresholds were used everywhere.
    assert {a.static_threshold for a in result.references.values()} == {th.static_threshold}


def test_years_exceeding_counts_and_statements():
    from summers_dsy.compare import years_exceeding

    years = pd.DataFrame({"swcdh": {2016: 100.0, 2017: 300.0, 2018: 900.0, 2019: 250.0, 2020: 50.0}})
    refs = pd.DataFrame({"swcdh": {"a": 200.0, "b": 1000.0, "c": 60.0}})
    labels = {"a": "DSY1 · 2050s · High emissions · 50th percentile", "b": "DSY3 · 2080s", "c": "TRY · current climate"}
    out = years_exceeding(years, refs, labels, "swcdh", "Zone 7", latest_complete_year=2020).set_index("reference")
    assert out.loc["a", "years_exceeded"] == 3 and out.loc["a", "exceeding_years"] == "2017, 2018, 2019"
    assert out.loc["b", "years_exceeded"] == 0 and out.loc["c", "years_exceeded"] == 4
    assert out.loc["a", "statement"] == ("3 of the last 5 years (2016–2020) exceeded "
                                         "DSY1 · 2050s · High emissions · 50th percentile for Zone 7 on SWCDH")
    # Not "last" when the range stops before the latest complete year; sorted hottest reference first.
    older = years_exceeding(years, refs, labels, "swcdh", "Zone 7", latest_complete_year=2025)
    assert older.iloc[0]["reference"] == "b"
    assert older.iloc[1]["statement"].startswith("3 of the 5 years (2016–2020) exceeded")
    one = years_exceeding(years.loc[[2018]], refs, labels, "swcdh", "Zone 7", latest_complete_year=2018).set_index("reference")
    assert one.loc["a", "statement"].startswith("2018 exceeded")
    assert one.loc["b", "statement"].startswith("2018 did not exceed")

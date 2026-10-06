import numpy as np
import pandas as pd
import pytest

from summers_dsy.comfort import comfort_temperature, daily_comfort, running_mean, seven_day_approximation
from summers_dsy.metrics import AnalysisConfig, analyse, season_mask, warm_events, weighted_degree_hours
from summers_dsy.model import WeatherSeries
from summers_dsy.synthetic import WarmSpell, synthetic_series


def daily(values, start="2019-01-01"):
    return pd.Series(values, index=pd.date_range(start, periods=len(values), freq="D"), dtype=float)


def test_comfort_temperature_tm52():
    assert comfort_temperature(20.0) == pytest.approx(25.4)
    assert comfort_temperature(5.0, clamp=True) == pytest.approx(0.33 * 10 + 18.8)


def test_running_mean_constant_is_constant():
    trm = running_mean(daily([15.0] * 30))
    assert trm.iloc[:7].isna().all()
    assert np.allclose(trm.iloc[7:], 15.0)


def test_seven_day_approximation_weights():
    # Tod-1 = 7 ... Tod-7 = 1 for day index 7.
    vals = [1, 2, 3, 4, 5, 6, 7, 0]
    got = seven_day_approximation(daily(vals)).iloc[7]
    expected = (7 + 0.8 * 6 + 0.6 * 5 + 0.5 * 4 + 0.4 * 3 + 0.3 * 2 + 0.2 * 1) / 3.8
    assert got == pytest.approx(expected)


def test_running_mean_recursion():
    vals = [10.0] * 10 + [20.0] * 10
    trm = running_mean(daily(vals), alpha=0.8)
    # After the step (first 20 degC day is index 10), Trm(11) = 0.2*20 + 0.8*10.
    assert trm.iloc[10] == pytest.approx(10.0)
    assert trm.iloc[11] == pytest.approx(0.2 * 20 + 0.8 * 10)
    assert trm.iloc[12] == pytest.approx(0.2 * 20 + 0.8 * trm.iloc[11])


def test_running_mean_reseeds_after_long_gap():
    vals = [12.0] * 10 + [np.nan] * 6 + [18.0] * 10
    trm = running_mean(daily(vals), max_fill_days=3)
    assert trm.iloc[17:23].isna().all()  # needs 7 valid days before reseeding
    assert trm.iloc[23] == pytest.approx(18.0)


def test_cyclic_running_mean_has_no_gaps():
    s = synthetic_series("x")
    table = daily_comfort(s.dry_bulb, cyclic=True)
    assert table["trm"].notna().all()
    assert (table["t_upper"] - table["t_comf"]).eq(3.0).all()


def test_weighted_degree_hours():
    t = pd.Series([20.0, 25.0, 27.0, 23.0])
    assert weighted_degree_hours(t, 24.0).tolist() == [0.0, 1.0, 9.0, 0.0]


def test_season_mask_tm59_dates():
    idx = pd.date_range("2019-04-30 23:00", "2019-10-01 00:00", freq="h")
    mask = season_mask(idx)
    assert not mask[0] and mask[1] and not mask[-1]
    assert idx[mask][-1] == pd.Timestamp("2019-09-30 23:00")


def test_warm_events():
    wcdh = daily([0, 5, 10, 0, 0, 1, 2, 3, 4, 0])
    tmax = daily([20, 26, 28, 20, 20, 25, 25, 26, 27, 20])
    events = warm_events(wcdh, tmax)
    assert [e.duration_days for e in events] == [2, 4]
    assert events[0].severity == 15 and events[0].intensity == 10
    assert events[1].peak_temperature == 27


def test_wcdh_hand_calculation():
    # Constant 20 degC all year gives Trm = 20, Tcomf = 25.4; one 5-hour spike at 30 degC in July.
    idx = pd.date_range("2019-01-01", "2019-12-31 23:00", freq="h")
    t = pd.Series(20.0, index=idx)
    spike = (idx >= "2019-07-10 12:00") & (idx < "2019-07-10 17:00")
    t[spike] = 30.0
    s = WeatherSeries("spike", t.to_frame("dry_bulb"))
    res = analyse(s, AnalysisConfig(static_threshold=28.0, twcdh_offset=1.0))
    # The spike nudges that day's mean, but Tcomf for 10 July uses days up to 9 July only.
    assert res.daily.loc["2019-07-10", "t_comf"] == pytest.approx(25.4)
    assert res.metrics["wcdh"] == pytest.approx(5 * (30 - 25.4) ** 2)
    assert res.metrics["twcdh"] == pytest.approx(5 * (30 - 26.4) ** 2)
    assert res.metrics["swcdh"] == pytest.approx(5 * 2.0**2)
    assert res.metrics["hours_above_upper"] == 5  # 30 - 28.4 = 1.6 -> rounds to 2 K
    assert res.metrics["max_delta_t"] == 2
    assert res.metrics["max_daily_weighted_exceedance"] == 10
    assert res.metrics["max_event_duration"] == 1


def test_intense_vs_long_spells_are_distinguished():
    intense = analyse(synthetic_series("dsy2", spells=[WarmSpell(210, 4, 10)], seed=1), AnalysisConfig(static_threshold=24, twcdh_offset=0))
    long = analyse(synthetic_series("dsy3", spells=[WarmSpell(210, 24, 5)], seed=1), AnalysisConfig(static_threshold=24, twcdh_offset=0))
    assert intense.metrics["peak_daily_wcdh"] > long.metrics["peak_daily_wcdh"]
    assert long.metrics["max_event_duration"] > intense.metrics["max_event_duration"]


def test_nightly_means_window_and_metrics():
    from summers_dsy.metrics import nightly_means

    idx = pd.date_range("2019-01-01", "2019-12-31 23:00", freq="h")
    t = pd.Series(15.0, index=idx)
    # The night starting on 24 July (22:00 on the 24th to 07:00 on the 25th) is warm: 29 degC.
    warm = (idx >= "2019-07-24 22:00") & (idx < "2019-07-25 07:00")
    t[warm] = 29.0
    t[idx == pd.Timestamp("2019-07-25 07:00")] = 40.0  # 07:00 is outside the night
    nights = nightly_means(t)
    assert nights.loc["2019-07-24", "night_hours"] == 9
    assert nights.loc["2019-07-24", "night_mean"] == pytest.approx(29.0)
    assert nights.loc["2019-07-25", "night_mean"] == pytest.approx(15.0)

    res = analyse(WeatherSeries("n", t.to_frame("dry_bulb")), AnalysisConfig(static_threshold=25, twcdh_offset=0))
    assert res.metrics["night_max_mean"] == pytest.approx(29.0)
    assert res.metrics["nights_above"] == 1
    res20 = analyse(WeatherSeries("n", t.to_frame("dry_bulb")),
                    AnalysisConfig(static_threshold=25, twcdh_offset=0, night_threshold=30))
    assert res20.metrics["nights_above"] == 0


def test_night_needs_enough_hours():
    from summers_dsy.metrics import nightly_means

    idx = pd.date_range("2019-07-01 22:00", periods=9, freq="h")
    t = pd.Series([20.0] * 6 + [np.nan] * 3, index=idx)
    assert np.isnan(nightly_means(t.dropna()).loc["2019-07-01", "night_mean"])


def test_seasonal_means_use_the_analysis_year_only():
    # A previous-December spin-up at 100 degC must not leak into this year's winter mean.
    idx = pd.date_range("2018-12-01", "2019-12-31 23:00", freq="h")
    t = pd.Series(np.where(idx.month.isin([12, 1, 2]), 2.0, 10.0), index=idx)
    t[idx.year == 2018] = 100.0
    t[(idx.month.isin([6, 7, 8])) & (idx.hour == 14)] = 25.0  # daily maximum in summer
    res = analyse(WeatherSeries("s", t.to_frame("dry_bulb")), AnalysisConfig(static_threshold=25, twcdh_offset=0))
    m = res.metrics
    assert m["mean_t_djf"] == pytest.approx(2.0)
    assert m["mean_t_mam"] == pytest.approx(10.0)
    assert m["mean_t_son"] == pytest.approx(10.0)
    assert m["mean_t_jja"] == pytest.approx((23 * 10 + 25) / 24)
    assert m["jja_mean_daily_max"] == pytest.approx(25.0)


def test_season_coverage_counts_against_the_whole_season():
    idx = pd.date_range("2019-01-01", "2019-08-24 23:00", freq="h")  # record stops in late August
    res = analyse(WeatherSeries("t", pd.DataFrame({"dry_bulb": 15.0}, index=idx)),
                  AnalysisConfig(static_threshold=25, twcdh_offset=0))
    assert res.metrics["season_coverage"] == pytest.approx(116 / 153, abs=1e-9)

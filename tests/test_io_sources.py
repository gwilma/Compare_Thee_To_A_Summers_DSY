import gzip
import json

import pandas as pd
import pytest

from summers_dsy.io import load_weather_file, parse_reference_name, reference_label
from summers_dsy.io.parsers import WeatherFileError
from summers_dsy.sources import fetchers
from summers_dsy.synthetic import synthetic_series, to_epw


def test_epw_round_trip_typical_year():
    s = synthetic_series("x")
    text = to_epw(s, year=1989)
    parsed = load_weather_file(text.encode(), "London_LHR_DSY2_2050High50.epw")
    assert parsed.is_typical_year
    assert len(parsed.data) == 8760
    assert parsed.meta["source_year"] == 1989
    assert parsed.meta["kind"] == "DSY2" and parsed.meta["period"] == "2050s"
    assert parsed.dry_bulb.iloc[0] == pytest.approx(round(s.dry_bulb.iloc[0], 1))
    assert parsed.data.index[0] == pd.Timestamp("2001-01-01 00:00")


def test_epw_actual_year_keeps_dates():
    s = synthetic_series("x", year=2019, typical=False)
    parsed = load_weather_file(to_epw(s), "heathrow_2019.epw")
    assert not parsed.is_typical_year
    assert parsed.data.index[0] == pd.Timestamp("2019-01-01 00:00")


def test_csv_month_day_hour():
    rows = ["Some title line", "Month,Day,Hour,Dry bulb temperature (C),RH"]
    for h in range(1, 25):
        rows.append(f"7,1,{h},{15 + h / 2},60")
    parsed = load_weather_file("\n".join(rows).encode(), "Manchester_TRY.csv")
    assert parsed.is_typical_year
    assert parsed.dry_bulb.iloc[0] == 15.5
    assert parsed.data.index[0] == pd.Timestamp("2001-07-01 00:00")


def test_csv_datetime_subhourly():
    rows = ["datetime,temperature"]
    for i, ts in enumerate(pd.date_range("2019-07-01", periods=12, freq="30min")):
        rows.append(f"{ts.isoformat()},{20 + i}")
    parsed = load_weather_file("\n".join(rows), "obs.csv")
    assert not parsed.is_typical_year
    # On-the-hour readings are preferred over the half-hourly ones; the final
    # 05:30 reading is the only one near 06:00.
    assert parsed.dry_bulb.tolist() == [20.0, 22.0, 24.0, 26.0, 28.0, 30.0, 31.0]


def test_csv_without_header_is_rejected():
    with pytest.raises(WeatherFileError):
        load_weather_file("1,2,3\n4,5,6\n", "bad.csv")


MIDAS_SAMPLE = """Conventions,G,BADC-CSV,1
title,G,uk-hourly-weather-obs
observation_station,G,heathrow
location,G,51.479,-0.449
data
ob_time,id,met_domain_name,air_temperature
2019-07-25 12:00:00,708,SYNOP,35.1
2019-07-25 13:00:00,708,SYNOP,36.4
2019-07-25 13:00:00,708,SYNOP,36.2
2019-07-25 15:00:00,708,SYNOP,35.9
end data
"""


def test_midas_parse():
    s = load_weather_file(MIDAS_SAMPLE, "midas.csv")
    assert s.source == "midas"
    assert s.meta["latitude"] == 51.479
    assert s.dry_bulb.loc["2019-07-25 13:00"] == pytest.approx(36.3)
    assert s.dry_bulb.loc["2019-07-25 14:00"] == pytest.approx(36.1)  # interpolated gap


def test_isd_parse_filters_missing_and_bad_quality():
    raw = (
        'STATION,DATE,TMP\n'
        '03772099999,2019-07-25T11:50:00,"+0340,1"\n'
        '03772099999,2019-07-25T12:00:00,"+0351,1"\n'
        '03772099999,2019-07-25T12:20:00,"+9999,9"\n'
        '03772099999,2019-07-25T12:50:00,"+0600,3"\n'
        '03772099999,2019-07-25T13:00:00,"+0364,5"\n'
    ).encode()
    t = fetchers.parse_isd(raw)
    assert t.tolist() == [34.0, 35.1, 36.4]
    hourly = fetchers.to_hourly(t)
    assert hourly.tolist() == [35.1, 36.4]  # 12:00 SYNOP beats the 11:50 METAR


def test_meteostat_parse():
    csv = "2019-07-25,12,35.1,,,,,,,,,,\n2019-07-25,13,36.4,,,,,,,,,,\n"
    t = fetchers.parse_meteostat(gzip.compress(csv.encode()))
    assert t.index[1] == pd.Timestamp("2019-07-25 13:00")
    assert t.iloc[1] == 36.4


def test_open_meteo_fetch(monkeypatch):
    times = pd.date_range("2019-01-01", "2019-12-31 23:00", freq="h")
    payload = {"elevation": 25, "hourly": {"time": [t.strftime("%Y-%m-%dT%H:%M") for t in times], "temperature_2m": [12.0] * len(times)}}
    seen = {}

    def fake_get(url, params=None, **kw):
        seen.update(params)
        return json.dumps(payload).encode()

    monkeypatch.setattr(fetchers, "get_bytes", fake_get)
    s = fetchers.fetch_open_meteo(51.48, -0.45, 2019, 2019)
    assert seen["start_date"] == "2019-01-01" and seen["end_date"] == "2019-12-31"
    assert len(s.data) == 8760 and s.years == [2019]


def test_meteostat_fetch_skips_missing_years(monkeypatch):
    csv = "\n".join(f"2019-01-01,{h},{5 + h / 10},,,,,,,,,," for h in range(24))

    def fake_get(url, **kw):
        return gzip.compress(csv.encode()) if "/2019/" in url else None

    monkeypatch.setattr(fetchers, "get_bytes", fake_get)
    s = fetchers.fetch_meteostat("03772", 2018, 2019)
    assert s.years == [2019]


def test_midas_fetch_requires_token(monkeypatch):
    monkeypatch.delenv("CEDA_TOKEN", raising=False)
    with pytest.raises(fetchers.SourceError):
        fetchers.fetch_midas("greater-london", 708, "heathrow", 2019, 2019)


@pytest.mark.parametrize(
    "name, expected",
    [
        ("London_LHR_DSY1_2050High50.epw", ("London (Heathrow)", "DSY1", "2050s", "High", 50)),
        ("Manchester_DSY3_2080s_Medium_90.csv", ("Manchester", "DSY3", "2080s", "Medium", 90)),
        ("GB_Edinburgh_TRY.epw", ("Edinburgh", "TRY", "Baseline", None, None)),
        ("LondonLWC_DSY2.epw", ("London (Weather Centre)", "DSY2", "Baseline", None, None)),
        ("Birmingham_DSY_2020_Low_10pc.epw", ("Birmingham", "DSY1", "2020s", "Low", 10)),
        ("Z1_DSY1_2020s_HIGH10_CIBSE_v1.1.epw", ("Zone 1", "DSY1", "2020s", "High", 10)),
        ("Z12_TRY_2030s_HIGH50_CIBSE_v1.1.epw", ("Zone 12", "TRY", "2030s", "High", 50)),
        ("Zone 28_DSY3_2080s_LOW90_CIBSE_v1.1.epw", ("Zone 28", "DSY3", "2080s", "Low", 90)),
        ("Glasgow_DSY2_2050High50.epw", ("Glasgow", "DSY2", "2050s", "High", 50)),
    ],
)
def test_reference_name_parsing(name, expected):
    m = parse_reference_name(name)
    assert (m["location"], m["kind"], m["period"], m["emissions"], m["percentile"]) == expected


def test_reference_label():
    meta = {"kind": "DSY1", "period": "2050s", "emissions": "High", "percentile": 50}
    assert reference_label(meta) == "DSY1 · 2050s · High emissions · 50th percentile"

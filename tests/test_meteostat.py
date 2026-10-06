import gzip
import json
import sqlite3

import numpy as np
import pandas as pd
import pytest

from summers_dsy.sources import meteostat as ms


def new_format(year: int, temp=lambda ts: 10.0, source=lambda ts: "isd_lite", months=None) -> bytes:
    """A file in Meteostat's current hourly layout (header; year, month, day, hour, values, *_source)."""
    idx = pd.date_range(f"{year}-01-01", f"{year}-12-31 23:00", freq="h")
    if months:
        idx = idx[idx.month.isin(months)]
    rows = ["year,month,day,hour,temp,temp_source,rhum,rhum_source"]
    rows += [f"{t.year},{t.month},{t.day},{t.hour},{temp(t):.1f},{source(t)},70,{source(t)}" for t in idx]
    return gzip.compress(("\n".join(rows) + "\n").encode())


def test_parse_current_format():
    df = ms.parse_hourly(new_format(2019, temp=lambda t: t.hour))
    assert df.index[0] == pd.Timestamp("2019-01-01 00:00")
    assert df.loc["2019-07-25 13:00", "temp"] == 13.0
    assert (df["source"] == "isd_lite").all()


def test_parse_legacy_format():
    csv = "2019-07-25,12,35.1,,,,,,,,,,\n2019-07-25,13,36.4,,,,,,,,,,\n"
    df = ms.parse_hourly(gzip.compress(csv.encode()))
    assert df.index[1] == pd.Timestamp("2019-07-25 13:00") and df["temp"].iloc[1] == 36.4


def test_fetch_years_drops_model_data_and_adds_december_spinup(monkeypatch):
    model_hours = lambda t: "metno_forecast" if (t.month == 7 and t.day == 20 and 3 <= t.hour < 5) else "metar"
    files = {2018: new_format(2018, temp=lambda t: 1.0), 2019: new_format(2019, source=model_hours),
             2020: None}

    def fake_get(url, **kw):
        year = int(url.split("/")[-2])
        return files.get(year) if "data.meteostat.net" in url else None

    monkeypatch.setattr(ms, "get_bytes", fake_get)
    calls = []
    s = ms.fetch_years("03772", 2019, 2020, progress=lambda k, n, y: calls.append((k, n, y)))
    assert calls[-1] == (3, 3, 2020)
    assert s.data.index[0] == pd.Timestamp("2018-12-01 00:00")  # previous December only
    assert s.meta["missing_years"] == [2020]
    assert s.meta["model_share"][2019] == pytest.approx(2 / 8760)
    # The 2-hour model gap is re-filled by interpolation from the observations either side.
    assert s.dry_bulb.loc["2019-07-20 03:00"] == pytest.approx(10.0)
    with_model = ms.fetch_years("03772", 2019, 2019, include_model=True)
    assert with_model.meta["include_model"] is True


def test_fetch_years_falls_back_to_legacy_endpoint(monkeypatch):
    legacy = gzip.compress("2019-07-01,0,15.0,,,,,,,,,,\n2019-07-01,1,15.5,,,,,,,,,,\n".encode())
    monkeypatch.setattr(ms, "get_bytes", lambda url, **kw: legacy if "bulk.meteostat.net" in url and "/2019/" in url else None)
    s = ms.fetch_years("03772", 2019, 2019)
    assert s.dry_bulb.tolist() == [15.0, 15.5]


def test_fetch_years_with_no_data_raises(monkeypatch):
    monkeypatch.setattr(ms, "get_bytes", lambda url, **kw: None)
    with pytest.raises(ms.SourceError):
        ms.fetch_years("XXXXX", 2019, 2020)


def test_long_gaps_are_left_empty():
    from summers_dsy.model import to_hourly

    idx = pd.date_range("2019-07-01", periods=20, freq="h")
    t = pd.Series(np.arange(20.0), index=idx)
    gappy = t.drop(idx[2:5]).drop(idx[8:16])  # a 3-hour and an 8-hour gap
    out = to_hourly(gappy, max_gap_hours=6)
    assert out.iloc[2:5].tolist() == [2.0, 3.0, 4.0]
    assert out.iloc[8:16].isna().all()


def make_db(path):
    con = sqlite3.connect(path)
    con.executescript("""
        CREATE TABLE stations (id, country, region, latitude, longitude, elevation, timezone);
        CREATE TABLE names (station, language, name);
        CREATE TABLE identifiers (station, key, value);
        CREATE TABLE inventory (station, provider, parameter, start, "end", completeness);
        INSERT INTO stations VALUES ('03772','GB','ENG',51.4833,-0.45,25,'Europe/London'),
                                    ('03334','GB','ENG',53.35,-2.2833,69,'Europe/London'),
                                    ('99999','GB','ENG',52.0,-1.0,10,'Europe/London'),
                                    ('07150','FR','IDF',48.97,2.44,60,'Europe/Paris');
        INSERT INTO names VALUES ('03772','en','London Heathrow Airport'),('03334','en','Manchester Airport'),
                                 ('99999','en','Forecast Only'),('07150','en','Paris Le Bourget');
        INSERT INTO identifiers VALUES ('03772','wmo','03772'),('03772','icao','EGLL');
        INSERT INTO inventory VALUES ('03772','isd_lite','temp','1948-12-01','2025-08-24',100),
                                     ('03772','metar','temp','2015-09-06','2026-10-04',95),
                                     ('03334','isd_lite','temp','1933-09-30','2025-08-24',100),
                                     ('99999','dwd_mosmix','temp','2018-01-01','2026-10-01',90),
                                     ('07150','isd_lite','temp','1950-01-01','2026-10-01',100);
    """)
    con.commit()
    con.close()


def test_stations_from_db(tmp_path):
    db = tmp_path / "stations.db"
    make_db(db)
    rows = ms.stations_from_db(db, ["GB"])
    assert [r["id"] for r in rows] == ["03772", "03334"]  # hourly observations only, UK only
    heathrow = rows[0]
    assert heathrow["icao"] == "EGLL" and heathrow["hourly_start"] == "1948-12-01" and heathrow["hourly_end"] == "2026-10-04"


def test_search_and_nearest():
    st = ms.bundled_stations()
    assert len(st) > 100
    assert "03772" in set(ms.search_stations(st, "heathrow")["id"])
    assert "03772" in set(ms.search_stations(st, "EGLL")["id"])
    near = ms.nearest_stations(st, 51.47, -0.45, 3)
    assert near.iloc[0]["id"] == "03772" and near.iloc[0]["distance_km"] < 3


def test_geocode_postcode(monkeypatch):
    payload = {"status": 200, "result": {"postcode": "SW1A 1AA", "latitude": 51.501, "longitude": -0.1416,
                                          "admin_district": "Westminster", "region": "London"}}
    monkeypatch.setattr(ms, "get_bytes", lambda url, **kw: json.dumps(payload).encode() if url.endswith("SW1A1AA") else None)
    lat, lon, label = ms.geocode_postcode("sw1a 1aa")
    assert (lat, lon) == (51.501, -0.1416) and "Westminster" in label
    with pytest.raises(ms.SourceError):
        ms.geocode_postcode("ZZ1 1ZZ")


def test_location_guess():
    assert ms.station_location_guess("Manchester Airport") == "Manchester"
    assert ms.station_location_guess("London Heathrow Airport") == "London (Heathrow)"
    assert ms.station_location_guess("Aberdaron") is None

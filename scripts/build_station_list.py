"""Build the bundled list of UK Meteostat stations that report hourly temperature.

    python scripts/build_station_list.py path/to/stations.db [--countries GB]

stations.db is Meteostat's station database (https://data.meteostat.net/stations.db, mirrored in
https://github.com/meteostat/weather-stations). Station data © Meteostat, CC BY 4.0.
"""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from summers_dsy.sources.meteostat import STATION_COLUMNS, stations_from_db  # noqa: E402

OUT = ROOT / "src/summers_dsy/sources/data/meteostat_stations_uk.csv"


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("db")
    ap.add_argument("--countries", nargs="+", default=["GB"])
    ap.add_argument("--out", default=str(OUT))
    args = ap.parse_args()
    stations = stations_from_db(args.db, args.countries)
    with open(args.out, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(STATION_COLUMNS)
        for s in stations:
            w.writerow([s[c] for c in STATION_COLUMNS])
    print(f"{len(stations)} stations written to {args.out}")


if __name__ == "__main__":
    main()

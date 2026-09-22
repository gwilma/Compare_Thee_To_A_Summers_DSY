"""Persistent local library of reference (CIBSE TRY/DSY) weather files.

Each entry keeps its metadata and its hourly dry-bulb column (8760 float32
values on the nominal year). Keeping the temperatures, rather than frozen
metric values, means every metric can be recomputed when thresholds or
settings change. The database lives in the local data directory, which is
git-ignored: CIBSE files are licensed and should stay on the licensee's machine.
"""

from __future__ import annotations

import json
import sqlite3
from contextlib import closing
from datetime import datetime, timezone
from pathlib import Path

import numpy as np
import pandas as pd

from .io.naming import reference_label
from .model import NOMINAL_YEAR, WeatherSeries
from .sources.http import data_dir

META_FIELDS = ["location", "kind", "period", "emissions", "percentile"]

SCHEMA = """
CREATE TABLE IF NOT EXISTS reference_files (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    filename TEXT NOT NULL,
    location TEXT NOT NULL,
    kind TEXT NOT NULL,
    period TEXT NOT NULL,
    emissions TEXT NOT NULL DEFAULT '',
    percentile INTEGER NOT NULL DEFAULT 0,
    source_year INTEGER,
    meta TEXT NOT NULL DEFAULT '{}',
    added_at TEXT NOT NULL,
    temps BLOB NOT NULL,
    UNIQUE (location, kind, period, emissions, percentile)
);
"""

_HOURS = pd.date_range(f"{NOMINAL_YEAR}-01-01", periods=8760, freq="h")


class Library:
    def __init__(self, path: str | Path | None = None):
        self.path = Path(path) if path else data_dir() / "library.sqlite"
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with closing(self._connect()) as con, con:
            con.executescript(SCHEMA)

    def _connect(self) -> sqlite3.Connection:
        return sqlite3.connect(self.path)

    # ---------------------------------------------------------------- writes

    def add(self, series: WeatherSeries, filename: str, meta: dict) -> int:
        """Store a typical-year file; replaces any entry with the same location/kind/period/scenario."""
        missing = [f for f in ("location", "kind", "period") if not meta.get(f)]
        if missing:
            raise ValueError(f"Missing metadata: {', '.join(missing)}")
        temps = series.dry_bulb.reindex(_HOURS)
        if temps.isna().mean() > 0.01:
            raise ValueError(f"{filename}: expected a full 8760-hour typical year ({temps.notna().sum()} hours matched)")
        temps = temps.interpolate(limit_direction="both").astype(np.float32)
        row = {
            "filename": filename,
            "location": meta["location"],
            "kind": meta["kind"],
            "period": meta["period"],
            "emissions": meta.get("emissions") or "",
            "percentile": int(meta.get("percentile") or 0),
            "source_year": series.meta.get("source_year"),
            "meta": json.dumps({k: v for k, v in series.meta.items() if isinstance(v, (str, int, float, bool)) or v is None}),
            "added_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "temps": temps.to_numpy().tobytes(),
        }
        with closing(self._connect()) as con, con:
            cur = con.execute(
                f"INSERT OR REPLACE INTO reference_files ({', '.join(row)}) VALUES ({', '.join('?' * len(row))})",
                list(row.values()),
            )
            return int(cur.lastrowid)

    def update(self, entry_id: int, **fields) -> None:
        allowed = {k: v for k, v in fields.items() if k in META_FIELDS}
        if not allowed:
            return
        if "emissions" in allowed:
            allowed["emissions"] = allowed["emissions"] or ""
        if "percentile" in allowed:
            allowed["percentile"] = int(allowed["percentile"] or 0)
        sets = ", ".join(f"{k} = ?" for k in allowed)
        with closing(self._connect()) as con, con:
            con.execute(f"UPDATE reference_files SET {sets} WHERE id = ?", [*allowed.values(), entry_id])

    def delete(self, ids: list[int]) -> None:
        with closing(self._connect()) as con, con:
            con.executemany("DELETE FROM reference_files WHERE id = ?", [(int(i),) for i in ids])

    # ----------------------------------------------------------------- reads

    def entries(self, location: str | None = None) -> pd.DataFrame:
        """Library contents (without the temperature data), sorted sensibly."""
        sql = "SELECT id, filename, location, kind, period, emissions, percentile, source_year, added_at FROM reference_files"
        params: list = []
        if location:
            sql += " WHERE location = ?"
            params.append(location)
        with closing(self._connect()) as con:
            df = pd.read_sql_query(sql, con, params=params)
        if df.empty:
            return df.assign(label=pd.Series(dtype=str))
        df["emissions"] = df["emissions"].replace("", None)
        df["percentile"] = df["percentile"].replace(0, None)
        df["label"] = df.apply(lambda r: reference_label(r.to_dict()), axis=1)
        order = {"Baseline": 0, "2020s": 1, "2050s": 2, "2080s": 3}
        return (
            df.assign(_p=df["period"].map(order).fillna(9), _e=df["emissions"].map({"Low": 0, "Medium": 1, "High": 2}).fillna(-1))
            .sort_values(["location", "_p", "kind", "_e", "percentile"])
            .drop(columns=["_p", "_e"])
            .reset_index(drop=True)
        )

    def locations(self) -> list[str]:
        with closing(self._connect()) as con:
            return [r[0] for r in con.execute("SELECT DISTINCT location FROM reference_files ORDER BY location")]

    def load(self, entry_id: int) -> WeatherSeries:
        with closing(self._connect()) as con:
            row = con.execute(
                "SELECT filename, location, kind, period, emissions, percentile, meta, temps FROM reference_files WHERE id = ?",
                (int(entry_id),),
            ).fetchone()
        if row is None:
            raise KeyError(entry_id)
        filename, location, kind, period, emissions, percentile, meta_json, blob = row
        temps = np.frombuffer(blob, dtype=np.float32).astype(float)
        meta = {
            **json.loads(meta_json),
            "location": location,
            "kind": kind,
            "period": period,
            "emissions": emissions or None,
            "percentile": percentile or None,
            "typical_year": True,
            "library_id": int(entry_id),
            "filename": filename,
        }
        name = f"{location} · {reference_label(meta)}"
        return WeatherSeries(name, pd.DataFrame({"dry_bulb": temps}, index=_HOURS), "library", meta)

    def __len__(self) -> int:
        with closing(self._connect()) as con:
            return int(con.execute("SELECT COUNT(*) FROM reference_files").fetchone()[0])

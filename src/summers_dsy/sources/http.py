"""HTTP download with a small on-disk cache, so repeat analyses don't refetch."""

from __future__ import annotations

import hashlib
import json
import os
import time
from pathlib import Path

import requests

USER_AGENT = "summers-dsy/0.1 (overheating weather analysis)"
DEFAULT_TIMEOUT = 60


class SourceError(RuntimeError):
    """A data source could not supply the requested data."""


def data_dir() -> Path:
    """Local data directory (library + cache); override with SUMMERS_DSY_DATA."""
    path = Path(os.environ.get("SUMMERS_DSY_DATA", Path.cwd() / "data"))
    path.mkdir(parents=True, exist_ok=True)
    return path


def cache_dir() -> Path:
    path = data_dir() / "cache"
    path.mkdir(parents=True, exist_ok=True)
    return path


def get_bytes(
    url: str,
    params: dict | None = None,
    headers: dict | None = None,
    max_age_s: float | None = None,
    session: requests.Session | None = None,
    not_found_ok: bool = False,
) -> bytes | None:
    """GET ``url`` with caching. ``max_age_s=None`` caches forever (use for past years).

    Returns None for a 404 when ``not_found_ok`` is set; raises :class:`SourceError` otherwise.
    """
    key = hashlib.sha256(json.dumps([url, sorted((params or {}).items())]).encode()).hexdigest()[:32]
    path = cache_dir() / f"{key}.bin"
    if path.exists() and (max_age_s is None or time.time() - path.stat().st_mtime < max_age_s):
        return path.read_bytes()

    sess = session or requests.Session()
    hdrs = {"User-Agent": USER_AGENT, **(headers or {})}
    try:
        resp = sess.get(url, params=params, headers=hdrs, timeout=DEFAULT_TIMEOUT)
    except requests.RequestException as exc:
        raise SourceError(f"Could not reach {url}: {exc}") from exc
    if resp.status_code == 404 and not_found_ok:
        return None
    if resp.status_code in (401, 403):
        raise SourceError(f"Access denied by {url} (HTTP {resp.status_code}); check credentials or network policy")
    if not resp.ok:
        raise SourceError(f"{url} returned HTTP {resp.status_code}: {resp.text[:200]}")
    path.write_bytes(resp.content)
    return resp.content

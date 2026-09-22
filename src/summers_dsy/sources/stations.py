"""UK observing stations near the CIBSE weather-file locations.

Identifiers:

* ``wmo`` - WMO block/station number; Meteostat uses it as the station id for these sites.
* ``usaf_wban`` - NOAA Integrated Surface Database id (USAF + WBAN).
* ``midas`` - Met Office MIDAS Open (county folder, src_id, site folder) on CEDA.

Coordinates and ids are from the public station registries. Check any
station against the source before relying on it; the app also takes custom ids.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Station:
    name: str
    cibse_location: str
    latitude: float
    longitude: float
    wmo: str
    usaf_wban: str | None = None
    midas: tuple[str, int, str] | None = None  # (county, src_id, site)

    @property
    def label(self) -> str:
        return f"{self.name} (CIBSE: {self.cibse_location})"


STATIONS: list[Station] = [
    Station("London Heathrow", "London (Heathrow)", 51.479, -0.449, "03772", "03772099999", ("greater-london", 708, "heathrow")),
    Station("London St James's Park", "London (Weather Centre)", 51.504, -0.129, "03770", "03770099999"),
    Station("London Gatwick", "London (Gatwick)", 51.148, -0.190, "03776", "03776099999"),
    Station("Belfast Aldergrove", "Belfast", 54.664, -6.224, "03917", "03917099999"),
    Station("Birmingham Airport", "Birmingham", 52.454, -1.748, "03534", "03534099999"),
    Station("Cardiff Airport (Rhoose)", "Cardiff", 51.397, -3.343, "03716", "03716099999"),
    Station("Edinburgh Airport", "Edinburgh", 55.950, -3.373, "03160", "03160099999"),
    Station("Glasgow Airport", "Glasgow", 55.872, -4.433, "03135", "03135099999"),
    Station("Leeds Bradford Airport", "Leeds", 53.866, -1.661, "03347", "03347099999"),
    Station("Manchester Airport", "Manchester", 53.354, -2.275, "03334", "03334099999"),
    Station("Newcastle Airport", "Newcastle", 55.037, -1.692, "03245", "03245099999"),
    Station("Norwich Airport", "Norwich", 52.676, 1.283, "03492", "03492099999"),
    Station("Nottingham Watnall", "Nottingham", 53.005, -1.250, "03354", "03354099999"),
    Station("Plymouth Mount Batten", "Plymouth", 50.354, -4.121, "03827", "03827099999"),
    Station("Southampton Airport", "Southampton", 50.950, -1.357, "03865", "03865099999"),
    Station("Brize Norton", "Swindon", 51.758, -1.576, "03649", "03649099999"),
]

BY_NAME = {s.name: s for s in STATIONS}


def nearest_station(lat: float, lon: float) -> Station:
    """Closest listed station (equirectangular distance is plenty at UK scale)."""
    import math

    def dist(s: Station) -> float:
        x = (s.longitude - lon) * math.cos(math.radians((s.latitude + lat) / 2))
        return math.hypot(x, s.latitude - lat)

    return min(STATIONS, key=dist)

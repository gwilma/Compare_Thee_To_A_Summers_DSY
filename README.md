# Compare thee to a summer's DSY

Overheating analysis of hourly weather for the UK. The app fetches observed hourly temperatures, computes the
metrics used to assess overheating and to select CIBSE Design Summer Years, and compares a real summer with your
library of CIBSE weather files. The result is a statement like:

> 2019 at London Heathrow was most similar to **DSY1 · 2050s · High emissions · 50th percentile** for London (Heathrow).

## Quick start

```bash
pip install -e ".[dev]"
streamlit run app/streamlit_app.py
pytest
```

1. **① Observed weather**: choose a source and station, pick a range of years, and fetch the data. Alternatively, upload EPW, CSV or MIDAS files.
2. **② DSY library**: upload your licensed CIBSE TRY/DSY files (EPW or CSV). Metadata such as location, DSY1/2/3,
   period, emissions and percentile is read from the file names, and you can edit it before adding. If you have no
   files yet, *Load synthetic demo set* adds clearly labelled fake data so you can try the app.
3. **③ Compare**: pick an observed year and a library location. The app ranks every reference file by similarity
   across the metrics and weights you choose.

### Several years from Meteostat

In the app, choose **Meteostat** as the data source. Then:
1. Find a station by name or code, or list the stations nearest to a UK postcode (via postcodes.io) or coordinates.
   A map shows them. The bundled list holds the 149 UK stations with hourly temperature observations; it can be
   refreshed from Meteostat.
2. Choose a range of years and download them. One file per year comes from `data.meteostat.net`, plus the
   December before for the running mean.

Model-forecast values that Meteostat uses to fill gaps are dropped unless you ask for them. Years with under 80%
of the May–September hours are not counted.

On **Compare**, the section *How many years exceeded each … reference file* gives statements such as
"5 of the last 10 years (2016–2025) exceeded DSY1 · 2050s · High emissions · 50th percentile for Zone 7 on SWCDH".
They use only the chosen location's files, on a metric you pick (default SWCDH).

The Excel workbook does the same with the `DownloadMeteostatYears` macro (Meteostat, Years and Years vs DSY
sheets). Downloading needs Excel for Windows. On a Mac, save the files into the download folder yourself.

Station data © Meteostat (CC BY 4.0). Check meteostat.net's terms for the hourly data before commercial use.
`python scripts/build_station_list.py stations.db` rebuilds the bundled station list.

### On an Ubuntu VPS

`scripts/deploy_vps.sh` installs the app as a service behind nginx, with a password and optional HTTPS.
Re-run it to update.

```bash
curl -fsSLO https://raw.githubusercontent.com/gwilma/Compare_Thee_To_A_Summers_DSY/claude/weather-overheating-analysis-y14bhk/scripts/deploy_vps.sh
sudo DOMAIN=dsy.example.com EMAIL=you@example.com bash deploy_vps.sh   # or just: sudo bash deploy_vps.sh
```

### Excel edition

`excel/summers_dsy.xlsx` does the same analysis with live Excel formulas, with no Python needed. Paste a year of
hourly data (Year, Month, Day, Hour, dry bulb: EPW columns 1–4 and 7) into **Hourly**. **Dashboard** and **Summary**
recalculate. Build a **Library** by running each CIBSE file through the workbook and pasting its Summary row as
values. **Compare** then ranks the library. Start with the *Read me* sheet. It ships with synthetic demo data.

**Importing a whole folder.** `excel/summers_dsy.xlsm` is the same workbook with an `ImportWeatherFolder` macro
(source in `excel/FolderImport.bas`):
- It walks a folder and all its sub-folders and loads each `.epw`/`.csv` file in turn.
- It adds each file's metrics to the Library, replacing any row for the same file identity.
- It reads location, file type, period, emissions and percentile from CIBSE 2016 or 2025 file names, falling back
  to the folder name.
- It derives each location's SWCDH/TWCDH thresholds from that location's TRY and stores them with each row.
  `UseLibraryThresholds` copies them into Settings.

Everything is recorded on the **Import log** sheet.

`python excel/lint_vba.py` checks `FolderImport.bas` against rules that Excel enforces but LibreOffice does not; `add_macros.py` runs it before building.

`python excel/build_workbook.py` regenerates the `.xlsx`. `python excel/add_macros.py` then compiles
`FolderImport.bas` into the `.xlsm` (this needs LibreOffice and `python3-uno`).
`python excel/verify_workbook.py <recalculated.xlsx>` checks every metric and the similarity ranking against the
Python package. `python excel/test_folder_import.py` runs the macro inside LibreOffice on a test folder tree and
checks every imported row.

## Metrics

| Metric | Definition |
|---|---|
| Running mean Trm | `(1−α)(Tod−1 + αTod−2 + α²Tod−3 + …)` with α = 0.8, evaluated recursively. The recursion is seeded with the BS EN 15251 7-day approximation. Design years wrap round from December; each observed year uses the preceding December. |
| Comfort temperature | `Tcomf = 0.33·Trm + 18.8`. Upper limit `Tmax = Tcomf + 2/3/4 K` for Category I/II/III (TM52/TM59 use II). |
| WCDH | `Σ max(0, T − Tcomf)²` over the season: the adaptive-comfort form of the metric (Eames 2016). |
| TWCDH | `Σ max(0, T − (Tcomf + ΔT_region))²` |
| SWCDH | `Σ max(0, T − T_static)²`, where `T_static` is the regional 93rd-centile temperature. CIBSE ranks DSY1 (a 1-in-7 year) by SWCDH; DSY2 has the most intense heat event and DSY3 the longest. |
| Warm events | Runs of days with daily WCDH > 0, characterised by duration, severity (total WCDH) and intensity (peak daily WCDH). These separate DSY2 (short and intense) from DSY3 (long). |
| TM52 analogues | Hours with ΔT ≥ 1 K above Tmax, the maximum daily weighted exceedance, and the maximum ΔT. All are computed on outdoor air. |
| Night-time (TM59:2026 analogue) | Mean outdoor temperature of each night, 22:00–07:00 (file time), labelled by the evening it starts. Reported as the warmest nightly mean, and the number of nights with a mean at or above a threshold. The threshold defaults to TM59:2026's 27 °C bedroom limit, where bedrooms are allowed no more than 4 such nights from May to September. |
| Seasonal climate | Mean air temperature for winter (Dec–Feb, using the analysis year's own December), spring, summer and autumn, for energy demand; and the summer (Jun–Aug) mean daily maximum, for overheating risk. |
| Other | Peak temperature, mean daily max and min, season mean, April–September mean (the original TM49 ranking), hot days, warm nights and CDH above 22 °C. |

The default season is the TM59 assessment period, 1 May to 30 September. TM49's April–September or a custom range can be chosen instead.

**Regional thresholds.** Unless you set them in the sidebar, `T_static` and `ΔT_region` are derived as 93rd
centiles (of T, and of T − Tcomf) from the location's current-climate TRY, or its DSY1 if there is no TRY. If the
library has neither, they come from the observed record. The same values are used for every file in a comparison.
If you have the TM49 regional values for your site, enter them in the sidebar.

**Similarity.** Skewed metrics (degree hours) are compared on a log scale. Each metric is then divided by its
spread across the compared files, and the distance is the weighted RMS difference. The score is `100·e^(−d)`.

## Data sources

| Source | Notes |
|---|---|
| Open-Meteo historical API (ERA5 / ERA5-Land) | Free with no key, 1940 onwards. Reanalysis smooths peaks and urban heat islands. |
| Meteostat (data.meteostat.net) | Free with no key. Station observations (ISD, METAR and others). Model gap-fill is dropped by default. |
| NOAA ISD global-hourly | Free with no key. Raw SYNOP/METAR reports; the report nearest each hour is used. |
| Met Office MIDAS Open (CEDA) | The authoritative UK record. Needs a CEDA access token (`CEDA_TOKEN`); alternatively, upload the yearly BADC-CSV files. |

Downloads are cached in `data/cache`. The library is stored in `data/library.sqlite`. Both paths are git-ignored,
and `SUMMERS_DSY_DATA` relocates them. CIBSE files are licensed, so they stay on your machine.

## Layout

```
src/summers_dsy/
  comfort.py    running mean, comfort temperature
  metrics.py    WCDH/TWCDH/SWCDH, warm events, season statistics
  compare.py    thresholds, similarity ranking, return periods
  library.py    SQLite store of reference files
  charts.py     Plotly figures
  io/           EPW, CSV and MIDAS parsers; CIBSE file-name metadata
  sources/      Open-Meteo, Meteostat, NOAA ISD and MIDAS fetchers; UK station list
  synthetic.py  synthetic weather for tests and the demo
app/streamlit_app.py
```

## References

* CIBSE TM52 (2013), *The limits of thermal comfort: avoiding overheating in European buildings*.
* CIBSE TM59 (2017), *Design methodology for the assessment of overheating risk in homes*.
* CIBSE TM49 (2014), *Design Summer Years for London*.
* Eames, M. (2016). An update of the UK's design summer years: probabilistic design summer years for enhanced
  overheating risk analysis in building design. *BSERT* 37(5).
* Liu, C., Kershaw, T., Eames, M. E. & Coley, D. A. (2016). Future probabilistic hot summer years for overheating
  risk assessments. *Building and Environment* 105.
* BS EN 15251:2007 / BS EN 16798-1:2019.

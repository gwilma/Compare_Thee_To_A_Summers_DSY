"""Build the Excel version of Compare thee to a summer's DSY.

Every calculation in the workbook is a live Excel formula. This script only lays
out the formulas and formatting and fills in demo data (synthetic, not CIBSE).

    python excel/build_workbook.py            # writes excel/summers_dsy.xlsx

The formulas mirror summers_dsy.metrics / summers_dsy.compare. excel/verify_workbook.py
checks the recalculated workbook against the Python package.
"""

from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.comments import Comment
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from summers_dsy.compare import LOG_METRICS, resolve_thresholds  # noqa: E402
from summers_dsy.metrics import DEFAULT_COMPARISON_METRICS, METRIC_INFO, AnalysisConfig, analyse  # noqa: E402
from summers_dsy.model import WeatherSeries  # noqa: E402
from summers_dsy.sources.meteostat import ATTRIBUTION, bundled_stations  # noqa: E402
from summers_dsy.synthetic import WarmSpell, demo_library, synthetic_year  # noqa: E402

OUT = Path(__file__).resolve().parent / "summers_dsy.xlsx"

N_HOURS = int(__import__("os").environ.get("WB_HOURS", 10_000))  # hourly rows available (a year plus a preceding December fits)
N_DAYS = 420
N_YEARS = 100  # rows in the Years table (one per analysed station-year)
N_REFS = 120  # reference files listed for one location on 'Years vs DSY'
N_LIB = 2500  # the full CIBSE 2025 set (28 zones × 4 file types × 18 scenarios) is 2,016 files
H0, H1 = 2, N_HOURS + 1  # hourly data rows
D0, D1 = 2, N_DAYS + 1  # daily rows
L0, L1 = 6, N_LIB + 5  # library rows

METRICS = [k for k in METRIC_INFO]  # 18 metrics, same order as the app

# ------------------------------------------------------------------ styling

FONT = "Arial"
INK, INK2, MUTED = "0B0B0B", "52514E", "898781"
BLUE, ORANGE, AQUA = "2A78D6", "EB6834", "1BAF7A"
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
HEAD_FILL = PatternFill("solid", fgColor="F3F2EE")
CALC_FILL = PatternFill("solid", fgColor="FCFCFB")
THIN = Side(style="thin", color="E1E0D9")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)


def f(size=10, bold=False, color=INK, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = f(16, True)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = f(10, color=INK2)


def header_row(ws, row, labels, start_col=1):
    for i, text in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=text)
        c.font = f(10, True)
        c.fill = HEAD_FILL
        c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")


def mark_input(cell):
    cell.fill = INPUT_FILL
    cell.font = f(10, color="0000FF")
    cell.border = BOX


def name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


# ------------------------------------------------------------------ demo data


def demo_inputs():
    """Synthetic observed year and a synthetic reference set with metrics from the Python package."""
    temps = synthetic_year(spells=[WarmSpell(208, 5, 8.5), WarmSpell(232, 9, 4.0)], warming=0.9, seed=7, year=2020)
    observed = WeatherSeries("2020", temps.to_frame("dry_bulb"), "synthetic", {"typical_year": False})
    lib = demo_library()
    refs = {}
    for fname, (series, meta) in lib.items():
        series.meta.update(meta)
        series.meta["typical_year"] = True
        refs[fname] = series
    baseline = next(s for s in refs.values() if s.meta["kind"] == "TRY")
    thr = resolve_thresholds(AnalysisConfig(), baseline, observed)
    # The workbook shows thresholds to 3 d.p.; compute the demo library with exactly those values.
    thr.static_threshold, thr.twcdh_offset = round(thr.static_threshold, 3), round(thr.twcdh_offset, 3)
    cfg = AnalysisConfig(static_threshold=thr.static_threshold, twcdh_offset=thr.twcdh_offset)
    rows = [(s.meta, analyse(s, cfg).metrics) for s in refs.values()]
    return observed, thr, rows


# ------------------------------------------------------------------ sheets


def build():
    observed, thr, lib_rows = demo_inputs()
    wb = Workbook()
    wb._named_styles["Normal"].font = Font(name=FONT, size=10)
    readme = wb.active
    readme.title = "Read me"
    dash = wb.create_sheet("Dashboard")
    st = wb.create_sheet("Settings")
    hr = wb.create_sheet("Hourly")
    dy = wb.create_sheet("Daily")
    sm = wb.create_sheet("Summary")
    lb = wb.create_sheet("Library")
    cp = wb.create_sheet("Compare")
    cc = wb.create_sheet("Compare calc")
    cd = wb.create_sheet("Chart data")
    lg = wb.create_sheet("Import log")
    ms_ws = wb.create_sheet("Meteostat", 2)
    yv = wb.create_sheet("Years vs DSY", 3)
    yr_ws = wb.create_sheet("Years")
    stn = wb.create_sheet("Stations")

    # ---------------------------------------------------------- Settings
    title(st, "Settings", "Yellow cells with blue text are inputs. Everything else is calculated.")
    st.column_dimensions["A"].width = 38
    st.column_dimensions["B"].width = 22
    st.column_dimensions["C"].width = 90
    settings = [
        # row, label, value, name, note, is_input
        (4, "Site name", "Demo site (synthetic)", "site_name", "Used in the headline.", True),
        (5, "Typical/design year file? (wrap running mean)", "No",
         "cyclic", "Yes for TRY/DSY files: the running mean at 1 January is seeded from late December of the same file. "
         "No for observed years (paste the preceding December too if you have it).", True),
        (6, "Nominal year for rows with no Year", 2001, "nominal_year", "TRY/DSY files are placed on this non-leap year.", True),
        (7, "Hour convention", "1-24 (hour ending)", "hour_conv", "EPW and CIBSE files use 1-24 (hour ending). Choose 0-23 for hour-beginning data.", True),
        (9, "Season start month", 5, "s_month", "TM59 assessment period is 1 May to 30 September (TM49 used 1 April).", True),
        (10, "Season start day", 1, "s_day", "", True),
        (11, "Season end month", 9, "e_month", "Season must not wrap past 31 December.", True),
        (12, "Season end day", 30, "e_day", "", True),
        (13, "Season start (MMDD)", "=s_month*100+s_day", "season_lo", "Calculated.", False),
        (14, "Season end (MMDD)", "=e_month*100+e_day", "season_hi", "Calculated.", False),
        (16, "Running-mean weight α", 0.8, "alpha", "CIBSE TM52 / BS EN 15251 recommend 0.8.", True),
        (17, "Running-mean method", "Exponential", "trm_method",
         "Exponential: Trm(n) = (1-α)·Tod(n-1) + α·Trm(n-1), seeded with the 7-day approximation. "
         "'7-day approximation' uses BS EN 15251's (Tod-1 + 0.8Tod-2 + … + 0.2Tod-7)/3.8 every day.", True),
        (18, "Comfort category", "II", "category", "Tmax = Tcomf + 2 K (I), 3 K (II, TM52/TM59 default) or 4 K (III).", True),
        (19, "Category offset (K)", '=IF(category="I",2,IF(category="III",4,3))', "cat_offset", "Calculated.", False),
        (20, "Clamp Trm to 10–30 °C?", "No", "clamp_trm", "BS EN 15251 gives the adaptive limit for 10 < Trm < 30 °C.", True),
        (22, "SWCDH threshold override (°C)", round(thr.static_threshold, 3), "swcdh_override",
         "Leave blank to use the value derived from this file (row 24). For comparisons, use the same value for every file: "
         "the 93rd centile from the location's current-climate TRY. The demo value came from the demo TRY.", True),
        (23, "TWCDH offset override (K)", round(thr.twcdh_offset, 3), "twcdh_override",
         "Leave blank to use row 25. Demo value from the demo TRY.", True),
        (24, "SWCDH threshold derived from this file (°C)", f'=IF(COUNT(Hourly!$R$2:$R${H1})=0,"",PERCENTILE(Hourly!$R$2:$R${H1},0.93))', "swcdh_auto",
         "93rd centile of in-season hourly dry-bulb (Eames 2016 regional threshold convention).", False),
        (25, "TWCDH offset derived from this file (K)", f'=IF(COUNT(Hourly!$S$2:$S${H1})=0,"",PERCENTILE(Hourly!$S$2:$S${H1},0.93))', "twcdh_auto",
         "93rd centile of in-season (T − Tcomf).", False),
        (26, "SWCDH threshold used (°C)", '=IF(swcdh_override<>"",swcdh_override,N(swcdh_auto))', "swcdh_thr", "Calculated.", False),
        (27, "TWCDH offset used (K)", '=IF(twcdh_override<>"",twcdh_override,N(twcdh_auto))', "twcdh_off", "Calculated.", False),
        (29, "Hot day: daily max ≥ (°C)", 28, "hot_thr", "", True),
        (30, "Warm night: daily min ≥ (°C)", 16, "night_thr", "", True),
        (31, "Warm-event day: daily WCDH > (K²h)", 0, "event_min", "A warm event is a run of consecutive in-season days above this.", True),
        (33, "First day in data", f"=INT(MIN(Hourly!$F$2:$F${H1}))", "first_day", "Calculated.", False),
        (34, "Number of days in data", f"=INT(MAX(Hourly!$F$2:$F${H1}))-first_day+1", "n_days",
         f"Calculated. Daily sheet holds up to {N_DAYS} days.", False),
        (35, "Analysis year", f"=YEAR(MAX(Hourly!$F$2:$F${H1}))", "analysis_year", "Year of the last day of data.", False),
        (36, "Rows of hourly data", f"=MAX(1,COUNT(Hourly!$F$2:$F${H1}))", "n_rows",
         "Calculated. Hourly rows must be in time order with no gaps between rows.", False),
        (41, "Folder to import (ImportWeatherFolder macro)", None, "import_folder",
         "Optional. Leave blank to be asked for a folder when the macro runs. All sub-folders are included.", True),
        (42, "Thresholds for imported files", "Each location's TRY", "import_thr_mode",
         "'Each location's TRY': the macro derives the SWCDH threshold and TWCDH offset for each location from its "
         "earliest-period TRY (else DSY1) and uses them for all of that location's files. 'Settings overrides': "
         "every file uses rows 22–23.", True),
        (44, "Night starts (hour, 0–23)", 22, "night_start",
         "TM59 treats 22:00–07:00 as bedroom night-time. Hours are file time; the window must cross midnight.", True),
        (45, "Night ends (hour, 0–23)", 7, "night_end", "", True),
        (46, "Night mean threshold (°C)", 27, "night_mean_thr",
         "TM59:2026: the mean bedroom temperature at night must stay below 27 °C, with no more than 4 exceedance nights "
         "May–September. Applied here to outdoor air; a lower value is often more telling for UK nights.", True),
        (47, "Minimum valid hours for a night mean", 7, "night_min_hours", "", True),
        (48, "Longest gap in daily means filled (days)", 3, "max_fill_days",
         "Days with under 18 hours of data have no daily mean. For the running mean, gaps of up to this many days are "
         "interpolated; after a longer gap the running mean restarts once 7 complete days are available.", True),
    ]
    for r, label, value, nm, note, is_input in settings:
        st.cell(row=r, column=1, value=label).font = f()
        c = st.cell(row=r, column=2, value=value)
        if is_input:
            mark_input(c)
        else:
            c.font = f()
        st.cell(row=r, column=3, value=note).font = f(9, color=INK2)
        st.cell(row=r, column=3).alignment = Alignment(wrap_text=True, vertical="top")
        name(wb, nm, f"Settings!$B${r}")
    for r in (13, 14, 19, 22, 23, 24, 25, 26, 27):
        st.cell(row=r, column=2).number_format = "0.00"
    st["B33"].number_format = "dd mmm yyyy"
    st["A38"] = "BS EN 15251 running-mean weights (Tod-1 … Tod-7)"
    st["A38"].font = f()
    for i, w in enumerate([1, 0.8, 0.6, 0.5, 0.4, 0.3, 0.2]):
        c = st.cell(row=38, column=2 + i, value=w)
        c.font = f()
    st["B39"] = "Source: BS EN 15251:2007 Annex A.2; the weights sum to 3.8."
    st["B39"].font = f(9, color=INK2)
    name(wb, "rm_weights", "Settings!$B$38:$H$38")
    st.freeze_panes = "A4"

    for rng, options in (
        ("B5", '"Yes,No"'), ("B7", '"1-24 (hour ending),0-23 (hour beginning)"'),
        ("B17", '"Exponential,7-day approximation"'), ("B18", '"I,II,III"'), ("B20", '"Yes,No"'),
        ("B42", '"Each location\'s TRY,Settings overrides"'),
    ):
        dv = DataValidation(type="list", formula1=options, allow_blank=False)
        st.add_data_validation(dv)
        dv.add(rng)

    # ---------------------------------------------------------- Hourly
    hr_cols = [
        ("Year", 7), ("Month", 7), ("Day", 6), ("Hour", 6), ("Dry bulb (°C)", 10),
        ("Date-time", 16), ("Date", 11), ("In season", 8), ("Day index", 8), ("Tcomf (°C)", 9), ("Tmax (°C)", 9),
        ("WCDH (K²h)", 10), ("TWCDH (K²h)", 10), ("SWCDH (K²h)", 10), ("ΔT above Tmax, rounded (K)", 11),
        ("Weighted exceedance (K·h)", 11), ("CDH above 22 °C (K·h)", 11), ("In-season T (°C)", 10),
        ("In-season T − Tcomf (K)", 11), ("Month no.", 7), ("Night of (date)", 11), ("Year no.", 7),
    ]
    header_row(hr, 1, [c[0] for c in hr_cols])
    for i, (_, w) in enumerate(hr_cols, start=1):
        hr.column_dimensions[col(i)].width = w
    hr.row_dimensions[1].height = 42
    hr.freeze_panes = "F2"
    hr["A1"].comment = Comment(
        "Paste hourly data into columns A–E (yellow). EPW files: columns 1–4 are Year, Month, Day, Hour and "
        "column 7 is dry bulb. Leave Year blank for typical/design years to use the nominal year in Settings. "
        "Delete any 29 February rows from TRY/DSY files.", "summers_dsy", width=320, height=140)

    dt = observed.dry_bulb
    for i in range(N_HOURS):
        r = H0 + i
        if i < len(dt):
            ts = dt.index[i]
            hr.cell(row=r, column=1, value=ts.year)
            hr.cell(row=r, column=2, value=ts.month)
            hr.cell(row=r, column=3, value=ts.day)
            hr.cell(row=r, column=4, value=ts.hour + 1)
            hr.cell(row=r, column=5, value=round(float(dt.iloc[i]), 2))
        hr[f"F{r}"] = (f'=IF($E{r}="","",DATE(IF($A{r}="",nominal_year,$A{r}),$B{r},$C{r})'
                       f'+($D{r}-IF(LEFT(hour_conv,1)="1",1,0))/24)')
        hr[f"G{r}"] = f'=IF(F{r}="","",INT(F{r}))'
        hr[f"H{r}"] = (f'=IF(G{r}="","",IF(AND(MONTH(G{r})*100+DAY(G{r})>=season_lo,'
                       f'MONTH(G{r})*100+DAY(G{r})<=season_hi),1,0))')
        hr[f"I{r}"] = f'=IF(G{r}="","",G{r}-first_day+1)'
        hr[f"J{r}"] = f'=IF(I{r}="","",INDEX(Daily!$J${D0}:$J${D1},I{r}))'
        hr[f"K{r}"] = f'=IF(J{r}="","",J{r}+cat_offset)'
        hr[f"L{r}"] = f'=IF(J{r}="","",MAX(0,E{r}-J{r})^2)'
        hr[f"M{r}"] = f'=IF(J{r}="","",MAX(0,E{r}-(J{r}+twcdh_off))^2)'
        hr[f"N{r}"] = f'=IF(E{r}="","",MAX(0,E{r}-swcdh_thr)^2)'
        hr[f"O{r}"] = f'=IF(K{r}="","",ROUND(E{r}-K{r},0))'
        hr[f"P{r}"] = f'=IF(O{r}="","",IF(O{r}>=1,O{r},0))'
        hr[f"Q{r}"] = f'=IF(E{r}="","",MAX(0,E{r}-22))'
        hr[f"R{r}"] = f'=IF(H{r}=1,E{r},"")'
        hr[f"S{r}"] = f'=IF(AND(H{r}=1,J{r}<>""),E{r}-J{r},"")'
        hr[f"T{r}"] = f'=IF(G{r}="","",MONTH(G{r}))'
        hr[f"U{r}"] = (f'=IF(G{r}="","",IF(HOUR(F{r})>=night_start,G{r},IF(HOUR(F{r})<night_end,G{r}-1,"")))')
        hr[f"V{r}"] = f'=IF(G{r}="","",YEAR(G{r}))'
    for c_idx in range(1, 6):
        for r in range(H0, H1 + 1):
            cell = hr.cell(row=r, column=c_idx)
            cell.fill = INPUT_FILL
            cell.font = f(10, color="0000FF")
    for r in range(H0, H1 + 1):
        hr[f"F{r}"].number_format = "dd mmm yyyy hh:mm"
        hr[f"G{r}"].number_format = "dd mmm yyyy"
        for c_ in "JKLMNQRS":
            hr[f"{c_}{r}"].number_format = "0.00"

    # ---------------------------------------------------------- Daily
    dy_cols = [
        ("Day no.", 6), ("Date", 11), ("Hours", 6), ("Daily mean (°C)", 9), ("Daily min (°C)", 9), ("Daily max (°C)", 9),
        ("Range (K)", 8), ("7-day approx. Trm (°C)", 10), ("Trm (°C)", 9), ("Tcomf (°C)", 9), ("Tmax (°C)", 9),
        ("In season", 7), ("WCDH (K²h)", 9), ("TWCDH (K²h)", 9), ("SWCDH (K²h)", 9), ("Daily weighted exceedance (K·h)", 11),
        ("Warm-event day", 8), ("Event length so far (days)", 9), ("Event WCDH so far (K²h)", 10), ("Event start", 7),
        ("Hot day", 6), ("Warm night", 7), ("", 2),
        ("Tod-1", 7), ("Tod-2", 7), ("Tod-3", 7), ("Tod-4", 7), ("Tod-5", 7), ("Tod-6", 7), ("Tod-7", 7),
        ("", 2), ("First hourly row", 8), ("Last hourly row", 8),
        ("Night hours", 7), ("Night mean (°C)", 9), ("Night ≥ threshold", 8), ("Month no.", 7), ("Year no.", 7),
        ("Previous day with a mean", 8), ("Next day with a mean", 8), ("Daily mean for running mean (°C)", 10),
    ]
    header_row(dy, 1, [c[0] for c in dy_cols])
    for i, (_, w) in enumerate(dy_cols, start=1):
        dy.column_dimensions[col(i)].width = w
    dy.row_dimensions[1].height = 56
    dy.freeze_panes = "C2"
    HR = lambda c: f"Hourly!${c}${H0}:${c}${H1}"  # noqa: E731
    # Only the filled rows of the (sorted) Date column, so MATCH's binary search sees numbers only.
    HDATES = f"INDEX({HR('G')},1):INDEX({HR('G')},n_rows)"
    for i in range(N_DAYS):
        r = D0 + i
        dy[f"A{r}"] = i + 1
        dy[f"B{r}"] = f'=IF(A{r}<=n_days,first_day+A{r}-1,"")'
        # Rows of this day's hours: binary search (MATCH type 1) on the sorted Date column.
        dy[f"AF{r}"] = f'=IF(B{r}="","",IFERROR(MATCH(B{r}-0.5,{HDATES},1),0)+1)'
        dy[f"AG{r}"] = f'=IF(B{r}="","",IFERROR(MATCH(B{r},{HDATES},1),0))'
        dy[f"C{r}"] = f'=IF(B{r}="","",MAX(0,AG{r}-AF{r}+1))'
        day = lambda c_: f"INDEX({HR(c_)},AF{r}):INDEX({HR(c_)},AG{r})"  # noqa: E731
        dy[f"D{r}"] = f'=IF(B{r}="","",IF(C{r}>=18,AVERAGE({day("E")}),""))'
        dy[f"E{r}"] = f'=IF(OR(B{r}="",N(C{r})=0),"",MIN({day("E")}))'
        # The night starting this evening ends next morning: search this day's and the next day's rows.
        night = lambda c_: f"INDEX({HR(c_)},AF{r}):INDEX({HR(c_)},MAX(AG{r},N(AG{r + 1})))"  # noqa: E731
        dy[f"AH{r}"] = f'=IF(OR(B{r}="",N(AF{r})=0),"",COUNTIF({night("U")},B{r}))'
        dy[f"AI{r}"] = f'=IF(N(AH{r})<night_min_hours,"",AVERAGEIFS({night("E")},{night("U")},B{r}))'
        dy[f"AJ{r}"] = f'=IF(AND(L{r}=1,ISNUMBER(AI{r})),IF(AI{r}>=night_mean_thr,1,0),0)'
        dy[f"AK{r}"] = f'=IF(B{r}="","",MONTH(B{r}))'
        dy[f"AL{r}"] = f'=IF(B{r}="","",YEAR(B{r}))'
        # Short gaps in the daily means are interpolated for the running mean (as in the Python package).
        dy[f"AM{r}"] = f'=IF(B{r}="","",IF(ISNUMBER(D{r}),A{r},N(AM{r - 1})))'
        dy[f"AN{r}"] = f'=IF(B{r}="","",IF(ISNUMBER(D{r}),A{r},N(AN{r + 1})))'
        dmean = f"$D${D0}:$D${D1}"
        dy[f"AO{r}"] = (f'=IF(B{r}="","",IF(ISNUMBER(D{r}),D{r},IF(AND(N(AM{r})>0,N(AN{r})>0,N(AN{r})-N(AM{r})-1<=max_fill_days),'
                        f'INDEX({dmean},AM{r})+(INDEX({dmean},AN{r})-INDEX({dmean},AM{r}))*(A{r}-AM{r})/(AN{r}-AM{r}),"")))')
        dy[f"AO{r}"].number_format = "0.00"
        dy[f"AI{r}"].number_format = "0.00"
        dy[f"F{r}"] = f'=IF(OR(B{r}="",N(C{r})=0),"",MAX({day("E")}))'
        dy[f"G{r}"] = f'=IF(E{r}="","",F{r}-E{r})'
        dy[f"H{r}"] = f'=IF(B{r}="","",IF(COUNT(X{r}:AD{r})=7,SUMPRODUCT(X{r}:AD{r},rm_weights)/SUM(rm_weights),""))'
        dy[f"I{r}"] = (f'=IF(B{r}="","",IF(trm_method="7-day approximation",H{r},'
                       f'IF(AND(ISNUMBER(I{r - 1}),ISNUMBER(AO{r - 1})),(1-alpha)*AO{r - 1}+alpha*I{r - 1},H{r})))')
        dy[f"J{r}"] = f'=IF(ISNUMBER(I{r}),0.33*IF(clamp_trm="Yes",MIN(MAX(I{r},10),30),I{r})+18.8,"")'
        dy[f"K{r}"] = f'=IF(J{r}="","",J{r}+cat_offset)'
        dy[f"L{r}"] = (f'=IF(B{r}="","",IF(AND(MONTH(B{r})*100+DAY(B{r})>=season_lo,'
                       f'MONTH(B{r})*100+DAY(B{r})<=season_hi),1,0))')
        dy[f"M{r}"] = f'=IF(N(C{r})=0,IF(B{r}="","",0),SUM({day("L")}))'
        dy[f"N{r}"] = f'=IF(N(C{r})=0,IF(B{r}="","",0),SUM({day("M")}))'
        dy[f"O{r}"] = f'=IF(N(C{r})=0,IF(B{r}="","",0),SUM({day("N")}))'
        dy[f"P{r}"] = f'=IF(N(C{r})=0,IF(B{r}="","",0),SUM({day("P")}))'
        dy[f"Q{r}"] = f'=IF(L{r}=1,IF(M{r}>event_min,1,0),0)'
        dy[f"R{r}"] = f'=IF(Q{r}=1,N(R{r - 1})+1,0)'
        dy[f"S{r}"] = f'=IF(Q{r}=1,N(S{r - 1})+M{r},0)'
        dy[f"T{r}"] = f'=IF(R{r}=1,1,0)'
        dy[f"U{r}"] = f'=IF(AND(L{r}=1,ISNUMBER(F{r})),IF(F{r}>=hot_thr,1,0),0)'
        dy[f"V{r}"] = f'=IF(AND(L{r}=1,ISNUMBER(E{r})),IF(E{r}>=night_thr,1,0),0)'
        for k in range(1, 8):
            c_ = col(23 + k)  # X..AD
            dy[f"{c_}{r}"] = (f'=IF(B{r}="","",IF(A{r}-{k}>=1,INDEX($AO${D0}:$AO${D1},A{r}-{k}),'
                              f'IF(cyclic="Yes",INDEX($AO${D0}:$AO${D1},A{r}-{k}+n_days),"")))')
        dy[f"B{r}"].number_format = "dd mmm yyyy"
        for c_ in "DEFGHIJKMNOPS":
            dy[f"{c_}{r}"].number_format = "0.00"
        for k in range(24, 31):
            dy.cell(row=r, column=k).number_format = "0.00"
    dy["X1"].comment = Comment("Daily means 1–7 days earlier, used for the BS EN 15251 7-day approximation. "
                               "For typical years (Settings: Yes) the start of the year wraps to December.", "summers_dsy")

    # ---------------------------------------------------------- Summary
    DY = lambda c: f"Daily!${c}${D0}:${c}${D1}"  # noqa: E731
    metric_formula = {
        "wcdh": f"=SUMIFS({HR('L')},{HR('H')},1)",
        "twcdh": f"=SUMIFS({HR('M')},{HR('H')},1)",
        "swcdh": f"=SUMIFS({HR('N')},{HR('H')},1)",
        "peak_daily_wcdh": f"=_xlfn.MAXIFS({DY('M')},{DY('L')},1)",
        "max_event_severity": f"=MAX({DY('S')})",
        "max_event_duration": f"=MAX({DY('R')})",
        "n_events": f"=SUM({DY('T')})",
        "hours_above_upper": f'=COUNTIFS({HR("O")},">=1",{HR("H")},1)',
        "max_delta_t": f"=MAX(0,_xlfn.MAXIFS({HR('O')},{HR('H')},1))",
        "max_daily_weighted_exceedance": f"=_xlfn.MAXIFS({DY('P')},{DY('L')},1)",
        "t_max": f"=_xlfn.MAXIFS({HR('E')},{HR('H')},1)",
        "mean_daily_max": f"=AVERAGEIFS({DY('F')},{DY('L')},1)",
        "mean_daily_min": f"=AVERAGEIFS({DY('E')},{DY('L')},1)",
        "season_mean": f"=AVERAGEIFS({HR('E')},{HR('H')},1)",
        "apr_sep_mean": f'=AVERAGEIFS({HR("E")},{HR("T")},">=4",{HR("T")},"<=9")',
        "hot_days": f"=SUM({DY('U')})",
        "warm_nights": f"=SUM({DY('V')})",
        "cdh_22": f"=SUMIFS({HR('Q')},{HR('H')},1)",
        "night_max_mean": f"=_xlfn.MAXIFS({DY('AI')},{DY('L')},1)",
        "nights_above": f"=SUM({DY('AJ')})",
        # Seasonal means use the analysis year only (the data may start with the previous December).
        "mean_t_djf": (f"=(SUMIFS({HR('E')},{HR('T')},12,{HR('V')},analysis_year)+SUMIFS({HR('E')},{HR('T')},1,{HR('V')},analysis_year)"
                       f"+SUMIFS({HR('E')},{HR('T')},2,{HR('V')},analysis_year))/(COUNTIFS({HR('T')},12,{HR('V')},analysis_year)"
                       f"+COUNTIFS({HR('T')},1,{HR('V')},analysis_year)+COUNTIFS({HR('T')},2,{HR('V')},analysis_year))"),
        "mean_t_mam": f'=AVERAGEIFS({HR("E")},{HR("T")},">=3",{HR("T")},"<=5",{HR("V")},analysis_year)',
        "mean_t_jja": f'=AVERAGEIFS({HR("E")},{HR("T")},">=6",{HR("T")},"<=8",{HR("V")},analysis_year)',
        "mean_t_son": f'=AVERAGEIFS({HR("E")},{HR("T")},">=9",{HR("T")},"<=11",{HR("V")},analysis_year)',
        "jja_mean_daily_max": f'=AVERAGEIFS({DY("F")},{DY("AK")},">=6",{DY("AK")},"<=8",{DY("AL")},analysis_year)',
    }
    title(sm, "Summary: overheating metrics for the season", None)
    sm["A2"] = '="Site: "&site_name&" · season "&TEXT(DATE(analysis_year,s_month,s_day),"d mmm")&" – "&TEXT(DATE(analysis_year,e_month,e_day),"d mmm yyyy")'
    sm["A2"].font = f(10, color=INK2)
    header_row(sm, 4, ["Metric", "Value", "Unit", "Definition"])
    for c_, w in zip("ABCD", (30, 14, 8, 95)):
        sm.column_dimensions[c_].width = w
    summary_row = {}
    for i, key in enumerate(METRICS):
        r = 5 + i
        label, unit, desc = METRIC_INFO[key]
        sm[f"A{r}"] = label
        sm[f"B{r}"] = f'=IFERROR({metric_formula[key][1:]},"")'
        sm[f"C{r}"] = unit
        sm[f"D{r}"] = desc
        for c_ in "ACD":
            sm[f"{c_}{r}"].font = f(10, color=INK if c_ == "A" else INK2)
        sm[f"B{r}"].font = f(10, True)
        sm[f"B{r}"].number_format = "#,##0" if unit in ("K²h", "K·h", "h", "days", "count", "nights") else "0.0"
        summary_row[key] = f"Summary!$B${r}"
    r = 5 + len(METRICS) + 1
    sm[f"A{r}"] = "Thresholds used"
    sm[f"A{r}"].font = f(10, True)
    sm[f"A{r + 1}"], sm[f"B{r + 1}"], sm[f"C{r + 1}"] = "SWCDH threshold", "=swcdh_thr", "°C"
    sm[f"A{r + 2}"], sm[f"B{r + 2}"], sm[f"C{r + 2}"] = "TWCDH offset above Tcomf", "=twcdh_off", "K"
    sm[f"A{r + 3}"] = "Season coverage"
    sm[f"B{r + 3}"] = (f"=COUNTIF({HR('H')},1)/((DATE(analysis_year,e_month,e_day)-DATE(analysis_year,s_month,s_day)+1)*24)")
    sm[f"C{r + 3}"] = "share of season hours with data"
    sm[f"B{r + 3}"].number_format = "0.0%"
    name(wb, "season_coverage", f"Summary!$B${r + 3}")
    for rr in (r + 1, r + 2, r + 3):
        if rr < r + 3:
            sm[f"B{rr}"].number_format = "0.00"
        for c_ in "ABC":
            sm[f"{c_}{rr}"].font = f()
    name(wb, "metric_labels", f"Summary!$A$5:$A${4 + len(METRICS)}")

    copy_r = r + 5
    sm[f"A{copy_r}"] = "Row to copy into the Library (select it, Copy, then Paste Special → Values on a Library row, from column H)"
    sm[f"A{copy_r}"].font = f(10, True)
    header_row(sm, copy_r + 1, [METRIC_INFO[k][0] for k in METRICS])
    for i, key in enumerate(METRICS):
        c = sm.cell(row=copy_r + 2, column=1 + i, value=f"={summary_row[key]}")
        c.font = f()
        c.number_format = "0.00"
    for i in range(len(METRICS)):
        sm.column_dimensions[col(1 + i)].width = max(sm.column_dimensions[col(1 + i)].width or 0, 12)

    # ---------------------------------------------------------- Library
    title(lb, "Library of reference weather files",
          "One row per CIBSE TRY/DSY file. Fill it with the ImportWeatherFolder macro (a folder and all its sub-folders), "
          "or by pasting a Summary row as values. Include = 1 marks the files to compare.")
    lb["A3"] = ("Demo rows are synthetic (not CIBSE data), calculated by the Python app with the Settings demo thresholds. "
                "Delete them before adding your own files.")
    lb["A3"].font = f(9, color="C00000", italic=True)
    lib_head = ["Include (1/0)", "Location", "File type", "Period", "Emissions", "Percentile", "Label"] + [
        f"{METRIC_INFO[k][0]} ({METRIC_INFO[k][1]})" for k in METRICS] + [
        "SWCDH threshold used (°C)", "TWCDH offset used (K)", "Source file"]
    c_thr_s, c_thr_t, c_src = 8 + len(METRICS), 9 + len(METRICS), 10 + len(METRICS)  # Z, AA, AB
    header_row(lb, 5, lib_head)
    lb.row_dimensions[5].height = 44
    for i, w in enumerate([9, 20, 9, 10, 11, 10, 44] + [12] * len(METRICS) + [12, 12, 60], start=1):
        lb.column_dimensions[col(i)].width = w
    lb.freeze_panes = "H6"
    for i in range(N_LIB):
        r = L0 + i
        for c_idx in list(range(1, 7)) + list(range(8, 8 + len(METRICS))) + [c_thr_s, c_thr_t, c_src]:
            mark_input(lb.cell(row=r, column=c_idx))
        lb[f"G{r}"] = (f'=IF(C{r}="","",C{r}&IF(D{r}="",""," · "&IF(D{r}="Baseline","current climate",D{r}))'
                       f'&IF(E{r}="",""," · "&E{r}&" emissions")&IF(F{r}="",""," · "&F{r}&"th percentile"))')
        lb[f"G{r}"].font = f()
        for c_idx in range(8, 8 + len(METRICS)):
            lb.cell(row=r, column=c_idx).number_format = "#,##0.0"
        if i < len(lib_rows):
            meta, m = lib_rows[i]
            lb[f"A{r}"] = 1
            lb[f"B{r}"] = meta["location"]
            lb[f"C{r}"] = meta["kind"]
            lb[f"D{r}"] = meta["period"]
            lb[f"E{r}"] = meta["emissions"]
            lb[f"F{r}"] = meta["percentile"]
            for j, key in enumerate(METRICS):
                lb.cell(row=r, column=8 + j, value=float(m[key]))
            lb.cell(row=r, column=c_thr_s, value=thr.static_threshold)
            lb.cell(row=r, column=c_thr_t, value=thr.twcdh_offset)
            lb.cell(row=r, column=c_src, value="synthetic demo data")
    for rng, options in ((f"A{L0}:A{L1}", '"1,0"'), (f"C{L0}:C{L1}", '"TRY,DSY1,DSY2,DSY3,Other"'),
                         (f"D{L0}:D{L1}", '"Baseline,2020s,2030s,2050s,2080s"'), (f"E{L0}:E{L1}", '"Low,Medium,High"'),
                         (f"F{L0}:F{L1}", '"10,50,90"')):
        dv = DataValidation(type="list", formula1=options, allow_blank=True)
        lb.add_data_validation(dv)
        dv.add(rng)
    name(wb, "lib_table", f"Library!$A${L0}:${col(c_src)}${L1}")

    # ---------------------------------------------------------- Compare calc
    nm = len(METRICS)
    first_m, last_m = 3, 2 + nm  # C .. T
    z_first = last_m + 2  # V
    z_last = z_first + nm - 1  # AM
    c_dist, c_score, c_key, c_label, c_loc = (col(z_last + 2), col(z_last + 3), col(z_last + 4), col(z_last + 5), col(z_last + 6))
    title(cc, "Compare calc", "Working for the similarity ranking. Skewed metrics use ln(1 + x); each metric is scaled by its "
          "population standard deviation across the included files plus the observed year.")
    labels_row, w_row, log_row, t_row, tt_row, n_row, mean_row, sd_row, sdu_row = 4, 5, 6, 7, 8, 9, 10, 11, 12
    for row_, text in ((labels_row, "Metric"), (w_row, "Weight (from Compare)"), (log_row, "Log transform (1/0)"),
                       (t_row, "Observed value"), (tt_row, "Observed, transformed"), (n_row, "n (files + observed)"),
                       (mean_row, "Mean"), (sd_row, "Spread (population SD)"), (sdu_row, "Spread used")):
        cc.cell(row=row_, column=2, value=text).font = f(10, True)
    cc.column_dimensions["A"].width = 6
    cc.column_dimensions["B"].width = 24
    R0, R1 = 15, 15 + N_LIB - 1
    for j, key in enumerate(METRICS):
        c_ = col(first_m + j)
        zc = col(z_first + j)
        lib_c = col(8 + j)
        cc.column_dimensions[c_].width = 11
        cc.column_dimensions[zc].width = 8
        cc[f"{c_}{labels_row}"] = METRIC_INFO[key][0]
        cc[f"{c_}{w_row}"] = f"=Compare!$H${13 + j}"
        cc[f"{c_}{log_row}"] = 1 if key in LOG_METRICS else 0
        cc[f"{c_}{t_row}"] = f"={summary_row[key]}"
        cc[f"{c_}{tt_row}"] = f'=IF(NOT(ISNUMBER({c_}{t_row})),0,IF({c_}{log_row}=1,LN(1+MAX(0,{c_}{t_row})),{c_}{t_row}))'
        rng = f"{c_}${R0}:{c_}${R1}"
        cc[f"{c_}{n_row}"] = f"=COUNT({rng})+1"
        cc[f"{c_}{mean_row}"] = f"=(SUM({rng})+{c_}{tt_row})/{c_}{n_row}"
        cc[f"{c_}{sd_row}"] = f"=SQRT(MAX(0,(SUMSQ({rng})+{c_}{tt_row}^2)/{c_}{n_row}-{c_}{mean_row}^2))"
        cc[f"{c_}{sdu_row}"] = f"=IF({c_}{sd_row}>0.000000001,{c_}{sd_row},1)"
        cc[f"{zc}{labels_row}"] = f"z² {METRIC_INFO[key][0]}"
        for r in range(R0, R1 + 1):
            lr = L0 + (r - R0)
            cc[f"{c_}{r}"] = (f'=IF(OR(Library!$A{lr}<>1,AND(compare_loc<>"",Library!$B{lr}<>compare_loc),Library!{lib_c}{lr}=""),"",'
                              f'IF({c_}${log_row}=1,LN(1+MAX(0,Library!{lib_c}{lr})),Library!{lib_c}{lr}))')
            cc[f"{zc}{r}"] = f'=IF({c_}{r}="","",(({c_}{r}-{c_}${tt_row})/{c_}${sdu_row})^2)'
            cc[f"{c_}{r}"].number_format = "0.000"
            cc[f"{zc}{r}"].number_format = "0.000"
        for rr in (tt_row, mean_row, sd_row, sdu_row, t_row):
            cc[f"{c_}{rr}"].number_format = "0.000"
    cc[f"{col(first_m)}{R0 - 1}"] = "Transformed library values →"
    cc[f"{col(z_first)}{R0 - 1}"] = "Squared standardised differences →"
    header_row(cc, labels_row, ["Distance", "Similarity (0–100)", "Rank key", "Label", "Location"],
               start_col=z_last + 2)
    wsum = f"SUM(${col(first_m)}${w_row}:${col(last_m)}${w_row})"
    for r in range(R0, R1 + 1):
        lr = L0 + (r - R0)
        zr = f"{col(z_first)}{r}:{col(z_last)}{r}"
        cc[f"{c_dist}{r}"] = (f'=IF(OR(Library!$A{lr}<>1,AND(compare_loc<>"",Library!$B{lr}<>compare_loc),COUNT({zr})=0,{wsum}<=0),"",'
                              f'SQRT(SUMPRODUCT({zr},${col(first_m)}${w_row}:${col(last_m)}${w_row})/{wsum}))')
        cc[f"{c_score}{r}"] = f'=IF({c_dist}{r}="","",100*EXP(-{c_dist}{r}))'
        cc[f"{c_key}{r}"] = f'=IF({c_score}{r}="","",{c_score}{r}-ROW()*0.000000001)'
        cc[f"{c_label}{r}"] = f'=IF({c_score}{r}="","",Library!$G{lr})'
        cc[f"{c_loc}{r}"] = f'=IF({c_score}{r}="","",Library!$B{lr})'
        for c_ in (c_dist, c_score):
            cc[f"{c_}{r}"].number_format = "0.00"
    cc.column_dimensions[c_label].width = 44
    cc.column_dimensions[c_loc].width = 20
    cc.freeze_panes = f"C{R0}"
    KEY = f"'Compare calc'!${c_key}${R0}:${c_key}${R1}"
    SCORE = f"'Compare calc'!${c_score}${R0}:${c_score}${R1}"
    LABEL = f"'Compare calc'!${c_label}${R0}:${c_label}${R1}"
    LOC = f"'Compare calc'!${c_loc}${R0}:${c_loc}${R1}"

    # ---------------------------------------------------------- Compare
    title(cp, "Compare with the library",
          "Ranks every included Library file by similarity to this workbook's season. Weights (yellow) choose which metrics count.")
    cp.column_dimensions["A"].width = 26
    cp["A4"] = "Result"
    cp["A4"].font = f(11, True)
    cp["B4"] = (f'=IF(COUNT(Summary!$B$5:$B${4 + len(METRICS)})<{len(METRICS)},"Load at least a full season of hourly data on the Hourly sheet.",'
                f'IF(COUNT({SCORE})=0,"Add reference files to the Library sheet and set Include to 1.",'
                f'analysis_year&" at "&site_name&" was most similar to "&INDEX({LABEL},MATCH(MAX({KEY}),{KEY},0))'
                f'&" for "&INDEX({LOC},MATCH(MAX({KEY}),{KEY},0))&" (similarity "&TEXT(MAX({SCORE}),"0")&"/100)."))')
    cp["B4"].font = f(12, True)
    header_row(cp, 12, ["Metric", "Weight (0 = ignore)", "Observed value"], start_col=7)
    cp.column_dimensions["F"].width = 3
    cp.column_dimensions["G"].width = 24
    cp.column_dimensions["H"].width = 11
    cp.column_dimensions["I"].width = 13
    for j, key in enumerate(METRICS):
        r = 13 + j
        cp.cell(row=r, column=7, value=f"{METRIC_INFO[key][0]} ({METRIC_INFO[key][1]})").font = f()
        mark_input(cp.cell(row=r, column=8, value=1 if key in DEFAULT_COMPARISON_METRICS else 0))
        c = cp.cell(row=r, column=9, value=f"={summary_row[key]}")
        c.font = f()
        c.number_format = "#,##0.0"
    cp.cell(row=14 + len(METRICS), column=7,
            value="Defaults match the app: WCDH, TWCDH, SWCDH, peak daily WCDH, max event severity, longest event, "
                  "peak temperature and season mean.").font = f(9, color=INK2, italic=True)

    header_row(cp, 12, ["Rank", "Reference file", "Location", "Similarity", "Short label"])
    cp.column_dimensions["B"].width = 44
    cp.column_dimensions["C"].width = 20
    cp.column_dimensions["D"].width = 10
    cp.column_dimensions["E"].width = 26
    for k in range(1, 11):
        r = 12 + k
        cp[f"A{r}"] = k
        cp[f"A{r}"].font = f()
        key_ = f"LARGE({KEY},{k})"
        cp[f"B{r}"] = f'=IF(COUNT({KEY})<{k},"",INDEX({LABEL},MATCH({key_},{KEY},0)))'
        cp[f"C{r}"] = f'=IF(COUNT({KEY})<{k},"",INDEX({LOC},MATCH({key_},{KEY},0)))'
        cp[f"D{r}"] = f'=IF(COUNT({KEY})<{k},0,INDEX({SCORE},MATCH({key_},{KEY},0)))'
        cp[f"D{r}"].number_format = "0"
        cp[f"E{r}"] = f'=SUBSTITUTE(SUBSTITUTE(B{r}," emissions","")," percentile","")'
        cp[f"E{r}"].font = f(9, color=INK2)
        for c_ in "BCD":
            cp[f"{c_}{r}"].font = f()

    # WCDH bracket sentence
    wc = col(8)  # Library WCDH column H
    incl = f"Library!$A${L0}:$A${L1}"
    wv = f"Library!${wc}${L0}:${wc}${L1}"
    tw = summary_row["wcdh"]
    cp["A5"] = "WCDH context"
    cp["A5"].font = f(10, True)
    locs = f"Library!$B${L0}:$B${L1}"
    lc = 'IF(compare_loc="","*",compare_loc)'
    incl = f"{incl},1,{locs},{lc}"  # included and in the Compare location
    lower = f'_xlfn.MAXIFS({wv},{incl},{wv},"<="&{tw})'
    upper = f'_xlfn.MINIFS({wv},{incl},{wv},">"&{tw})'
    lab = lambda v: f"INDEX(Library!$G${L0}:$G${L1},MATCH({v},{wv},0))"  # noqa: E731
    cp["B5"] = (f'=IF(COUNTIFS({incl})=0,"",'
                f'"WCDH "&TEXT({tw},"#,##0")&" K²h "&IF(COUNTIFS({incl},{wv},"<="&{tw})=0,'
                f'"is below every included file (lowest: "&{lab(upper)}&").",'
                f'IF(COUNTIFS({incl},{wv},">"&{tw})=0,"exceeds every included file (highest: "&{lab(lower)}&").",'
                f'"lies between "&{lab(lower)}&" and "&{lab(upper)}&".")))')
    cp["B5"].font = f(10, color=INK2)

    cp["A6"] = "Compare with location"
    cp["A6"].font = f(10, True)
    mark_input(cp["B6"])
    cp["C6"] = ("Choose from the list of Library locations. Blank = every included Library row. "
                "The Years vs DSY sheet compares only with this location.")
    cp["C6"].font = f(9, color=INK2)
    name(wb, "compare_loc", "Compare!$B$6")
    thr_s = f"Library!${col(c_thr_s)}${L0}:${col(c_thr_s)}${L1}"
    thr_t = f"Library!${col(c_thr_t)}${L0}:${col(c_thr_t)}${L1}"
    in_scope = f'(Library!$A${L0}:$A${L1}=1)*((compare_loc="")+({locs}=compare_loc)>0)'
    mismatch = (f'SUMPRODUCT({in_scope}*({thr_s}<>"")*((ABS({thr_s}-swcdh_thr)>0.005)+(ABS({thr_t}-twcdh_off)>0.005)>0))')
    cp["B7"] = (f'=IF({mismatch}=0,"","Check thresholds: "&{mismatch}&" of the files being compared were imported with a '
                f'different SWCDH threshold or TWCDH offset from Settings ("&TEXT(swcdh_thr,"0.00")&" °C, "&TEXT(twcdh_off,"0.00")&" K). '
                f'Run the UseLibraryThresholds macro, or set the overrides to match, so every file is compared on the same basis.")')
    cp["B7"].font = f(10, True, color="C00000")

    bar = BarChart()
    bar.type = "bar"
    bar.style = 10
    bar.title = None
    bar.y_axis.title = "Similarity (0–100)"
    bar.y_axis.scaling.min = 0
    bar.y_axis.scaling.max = 100
    bar.x_axis.scaling.orientation = "maxMin"
    bar.add_data(Reference(cp, min_col=4, min_row=12, max_row=22), titles_from_data=True)
    bar.set_categories(Reference(cp, min_col=5, min_row=13, max_row=22))
    bar.series[0].graphicalProperties = GraphicalProperties(solidFill=BLUE)
    bar.series[0].graphicalProperties.line = LineProperties(noFill=True)
    bar.legend = None
    bar.gapWidth = 60
    bar.height, bar.width = 9, 18
    cp.add_chart(bar, "A25")

    # ---------------------------------------------------------- Chart data (1 Apr – 30 Sep)
    title(cd, "Chart data", "1 April – 30 September of the analysis year, looked up from the Daily sheet.")
    header_row(cd, 4, ["Date", "Daily min (°C)", "Range (K)", "Daily max (°C)", "Tcomf (°C)", "Tmax (°C)", "Cumulative WCDH (K²h)"])
    for c_, w in zip("ABCDEFG", (12, 10, 10, 10, 10, 10, 12)):
        cd.column_dimensions[c_].width = w
    C0 = 5
    n_chart = 183
    for i in range(n_chart):
        r = C0 + i
        cd[f"A{r}"] = f"=DATE(analysis_year,4,1)+{i}"
        idx = f"A{r}-first_day+1"
        look = lambda c_: f'IFERROR(INDEX(Daily!${c_}${D0}:${c_}${D1},{idx}),"")'  # noqa: E731
        cd[f"B{r}"] = f"={look('E')}"
        cd[f"C{r}"] = f"={look('G')}"
        cd[f"D{r}"] = f"={look('F')}"
        cd[f"E{r}"] = f"={look('J')}"
        cd[f"F{r}"] = f"={look('K')}"
        prev = f"N(G{r - 1})" if i else "0"
        cd[f"G{r}"] = f'=IF(IFERROR(INDEX(Daily!$L${D0}:$L${D1},{idx}),0)=1,{prev}+N({look("M")}),{prev})'
        cd[f"A{r}"].number_format = "d mmm"
        for c_ in "BCDEFG":
            cd[f"{c_}{r}"].number_format = "0.0"
    CR = C0 + n_chart - 1

    # ---------------------------------------------------------- Dashboard
    title(dash, "Compare thee to a summer's DSY", "Overheating metrics for one season, compared with a library of CIBSE Design Summer Years.")
    dash.column_dimensions["A"].width = 3
    for c_ in "BCDEFGHIJKLMN":
        dash.column_dimensions[c_].width = 13
    dash["B4"] = "=Compare!B4"
    dash["B4"].font = f(14, True)
    dash["B5"] = "=Compare!B5"
    dash["B5"].font = f(10, color=INK2)
    dash["B6"] = '="Site: "&site_name&" · SWCDH threshold "&TEXT(swcdh_thr,"0.0")&" °C · TWCDH offset "&TEXT(twcdh_off,"+0.0;-0.0")&" K"'
    dash["B6"].font = f(9, color=MUTED)
    dash["B7"] = "=Compare!B7"
    dash["B7"].font = f(10, True, color="C00000")
    tiles = [("WCDH (K²h)", summary_row["wcdh"], "#,##0"), ("Peak temperature (°C)", summary_row["t_max"], "0.0"),
             ("Longest warm event (days)", summary_row["max_event_duration"], "0"),
             ("Hours above Tmax", summary_row["hours_above_upper"], "#,##0"),
             ("Hot days", summary_row["hot_days"], "0"),
             ("Nights ≥ night threshold", summary_row["nights_above"], "0"),
             ("Summer mean daily max (°C)", summary_row["jja_mean_daily_max"], "0.0")]
    for i, (label, ref, fmt) in enumerate(tiles):
        c0 = 2 + i * 2
        lab_c = dash.cell(row=8, column=c0, value=label)
        lab_c.font = f(9, color=INK2)
        v = dash.cell(row=9, column=c0, value=f"={ref}")
        v.font = f(20, True)
        v.number_format = fmt
        v.alignment = Alignment(horizontal="left")
    dash.row_dimensions[9].height = 30

    # One line chart on a single temperature axis, with the diurnal range as columns in a chart below.
    prof = LineChart()
    for c_idx in (2, 5, 6, 4):
        prof.add_data(Reference(cd, min_col=c_idx, min_row=4, max_row=CR), titles_from_data=True)
    for s_, color, dash_style, width in zip(prof.series, (BLUE, AQUA, MUTED, ORANGE), (None, None, "sysDot", None), (2, 2, 1, 2)):
        s_.graphicalProperties.line.solidFill = color
        s_.graphicalProperties.line.width = int(width * 12700)
        if dash_style:
            s_.graphicalProperties.line.dashStyle = dash_style
        s_.smooth = False
        s_.marker.symbol = "none"
    cats = Reference(cd, min_col=1, min_row=C0, max_row=CR)
    prof.set_categories(cats)
    prof.title = None
    prof.y_axis.title = "°C"
    prof.x_axis.number_format = "d mmm"
    prof.x_axis.tickLblSkip = 30
    prof.x_axis.tickMarkSkip = 30
    prof.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E1E0D9"))
    prof.legend.position = "b"
    prof.height, prof.width = 10, 26
    dash["B11"] = "Daily maximum and minimum, comfort temperature Tcomf and upper limit Tmax, 1 April – 30 September"
    dash["B11"].font = f(11, True)

    rng = BarChart()
    rng.add_data(Reference(cd, min_col=3, min_row=4, max_row=CR), titles_from_data=True)
    rng.set_categories(cats)
    rng.series[0].graphicalProperties = GraphicalProperties(solidFill="B7B6AF")
    rng.series[0].graphicalProperties.line = LineProperties(noFill=True)
    rng.gapWidth = 40
    rng.title = None
    rng.legend = None
    rng.y_axis.title = "Range (K)"
    rng.y_axis.scaling.min = 0
    rng.x_axis.number_format = "d mmm"
    rng.x_axis.tickLblSkip = 30
    rng.x_axis.tickMarkSkip = 30
    rng.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E1E0D9"))
    rng.height, rng.width = 5, 26
    dash["B34"] = "Diurnal range: daily maximum minus daily minimum (K)"
    dash["B34"].font = f(10, True)
    dash.add_chart(rng, "B35")
    dash.add_chart(prof, "B13")

    cum = LineChart()
    cum.add_data(Reference(cd, min_col=7, min_row=4, max_row=CR), titles_from_data=True)
    cum.set_categories(cats)
    cum.series[0].graphicalProperties.line.solidFill = BLUE
    cum.series[0].graphicalProperties.line.width = 25400
    cum.series[0].marker.symbol = "none"
    cum.title = "Cumulative WCDH through the season"
    cum.y_axis.title = "K²h"
    cum.x_axis.number_format = "d mmm"
    cum.x_axis.tickLblSkip = 30
    cum.x_axis.tickMarkSkip = 30
    cum.legend = None
    cum.height, cum.width = 8, 13
    dash.add_chart(cum, "B49")
    dash["B48"] = "Steep steps are short, intense spells (DSY2-like). Long steady climbs are prolonged spells (DSY3-like)."
    dash["B48"].font = f(9, color=INK2)

    bar2 = BarChart()
    bar2.type = "bar"
    bar2.title = "Closest reference files (similarity 0–100)"
    bar2.add_data(Reference(cp, min_col=4, min_row=12, max_row=22), titles_from_data=True)
    bar2.set_categories(Reference(cp, min_col=5, min_row=13, max_row=22))
    bar2.series[0].graphicalProperties = GraphicalProperties(solidFill=BLUE)
    bar2.series[0].graphicalProperties.line = LineProperties(noFill=True)
    bar2.x_axis.scaling.orientation = "maxMin"
    bar2.y_axis.scaling.min = 0
    bar2.y_axis.scaling.max = 100
    bar2.legend = None
    bar2.gapWidth = 60
    bar2.height, bar2.width = 8, 13
    dash.add_chart(bar2, "I49")

    # ---------------------------------------------------------- Read me
    readme.column_dimensions["A"].width = 3
    readme.column_dimensions["B"].width = 120
    lines = [
        ("Compare thee to a summer's DSY: Excel edition", f(16, True)),
        ("Overheating metrics for hourly weather, and comparison with CIBSE Design Summer Years. "
         "Every result is a live formula: paste new data and it recalculates.", f(10, color=INK2)),
        ("", None),
        ("Colour legend", f(11, True)),
        ("Yellow cells with blue text are inputs. Everything else is calculated; don't overwrite it.", f()),
        ("", None),
        ("How to use it", f(11, True)),
        ("1. Settings: set the site name, whether the file is a typical/design year, and the season (default TM59, 1 May – 30 Sep).", f()),
        ("2. Hourly: paste Year, Month, Day, Hour (1–24) and dry bulb (°C) into columns A–E. From an EPW, copy columns 1–4 and 7. "
         "Leave Year blank for TRY/DSY files. Clear old rows first. Up to 10,000 hours (one year plus the preceding December).", f()),
        ("3. Dashboard and Summary update automatically: daily max/min/range, comfort temperature, WCDH, TWCDH, SWCDH, warm events.", f()),
        ("4. Library, the quick way (summers_dsy.xlsm): press Alt+F8 (Mac: Tools → Macro → Macros), run ImportWeatherFolder and pick "
         "a folder. Every .epw and .csv file in it and its sub-folders is loaded in turn and its metrics added to the Library. "
         "Location, file type, period, emissions and percentile come from the file name (CIBSE 2016 names like "
         "London_LHR_DSY1_2050High50.epw, or 2025 names like Z1_DSY1_2050s_HIGH50_CIBSE_v1.1.epw), else from the folder name. "
         "Re-importing a file replaces its row. The Import log sheet lists every file and anything skipped. Your Hourly data and "
         "Settings are put back afterwards. Allow about 2–5 seconds per file.", f()),
        ("   The quick way needs macros enabled. A downloaded file may be blocked: in Windows Explorer right-click it → Properties → "
         "tick Unblock. If Excel reports a problem with the macros, open the VBA editor (Alt+F11) → File → Import File → "
         "FolderImport.bas (in the repository's excel folder), then save as .xlsm.", f(10, color=INK2)),
        ("   By hand: paste a CIBSE file into Hourly (Settings → Typical = Yes), then copy the Summary sheet's 'Row to copy' and "
         "Paste Special → Values into a Library row from column H. Fill in location, type, period, emissions and percentile.", f(10, color=INK2)),
        ("5. Paste your observed year back into Hourly (Typical = No). On Compare, type the location to compare with (e.g. Zone 7) "
         "in B6. If Compare warns that thresholds differ, run the UseLibraryThresholds macro. Compare and Dashboard then show "
         "the closest files.", f()),
        ("", None),
        ("Several years from Meteostat (summers_dsy.xlsm)", f(11, True)),
        ("1. On Compare, choose the location to compare with in B6 (the dropdown lists the Library's locations). Years are compared "
         "only with that location's files, using its thresholds.", f()),
        ("2. On Meteostat, choose a station from the list. To find one, type a UK postcode (Excel for Windows looks it up) or a "
         "latitude and longitude: the five nearest stations are listed with their distance and years of hourly data.", f()),
        ("3. Set the first and last years and run DownloadMeteostatYears (Alt+F8). Each year's file is downloaded from "
         "data.meteostat.net (with the December before, to start the running mean), analysed and added to the Years sheet. "
         "Years with under 80% of the May–September hours are listed but not counted.", f()),
        ("4. Years vs DSY lists every file of the location, hottest first, with how many years exceeded it on the metric you "
         "choose (default SWCDH, the metric CIBSE ranks DSY1 by), e.g. '5 of the last 10 years (2016–2025) exceeded DSY1 · "
         "2050s · High emissions · 50th percentile for Zone 7 on SWCDH'.", f()),
        ("   Downloading needs Excel for Windows. On a Mac or offline, save each https://data.meteostat.net/hourly/<year>/<station>.csv.gz, "
         "unzip it, name it <station>_<year>.csv (e.g. 03772_2019.csv), put it in the Download folder set on the Meteostat sheet "
         "and run the macro: files already there are used. By default model-forecast values that Meteostat uses to fill gaps are "
         "dropped. Station data © Meteostat, CC BY 4.0; check meteostat.net's terms for the hourly data before commercial use.",
         f(10, color=INK2)),
        ("", None),
        ("Important: fix the thresholds before building the library", f(11, True)),
        ("SWCDH and TWCDH depend on regional thresholds, and every file in a comparison must use the same values. The folder import "
         "does this for you: with Settings → 'Thresholds for imported files' = Each location's TRY, it derives each location's values "
         "from its earliest-period TRY (else DSY1) and stores them in Library columns Z–AA. By hand: load the location's current-climate "
         "TRY first, copy Settings rows 24–25 into the overrides in rows 22–23, and keep them fixed while you add the other files.", f()),
        ("", None),
        ("Definitions", f(11, True)),
        ("Running mean Trm = (1−α)(Tod-1 + α·Tod-2 + α²·Tod-3 + …), evaluated recursively with α = 0.8 and seeded with the BS EN 15251 "
         "7-day approximation. Days with fewer than 18 hours of data have no daily mean.", f()),
        ("Comfort temperature Tcomf = 0.33·Trm + 18.8 (CIBSE TM52, BS EN 15251). Tmax = Tcomf + 3 K for Category II.", f()),
        ("WCDH = Σ max(0, T − Tcomf)² over the season's hours: the adaptive-comfort form of the metric (Eames 2016).", f()),
        ("TWCDH = Σ max(0, T − (Tcomf + regional offset))². SWCDH = Σ max(0, T − regional 93rd-centile temperature)². CIBSE selects "
         "DSY1 as the 1-in-7 year ranked by SWCDH; DSY2 is the year with the most intense heat event and DSY3 the longest.", f()),
        ("Warm event: a run of consecutive in-season days with daily WCDH above the Settings threshold. Severity = its total WCDH; "
         "intensity = its peak daily WCDH. DSY2 years have short intense events; DSY3 years have long ones.", f()),
        ("TM52-style indicators are computed on outdoor air: hours with rounded ΔT ≥ 1 K above Tmax, the largest daily weighted "
         "exceedance Σ ΔT, and the largest ΔT. They indicate climate severity; they are not building compliance checks.", f()),
        ("Night-time (TM59:2026 analogue): each night runs from 22:00 to 07:00 (Settings rows 44–47) and is labelled by the "
         "evening it starts. 'Warmest night' is the highest nightly mean outdoor temperature in the season. 'Nights ≥ night "
         "threshold' counts nights with a mean at or above 27 °C by default; TM59:2026 allows bedrooms no more than 4 such nights.", f()),
        ("Seasonal climate: mean air temperature for winter (Dec–Feb, using the analysis year's own December), spring, summer and "
         "autumn, as a guide to energy demand; and the summer (Jun–Aug) mean daily maximum, as a guide to overheating risk.", f()),
        ("Similarity: skewed metrics use ln(1+x); each metric is divided by its spread across the included files plus the observed "
         "year; distance = weighted RMS difference; score = 100·e^(−distance).", f()),
        ("", None),
        ("Limitations compared with the web app", f(11, True)),
        ("One weather file analysed at a time (the folder import loops over files for you). No automatic downloads: paste data from Open-Meteo, Meteostat, NOAA ISD, MIDAS Open or the web "
         "app's CSV export. Hourly gaps in pasted data aren't interpolated (the Meteostat macro fills gaps of up to 6 hours). The charts always show 1 April – 30 September. The season must not "
         "wrap past 31 December.", f()),
        ("For typical years, 1 January's running mean is seeded from the last 7 days of December; the web app spins up over 30 days. "
         "Results differ only in early January, not in the summer season.", f()),
        ("", None),
        ("Demo data", f(11, True)),
        ("The workbook ships with a synthetic observed year and a synthetic library ('Demo (synthetic)'). Neither is CIBSE or "
         "Met Office data. Replace both with real data.", f(10, color="C00000")),
        ("", None),
        ("References", f(11, True)),
        ("CIBSE TM52 (2013); CIBSE TM59 (2017, updated 2026); CIBSE TM49 (2014); CIBSE Weather Data 2025 technical briefing; "
         "BS EN 15251:2007 / BS EN 16798-1:2019; "
         "Eames, M. (2016) BSERT 37(5); Liu, Kershaw, Eames & Coley (2016) Building and Environment 105.", f()),
    ]
    for i, (text, font) in enumerate(lines, start=2):
        c = readme.cell(row=i, column=2, value=text)
        if font:
            c.font = font
        c.alignment = Alignment(wrap_text=True, vertical="top")
    readme.sheet_view.showGridLines = False
    dash.sheet_view.showGridLines = False

    # Shade calculated cells lightly on the calculation sheets.
    for ws in (hr,):
        ws.conditional_formatting.add(f"F{H0}:V{H1}", FormulaRule(formula=['$E2=""'], font=Font(color="BFBFBF")))

    # Tab colours
    for ws, color in ((readme, "898781"), (dash, BLUE), (st, "EDA100"), (hr, "EDA100"), (lb, "EDA100"), (cp, BLUE),
                      (ms_ws, "EDA100"), (yv, BLUE)):
        ws.sheet_properties.tabColor = color

    # ---------------------------------------------------------- Import log (written by the macro)
    title(lg, "Import log", "Written by the ImportWeatherFolder macro: one line per file found.")
    header_row(lg, 3, ["File", "Result", "Location", "File type", "Period", "Emissions", "Percentile", "Library row", "Notes"])
    for c_, w in zip("ABCDEFGHI", (70, 16, 18, 9, 9, 10, 10, 10, 70)):
        lg.column_dimensions[c_].width = w
    lg.freeze_panes = "A4"
    name(wb, "log_start", "'Import log'!$A$4")
    name(wb, "hourly_input", f"Hourly!$A${H0}:$E${H1}")
    name(wb, "summary_values", f"Summary!$B$5:$B${4 + len(METRICS)}")

    build_meteostat_sheets(wb, ms_ws, yv, yr_ws, stn, cc, cp, summary_row, R0, R1, c_loc, c_thr_s, c_thr_t)

    # Code names, so the VBA project's document modules bind to the workbook and its sheets.
    wb.code_name = "ThisWorkbook"
    for i, ws in enumerate(wb.worksheets, start=1):
        ws.sheet_properties.codeName = f"Sheet{i}"

    wb.active = 1
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUT)
    return OUT


def build_meteostat_sheets(wb, ms_ws, yv, yr_ws, stn, cc, cp, summary_row, R0, R1, c_loc, c_thr_s, c_thr_t):
    """Stations list, Meteostat inputs, the Years table, 'Years vs DSY' statements and the location dropdown."""
    nm = len(METRICS)
    lib_metrics = f"Library!$H${L0}:${col(7 + nm)}${L1}"

    # ---------------------------------------------------------- Stations (bundled UK list)
    stations = bundled_stations()
    title(stn, "Stations", f"UK Meteostat stations with hourly temperature observations. {ATTRIBUTION}")
    heads = ["Meteostat ID", "Name", "Region", "Latitude", "Longitude", "Elevation (m)", "WMO", "ICAO",
             "Hourly data from", "Hourly data to", "Label", "Distance (km)"]
    header_row(stn, 4, heads)
    for c_, w in zip("ABCDEFGHIJKL", (12, 34, 8, 9, 10, 9, 8, 7, 9, 9, 44, 10)):
        stn.column_dimensions[c_].width = w
    S0, S1 = 5, 4 + len(stations)
    for i, rec in enumerate(stations.itertuples(index=False)):
        r = S0 + i
        vals = [rec.id, rec.name, rec.region if isinstance(rec.region, str) else "", rec.latitude, rec.longitude,
                None if rec.elevation != rec.elevation else int(rec.elevation),
                rec.wmo if isinstance(rec.wmo, str) else "", rec.icao if isinstance(rec.icao, str) else "",
                None if rec.start_year != rec.start_year else int(rec.start_year),
                None if rec.end_year != rec.end_year else int(rec.end_year), rec.label]
        for j, v in enumerate(vals, start=1):
            stn.cell(row=r, column=j, value=v).font = f()
        stn[f"L{r}"] = (f'=IF(OR(ms_lat="",ms_lon=""),"",6371*2*ASIN(SQRT(SIN(RADIANS(D{r}-ms_lat)/2)^2+'
                        f'COS(RADIANS(ms_lat))*COS(RADIANS(D{r}))*SIN(RADIANS(E{r}-ms_lon)/2)^2)))')
        stn[f"L{r}"].number_format = "0.0"
    stn.freeze_panes = "C5"
    name(wb, "station_labels", f"Stations!$K${S0}:$K${S1}")
    st_rng = lambda c_: f"Stations!${c_}${S0}:${c_}${S1}"  # noqa: E731

    # ---------------------------------------------------------- Meteostat inputs
    title(ms_ws, "Meteostat: download and analyse several years",
          "Pick a station and a range of years, then run the DownloadMeteostatYears macro (Alt+F8). Each year is "
          "downloaded, analysed with the Compare location's thresholds and added to the Years sheet. "
          "Years vs DSY then counts how many years exceeded each reference file.")
    ms_ws.column_dimensions["A"].width = 34
    ms_ws.column_dimensions["B"].width = 44
    ms_ws.column_dimensions["C"].width = 12
    ms_ws.column_dimensions["D"].width = 80
    def inp(r, label, value, nm_, note="", fmt=None):
        ms_ws[f"A{r}"] = label
        ms_ws[f"A{r}"].font = f()
        c = ms_ws[f"B{r}"]
        c.value = value
        mark_input(c)
        if fmt:
            c.number_format = fmt
        ms_ws[f"D{r}"] = note
        ms_ws[f"D{r}"].font = f(9, color=INK2)
        if nm_:
            name(wb, nm_, f"Meteostat!$B${r}")
    def calc(r, label, formula, nm_=None, fmt=None):
        ms_ws[f"A{r}"] = label
        ms_ws[f"A{r}"].font = f()
        ms_ws[f"B{r}"] = formula
        ms_ws[f"B{r}"].font = f(10, True)
        if fmt:
            ms_ws[f"B{r}"].number_format = fmt
        if nm_:
            name(wb, nm_, f"Meteostat!$B${r}")

    ms_ws["A4"] = "1. Station"
    ms_ws["A4"].font = f(11, True)
    inp(5, "Station", "London Heathrow Airport (03772)", "ms_station",
        "Choose from the list, or type a Meteostat station ID. Use the nearest-station helper below to find one.")
    calc(6, "Meteostat station ID",
         f'=IFERROR(INDEX({st_rng("A")},MATCH(ms_station,{st_rng("K")},0)),TRIM(ms_station))', "ms_station_id")
    calc(7, "Hourly data available",
         f'=IFERROR(INDEX({st_rng("I")},MATCH(ms_station_id,{st_rng("A")},0))&"–"&INDEX({st_rng("J")},MATCH(ms_station_id,{st_rng("A")},0)),"not in the station list")')
    calc(8, "Station name", f'=IFERROR(INDEX({st_rng("B")},MATCH(ms_station_id,{st_rng("A")},0)),ms_station_id)', "ms_station_name")
    calc(9, "Station key (Years sheet)", '=ms_station_name&" ("&ms_station_id&")"', "ms_key")

    ms_ws["A10"] = "Find the nearest stations to a site"
    ms_ws["A10"].font = f(11, True)
    inp(11, "UK postcode", None, "ms_postcode",
        "Looked up with postcodes.io (Excel for Windows only, via WEBSERVICE). Otherwise type the latitude and longitude.")
    inp(12, "or latitude", None, None, "Decimal degrees, e.g. 51.5072", "0.0000")
    inp(13, "and longitude", None, None, "Decimal degrees, e.g. -0.1276 (west is negative)", "0.0000")
    ms_ws["F11"] = ('=IF(ms_postcode="","",IFERROR(_xlfn.WEBSERVICE("https://api.postcodes.io/postcodes/"'
                    '&SUBSTITUTE(ms_postcode," ","")),""))')
    ms_ws["F11"].font = f(8, color=MUTED)
    ms_ws.column_dimensions["F"].hidden = True
    jnum = lambda key: (f'IFERROR(_xlfn.NUMBERVALUE(MID($F$11,SEARCH("""{key}"":",$F$11)+{len(key) + 3},'  # noqa: E731
                        f'SEARCH(",",$F$11,SEARCH("""{key}"":",$F$11))-SEARCH("""{key}"":",$F$11)-{len(key) + 3}),"."),"")')
    calc(14, "Site latitude", f'=IF(B12<>"",B12,{jnum("latitude")})', "ms_lat", "0.0000")
    calc(15, "Site longitude", f'=IF(B13<>"",B13,{jnum("longitude")})', "ms_lon", "0.0000")
    ms_ws["A17"] = "Nearest stations (copy one into Station above)"
    ms_ws["A17"].font = f(10, True)
    for k in range(1, 6):
        r = 17 + k
        ms_ws[f"A{r}"] = k
        ms_ws[f"A{r}"].alignment = Alignment(horizontal="right")
        small = f"SMALL({st_rng('L')},{k})"
        ms_ws[f"B{r}"] = f'=IFERROR(INDEX({st_rng("K")},MATCH({small},{st_rng("L")},0)),"")'
        ms_ws[f"C{r}"] = f'=IFERROR({small},"")'
        ms_ws[f"C{r}"].number_format = '0.0" km"'
        ms_ws[f"D{r}"] = (f'=IFERROR("hourly data "&INDEX({st_rng("I")},MATCH({small},{st_rng("L")},0))&"–"'
                          f'&INDEX({st_rng("J")},MATCH({small},{st_rng("L")},0)),"")')
        for c_ in "ABCD":
            ms_ws[f"{c_}{r}"].font = f(10, color=INK2 if c_ != "B" else INK)

    ms_ws["A24"] = "2. Years"
    ms_ws["A24"].font = f(11, True)
    inp(25, "From year", 2016, "ms_year_from", "Whole calendar years. The December before is downloaded too, to start the running mean.")
    inp(26, "To year", 2025, "ms_year_to", "The latest complete May–September season is the most recent useful year.")
    inp(27, "Fill gaps with model data?", "No", "ms_include_model",
        "No (recommended): observations only; gaps of up to 6 hours are interpolated and years with under 80% of the "
        "season are not counted. Yes: keep Meteostat's model forecasts where observations are missing.")
    inp(28, "Download folder", None, "ms_folder",
        "Optional. Blank = a summers_dsy_meteostat folder in your temporary folder. Files are named <station>_<year>.csv "
        "and are reused if already there; on a Mac, or offline, save Meteostat's files there yourself (see Read me).")
    ms_ws["A30"] = "3. Run DownloadMeteostatYears (Alt+F8)"
    ms_ws["A30"].font = f(11, True)
    ms_ws["A31"] = "Last run"
    ms_ws["A31"].font = f()
    ms_ws["B31"].font = f(10, color=INK2)
    name(wb, "ms_status", "Meteostat!$B$31")
    ms_ws["A33"] = ATTRIBUTION + " Hourly data © Meteostat, from data.meteostat.net; see meteostat.net for its licence terms."
    ms_ws["A33"].font = f(9, color=MUTED, italic=True)
    for rng, options in (("B27", '"Yes,No"'),):
        dv = DataValidation(type="list", formula1=options, allow_blank=False)
        ms_ws.add_data_validation(dv)
        dv.add(rng)
    dv = DataValidation(type="list", formula1="station_labels", allow_blank=True, showErrorMessage=False)
    ms_ws.add_data_validation(dv)
    dv.add("B5")

    # ---------------------------------------------------------- Years (written by the macro)
    Y0, Y1 = 6, 5 + N_YEARS
    title(yr_ws, "Years", "One row per station-year, written by the DownloadMeteostatYears macro. "
          "Metrics are blank where the season has under 80% of its hours.")
    y_head = (["Year", "Station", "Station ID", "Season coverage", "Model-filled share", "Notes"]
              + [f"{METRIC_INFO[k][0]} ({METRIC_INFO[k][1]})" for k in METRICS]
              + ["SWCDH threshold used (°C)", "TWCDH offset used (K)"])
    header_row(yr_ws, 5, y_head)
    yr_ws.row_dimensions[5].height = 44
    for i, w in enumerate([7, 34, 10, 9, 9, 30] + [11] * nm + [11, 11], start=1):
        yr_ws.column_dimensions[col(i)].width = w
    for r in range(Y0, Y1 + 1):
        yr_ws[f"C{r}"].number_format = "@"  # keep IDs like 03772 as text
        yr_ws[f"D{r}"].number_format = "0.0%"
        yr_ws[f"E{r}"].number_format = "0.0%"
        for j in range(nm):
            yr_ws.cell(row=r, column=7 + j).number_format = "#,##0.0"
    yr_ws.freeze_panes = "D6"
    name(wb, "years_table", f"Years!$A${Y0}:${col(8 + nm)}${Y1}")
    YA, YB = f"Years!$A${Y0}:$A${Y1}", f"Years!$B${Y0}:$B${Y1}"
    ymetric = f"INDEX(Years!$G${Y0}:${col(6 + nm)}${Y1},0,yrs_mi)"

    # ---------------------------------------------------------- Compare calc helpers
    c_cnt, c_dl, c_in, c_val, c_skey = (col(cc[f"{c_loc}{R0}"].column + k) for k in range(1, 6))
    header_row(cc, 4, ["Distinct location count", "Location list", "In Years-vs-DSY location", "Metric value", "Sort key"],
               start_col=cc[f"{c_loc}{R0}"].column + 1)
    for r in range(R0, R1 + 1):
        lr = L0 + (r - R0)
        prev = f"N({c_cnt}{r - 1})"
        cc[f"{c_cnt}{r}"] = (f'=IF(Library!$B{lr}="",{prev},IF(MATCH(Library!$B{lr},Library!$B${L0}:$B${L1},0)={r - R0 + 1},'
                             f'{prev}+1,{prev}))')
        cc[f"{c_dl}{r}"] = f'=IFERROR(INDEX(Library!$B${L0}:$B${L1},MATCH({r - R0 + 1},${c_cnt}${R0}:${c_cnt}${R1},0)),"")'
        cc[f"{c_in}{r}"] = (f'=IF(AND(compare_loc<>"",Library!$A{lr}=1,Library!$B{lr}=compare_loc,'
                            f'ISNUMBER(INDEX(Library!$H{lr}:${col(7 + nm)}{lr},1,yrs_mi))),1,0)')
        cc[f"{c_val}{r}"] = f'=IF({c_in}{r}=1,INDEX(Library!$H{lr}:${col(7 + nm)}{lr},1,yrs_mi),"")'
        cc[f"{c_skey}{r}"] = f'=IF({c_in}{r}=1,{c_val}{r}-ROW()*0.000000001,"")'
    name(wb, "location_list", f"OFFSET('Compare calc'!${c_dl}${R0},0,0,MAX(1,MAX('Compare calc'!${c_cnt}${R0}:${c_cnt}${R1})),1)")
    dv = DataValidation(type="list", formula1="location_list", allow_blank=True, showErrorMessage=False)
    cp.add_data_validation(dv)
    dv.add("B6")
    SKEY = f"'Compare calc'!${c_skey}${R0}:${c_skey}${R1}"
    SVAL = f"'Compare calc'!${c_val}${R0}:${c_val}${R1}"

    # ---------------------------------------------------------- Years vs DSY
    title(yv, "Years vs DSY", "How many of the downloaded years exceeded each reference file of one location. "
          "Only files for the Compare sheet's location are used.")
    yv.column_dimensions["A"].width = 30
    yv.column_dimensions["B"].width = 48
    for c_, w in zip("CDEFG", (12, 10, 8, 30, 110)):
        yv.column_dimensions[c_].width = w
    rows = [
        (4, "Location (choose on Compare, B6)", '=IF(compare_loc="","(choose a location on the Compare sheet)",compare_loc)', None),
        (5, "Metric", "SWCDH", "yrs_metric"),
        (6, "Station", '=ms_key', None),
        (7, "Years counted", f'=COUNTIFS({YB},ms_key,{YA},">="&ms_year_from,{YA},"<="&ms_year_to,{ymetric},"<>")', "yrs_n"),
        (8, "First year", f'=IF(yrs_n=0,"",_xlfn.MINIFS({YA},{YB},ms_key,{YA},">="&ms_year_from,{YA},"<="&ms_year_to,{ymetric},"<>"))', "yrs_first"),
        (9, "Last year", f'=IF(yrs_n=0,"",_xlfn.MAXIFS({YA},{YB},ms_key,{YA},">="&ms_year_from,{YA},"<="&ms_year_to,{ymetric},"<>"))', "yrs_last"),
        (10, "Latest complete season", '=IF(TODAY()>=DATE(YEAR(TODAY()),10,1),YEAR(TODAY()),YEAR(TODAY())-1)', "yrs_latest"),
    ]
    for r, label, value, nm_ in rows:
        yv[f"A{r}"] = label
        yv[f"A{r}"].font = f()
        yv[f"B{r}"] = value
        if r == 5:
            mark_input(yv["B5"])
        else:
            yv[f"B{r}"].font = f(10, True)
        if nm_:
            name(wb, nm_, f"'Years vs DSY'!$B${r}")
    yv["Z5"] = "=MATCH(yrs_metric,metric_labels,0)"
    yv["Z5"].font = f(8, color=MUTED)
    yv.column_dimensions["Z"].hidden = True
    name(wb, "yrs_mi", "'Years vs DSY'!$Z$5")
    yv["C5"] = "CIBSE ranks DSY1 by SWCDH. A year exceeds a file when its value is higher."
    yv["C5"].font = f(9, color=INK2)
    dv = DataValidation(type="list", formula1="metric_labels", allow_blank=False)
    yv.add_data_validation(dv)
    dv.add("B5")
    yv["A11"] = ('=IF(yrs_n=0,"No analysed years for this station and range yet: run DownloadMeteostatYears.",'
                 'IF(compare_loc="","Choose the location to compare with on the Compare sheet (B6).",""))')
    yv["A11"].font = f(10, True, color="C00000")
    # Years analysed with thresholds other than the location's
    loc_thr_s = (f'_xlfn.MINIFS(Library!${col(c_thr_s)}${L0}:${col(c_thr_s)}${L1},Library!$B${L0}:$B${L1},compare_loc,'
                 f'Library!$A${L0}:$A${L1},1)')
    yv["A12"] = (f'=IF(OR(compare_loc="",yrs_n=0,COUNTIFS(Library!$B${L0}:$B${L1},compare_loc,'
                 f'Library!${col(c_thr_s)}${L0}:${col(c_thr_s)}${L1},"<>")=0),"",IF(SUMPRODUCT(({YB}=ms_key)*'
                 f'(Years!${col(7 + nm)}${Y0}:${col(7 + nm)}${Y1}<>"")*(ABS(Years!${col(7 + nm)}${Y0}:${col(7 + nm)}${Y1}-{loc_thr_s})>0.005))>0,'
                 f'"Some years were analysed with different thresholds from this location\'s files: run DownloadMeteostatYears again.",""))')
    yv["A12"].font = f(10, True, color="C00000")

    T0 = 15
    header_row(yv, T0 - 1, ["Reference file", "Short label", "File value", "Years exceeded", "Years", "Which years", "Statement"])
    which_rng = f"(({YB}=ms_key)*({YA}>=ms_year_from)*({YA}<=ms_year_to))"
    for k in range(1, N_REFS + 1):
        r = T0 + k - 1
        key = f"LARGE({SKEY},{k})"
        idx = f"MATCH({key},{SKEY},0)"
        yv[f"A{r}"] = f'=IF(COUNT({SKEY})<{k},"",INDEX(Library!$G${L0}:$G${L1},{idx}))'
        yv[f"B{r}"] = f'=SUBSTITUTE(SUBSTITUTE(A{r}," emissions","")," percentile","")'
        yv[f"C{r}"] = f'=IF(A{r}="","",INDEX({SVAL},{idx}))'
        yv[f"D{r}"] = (f'=IF(A{r}="","",COUNTIFS({ymetric},">"&C{r},{YB},ms_key,{YA},">="&ms_year_from,'
                       f'{YA},"<="&ms_year_to))')
        yv[f"E{r}"] = f'=IF(A{r}="","",yrs_n)'
        yv[f"F{r}"] = ArrayFormula(f"F{r}", f'=IF(A{r}="","",_xlfn.TEXTJOIN(", ",TRUE,IF({which_rng}*ISNUMBER({ymetric})'
                                            f'*(IFERROR({ymetric}*1,0)>C{r}),{YA},"")))')
        yv[f"G{r}"] = (f'=IF(OR(A{r}="",yrs_n=0),"",IF(yrs_n>1,D{r}&" of "&IF(yrs_last>=yrs_latest,"the last ","the ")'
                       f'&yrs_n&" years ("&yrs_first&"–"&yrs_last&") exceeded ",IF(D{r}>0,yrs_first&" exceeded ",'
                       f'yrs_first&" did not exceed "))&A{r}&" for "&compare_loc&" on "&yrs_metric)')
        yv[f"C{r}"].number_format = "#,##0.0"
        yv.row_dimensions[r].height = 15
        for c_ in "ABCDEFG":
            yv[f"{c_}{r}"].font = f(10, color=INK2 if c_ in "BEF" else INK)
    yv.freeze_panes = f"A{T0}"
    chart = BarChart()
    chart.type = "bar"
    chart.title = "Years exceeding each file (first 24, hottest first)"
    chart.add_data(Reference(yv, min_col=4, min_row=T0 - 1, max_row=T0 + 23), titles_from_data=True)
    chart.set_categories(Reference(yv, min_col=2, min_row=T0, max_row=T0 + 23))
    chart.series[0].graphicalProperties = GraphicalProperties(solidFill=BLUE)
    chart.series[0].graphicalProperties.line = LineProperties(noFill=True)
    chart.x_axis.scaling.orientation = "maxMin"
    chart.y_axis.scaling.min = 0
    chart.legend = None
    chart.gapWidth = 60
    chart.height, chart.width = 14, 16
    yv.add_chart(chart, "I4")


if __name__ == "__main__":
    print(build())

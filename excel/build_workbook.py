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

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from summers_dsy.compare import LOG_METRICS, resolve_thresholds  # noqa: E402
from summers_dsy.metrics import DEFAULT_COMPARISON_METRICS, METRIC_INFO, AnalysisConfig, analyse  # noqa: E402
from summers_dsy.model import WeatherSeries  # noqa: E402
from summers_dsy.synthetic import WarmSpell, demo_library, synthetic_year  # noqa: E402

OUT = Path(__file__).resolve().parent / "summers_dsy.xlsx"

N_HOURS = int(__import__("os").environ.get("WB_HOURS", 10_000))  # hourly rows available (a year plus a preceding December fits)
N_DAYS = 420
N_LIB = 100
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
        ("In-season T − Tcomf (K)", 11), ("Month no.", 7),
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
        dy[f"F{r}"] = f'=IF(OR(B{r}="",N(C{r})=0),"",MAX({day("E")}))'
        dy[f"G{r}"] = f'=IF(E{r}="","",F{r}-E{r})'
        dy[f"H{r}"] = f'=IF(B{r}="","",IF(COUNT(X{r}:AD{r})=7,SUMPRODUCT(X{r}:AD{r},rm_weights)/SUM(rm_weights),""))'
        dy[f"I{r}"] = (f'=IF(B{r}="","",IF(trm_method="7-day approximation",H{r},'
                       f'IF(AND(ISNUMBER(I{r - 1}),ISNUMBER(D{r - 1})),(1-alpha)*D{r - 1}+alpha*I{r - 1},H{r})))')
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
            dy[f"{c_}{r}"] = (f'=IF(B{r}="","",IF(A{r}-{k}>=1,INDEX($D${D0}:$D${D1},A{r}-{k}),'
                              f'IF(cyclic="Yes",INDEX($D${D0}:$D${D1},A{r}-{k}+n_days),"")))')
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
    for rr in (r + 1, r + 2):
        sm[f"B{rr}"].number_format = "0.00"
        for c_ in "ABC":
            sm[f"{c_}{rr}"].font = f()

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
          "One row per CIBSE TRY/DSY file. Run each file through the workbook with the same Settings thresholds, "
          "then paste its Summary row here as values. Set Include to 1 for the files to compare.")
    lb["A3"] = ("Demo rows are synthetic (not CIBSE data), calculated by the Python app with the Settings demo thresholds. "
                "Delete them before adding your own files.")
    lb["A3"].font = f(9, color="C00000", italic=True)
    lib_head = ["Include (1/0)", "Location", "File type", "Period", "Emissions", "Percentile", "Label"] + [
        f"{METRIC_INFO[k][0]} ({METRIC_INFO[k][1]})" for k in METRICS]
    header_row(lb, 5, lib_head)
    lb.row_dimensions[5].height = 44
    for i, w in enumerate([9, 20, 9, 10, 11, 10, 44] + [12] * len(METRICS), start=1):
        lb.column_dimensions[col(i)].width = w
    lb.freeze_panes = "H6"
    for i in range(N_LIB):
        r = L0 + i
        for c_idx in list(range(1, 7)) + list(range(8, 8 + len(METRICS))):
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
    for rng, options in ((f"A{L0}:A{L1}", '"1,0"'), (f"C{L0}:C{L1}", '"TRY,DSY1,DSY2,DSY3"'),
                         (f"D{L0}:D{L1}", '"Baseline,2020s,2050s,2080s"'), (f"E{L0}:E{L1}", '"Low,Medium,High"'),
                         (f"F{L0}:F{L1}", '"10,50,90"')):
        dv = DataValidation(type="list", formula1=options, allow_blank=True)
        lb.add_data_validation(dv)
        dv.add(rng)

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
            cc[f"{c_}{r}"] = (f'=IF(OR(Library!$A{lr}<>1,Library!{lib_c}{lr}=""),"",'
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
        cc[f"{c_dist}{r}"] = (f'=IF(OR(Library!$A{lr}<>1,COUNT({zr})=0,{wsum}<=0),"",'
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
    lower = f'_xlfn.MAXIFS({wv},{incl},1,{wv},"<="&{tw})'
    upper = f'_xlfn.MINIFS({wv},{incl},1,{wv},">"&{tw})'
    lab = lambda v: f"INDEX(Library!$G${L0}:$G${L1},MATCH({v},{wv},0))"  # noqa: E731
    cp["B5"] = (f'=IF(COUNTIFS({incl},1)=0,"",'
                f'"WCDH "&TEXT({tw},"#,##0")&" K²h "&IF(COUNTIFS({incl},1,{wv},"<="&{tw})=0,'
                f'"is below every included file (lowest: "&{lab(upper)}&").",'
                f'IF(COUNTIFS({incl},1,{wv},">"&{tw})=0,"exceeds every included file (highest: "&{lab(lower)}&").",'
                f'"lies between "&{lab(lower)}&" and "&{lab(upper)}&".")))')
    cp["B5"].font = f(10, color=INK2)

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
    tiles = [("WCDH (K²h)", summary_row["wcdh"], "#,##0"), ("Peak temperature (°C)", summary_row["t_max"], "0.0"),
             ("Longest warm event (days)", summary_row["max_event_duration"], "0"),
             ("Hours above Tmax", summary_row["hours_above_upper"], "#,##0"),
             ("Hot days", summary_row["hot_days"], "0")]
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
        ("4. Library: for each CIBSE file, paste it into Hourly (with Settings → Typical = Yes), then copy the Summary sheet's "
         "'Row to copy' and Paste Special → Values into a Library row from column H. Fill in location, type, period, emissions, percentile.", f()),
        ("5. Paste your observed year back into Hourly (Typical = No). Compare and Dashboard show the closest files.", f()),
        ("", None),
        ("Important: fix the thresholds before building the library", f(11, True)),
        ("SWCDH and TWCDH depend on regional thresholds. Every file in a comparison must use the same values. Load the location's "
         "current-climate TRY first, copy Settings rows 24–25 (derived values) into the override cells in rows 22–23, then keep them "
         "fixed while you add the other files and your observed years.", f()),
        ("", None),
        ("Definitions", f(11, True)),
        ("Running mean Trm = (1−α)(Tod-1 + α·Tod-2 + α²·Tod-3 + …), evaluated recursively with α = 0.8 and seeded with the BS EN 15251 "
         "7-day approximation. Days with fewer than 18 hours of data have no daily mean.", f()),
        ("Comfort temperature Tcomf = 0.33·Trm + 18.8 (CIBSE TM52, BS EN 15251). Tmax = Tcomf + 3 K for Category II.", f()),
        ("WCDH = Σ max(0, T − Tcomf)² over the season's hours (Eames 2016, the metric behind the CIBSE 2016 probabilistic DSYs).", f()),
        ("TWCDH = Σ max(0, T − (Tcomf + regional offset))². SWCDH = Σ max(0, T − regional 93rd-centile temperature)².", f()),
        ("Warm event: a run of consecutive in-season days with daily WCDH above the Settings threshold. Severity = its total WCDH; "
         "intensity = its peak daily WCDH. DSY2 years have short intense events; DSY3 years have long ones.", f()),
        ("TM52-style indicators are computed on outdoor air: hours with rounded ΔT ≥ 1 K above Tmax, the largest daily weighted "
         "exceedance Σ ΔT, and the largest ΔT. They indicate climate severity; they are not building compliance checks.", f()),
        ("Similarity: skewed metrics use ln(1+x); each metric is divided by its spread across the included files plus the observed "
         "year; distance = weighted RMS difference; score = 100·e^(−distance).", f()),
        ("", None),
        ("Limitations compared with the web app", f(11, True)),
        ("One weather file at a time. No automatic downloads: paste data from Open-Meteo, Meteostat, NOAA ISD, MIDAS Open or the web "
         "app's CSV export. Daily gaps aren't interpolated. The charts always show 1 April – 30 September. The season must not "
         "wrap past 31 December.", f()),
        ("For typical years, 1 January's running mean is seeded from the last 7 days of December; the web app spins up over 30 days. "
         "Results differ only in early January, not in the summer season.", f()),
        ("", None),
        ("Demo data", f(11, True)),
        ("The workbook ships with a synthetic observed year and a synthetic library ('Demo (synthetic)'). Neither is CIBSE or "
         "Met Office data. Replace both with real data.", f(10, color="C00000")),
        ("", None),
        ("References", f(11, True)),
        ("CIBSE TM52 (2013); CIBSE TM59 (2017); CIBSE TM49 (2014); BS EN 15251:2007 / BS EN 16798-1:2019; "
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
        ws.conditional_formatting.add(f"F{H0}:T{H1}", FormulaRule(formula=['$E2=""'], font=Font(color="BFBFBF")))

    # Tab colours
    for ws, color in ((readme, "898781"), (dash, BLUE), (st, "EDA100"), (hr, "EDA100"), (lb, "EDA100"), (cp, BLUE)):
        ws.sheet_properties.tabColor = color

    wb.active = 1
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUT)
    return OUT


if __name__ == "__main__":
    print(build())

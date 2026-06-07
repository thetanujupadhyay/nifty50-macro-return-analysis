import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00
import warnings
warnings.filterwarnings("ignore")

# ── COLOUR PALETTE (industry standard) ───────────────────
NAVY        = "1B3A6B"   # headers
LIGHT_BLUE  = "D6E4F0"   # subheaders
DARK_GREY   = "2C2C2C"   # body text
MID_GREY    = "F2F2F2"   # alternating rows
WHITE       = "FFFFFF"
GREEN       = "1E7145"   # significant / good
RED         = "C0392B"   # not significant / bad
AMBER       = "D4870A"   # borderline
GOLD        = "F0C040"   # highlight accent

def navy_fill():   return PatternFill("solid", fgColor=NAVY)
def blue_fill():   return PatternFill("solid", fgColor=LIGHT_BLUE)
def grey_fill():   return PatternFill("solid", fgColor=MID_GREY)
def green_fill():  return PatternFill("solid", fgColor="C6EFCE")
def red_fill():    return PatternFill("solid", fgColor="FADADD")
def amber_fill():  return PatternFill("solid", fgColor="FFEEBA")
def white_fill():  return PatternFill("solid", fgColor=WHITE)

def hdr_font(size=11, bold=True, color=WHITE):
    return Font(name="Arial", bold=bold, size=size, color=color)
def body_font(size=10, bold=False, color=DARK_GREY):
    return Font(name="Arial", bold=bold, size=size, color=color)
def title_font(size=16, color=NAVY):
    return Font(name="Arial", bold=True, size=size, color=color)

def thin_border():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def center(): return Alignment(horizontal="center", vertical="center", wrap_text=True)
def left():   return Alignment(horizontal="left",   vertical="center", wrap_text=True)
def right():  return Alignment(horizontal="right",  vertical="center")

def set_col_width(ws, col, width):
    ws.column_dimensions[get_column_letter(col)].width = width

def write_header_row(ws, row, cols, texts, widths=None):
    for i, (col, text) in enumerate(zip(cols, texts)):
        c = ws.cell(row=row, column=col, value=text)
        c.font      = hdr_font()
        c.fill      = navy_fill()
        c.alignment = center()
        c.border    = thin_border()
        if widths: set_col_width(ws, col, widths[i])

def write_data_row(ws, row, cols, values, alt=False, bold=False):
    fill = grey_fill() if alt else white_fill()
    for col, val in zip(cols, values):
        c = ws.cell(row=row, column=col, value=val)
        c.font      = body_font(bold=bold)
        c.fill      = fill
        c.alignment = center()
        c.border    = thin_border()

def section_label(ws, row, col, text, span_end_col):
    c = ws.cell(row=row, column=col, value=text)
    c.font      = Font(name="Arial", bold=True, size=10, color=NAVY)
    c.fill      = blue_fill()
    c.alignment = left()
    c.border    = thin_border()
    ws.merge_cells(start_row=row, start_column=col,
                   end_row=row,   end_column=span_end_col)

# ════════════════════════════════════════════════════════════
# LOAD ALL RESULT FILES
# ════════════════════════════════════════════════════════════
print("Loading result files...")

processed = pd.read_excel("nifty_macro_final.xlsx",
                          sheet_name="Processed_Data", index_col=0)
processed.index = pd.to_datetime(processed.index)

adf_raw = pd.read_excel("phase3_adf_results.xlsx")

vif_raw = pd.read_excel("phase4_vif_results.xlsx")

reg_raw = pd.read_excel("phase5_results.xlsx",
                         sheet_name="Macro_Regression")
fc_raw  = pd.read_excel("phase5_results.xlsx",
                         sheet_name="Forecast_Comparison",
                         index_col=0)
rmse_raw = pd.read_excel("phase5_results.xlsx",
                          sheet_name="RMSE_MAE_Summary")

hac_raw   = pd.read_excel("phase6_robustness_results.xlsx",
                           sheet_name="HAC_vs_OLS")
split_raw = pd.read_excel("phase6_robustness_results.xlsx",
                           sheet_name="PrePost_COVID")
roll_raw  = pd.read_excel("phase6_robustness_results.xlsx",
                           sheet_name="Rolling_Coefficients")
roll_raw["Date"] = pd.to_datetime(roll_raw["Date"])

print("  All files loaded successfully.")

# ════════════════════════════════════════════════════════════
# CREATE WORKBOOK
# ════════════════════════════════════════════════════════════
wb = Workbook()
wb.remove(wb.active)

# ════════════════════════════════════════════════════════════
# SHEET 1 — PROJECT SUMMARY
# ════════════════════════════════════════════════════════════
print("Building Sheet 1: Project Summary...")
ws1 = wb.create_sheet("1. Project Summary")
ws1.sheet_view.showGridLines = False
ws1.row_dimensions[1].height = 10

# Title block
ws1.merge_cells("B2:H4")
c = ws1["B2"]
c.value     = "MACROECONOMIC DETERMINANTS & FORECASTABILITY OF NIFTY 50 RETURNS"
c.font      = Font(name="Arial", bold=True, size=15, color=WHITE)
c.fill      = navy_fill()
c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws1.row_dimensions[2].height = 20
ws1.row_dimensions[3].height = 20
ws1.row_dimensions[4].height = 20

ws1.merge_cells("B5:H5")
c = ws1["B5"]
c.value     = "A Time Series Econometric Study  |  Jan 2008 – Dec 2025  |  212 Monthly Observations"
c.font      = Font(name="Arial", bold=False, size=11, color=NAVY)
c.fill      = blue_fill()
c.alignment = center()
ws1.row_dimensions[5].height = 18

# Study Objective
ws1.row_dimensions[7].height = 14
section_label(ws1, 7, 2, "STUDY OBJECTIVE", 8)
obj_text = ("This study tests whether macroeconomic variables — US equity returns, "
            "India VIX, RBI Repo Rate, and USD/INR exchange rate movements — "
            "contribute statistically significant and economically meaningful predictive "
            "information to NIFTY 50 monthly returns, using rigorous time-series econometrics.")
ws1.merge_cells("B8:H10")
c = ws1["B8"]
c.value     = obj_text
c.font      = body_font(size=10)
c.fill      = white_fill()
c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ws1.row_dimensions[8].height = 16
ws1.row_dimensions[9].height = 16
ws1.row_dimensions[10].height = 16

# Key findings
ws1.row_dimensions[12].height = 14
section_label(ws1, 12, 2, "KEY FINDINGS AT A GLANCE", 8)

findings = [
    ("AUTOCORRELATION",  "Lag_NIFTY p=0.632, R²=0.001",
     "❌ Past returns do NOT predict future returns → supports Weak-Form EMH"),
    ("GLOBAL FACTOR",    "SP500 coef=+0.82, p<0.001",
     "✅ Strongest driver — every 1% US gain adds ~0.82% to NIFTY"),
    ("MONETARY POLICY",  "Lag_Repo coef=-0.007, p=0.027",
     "✅ Tighter policy dampens returns — robust to HAC correction"),
    ("FEAR INDEX",       "VIX coef=-0.0005, p=0.199",
     "❌ Insignificant when SP500 included — risk already captured"),
    ("EXCHANGE RATE",    "Lag_ΔUSDINR coef=-0.003, p=0.361",
     "❌ Monthly FX change not a reliable predictor"),
    ("MODEL FIT",        "Adj R² = 0.436",
     "✅ Strong for equity return models (project expected 0.10–0.30)"),
    ("FORECASTING",      "Model 2 RMSE = 0.0382 vs M1 = 0.0320",
     "⚠️  Macro model overfits — doesn't improve out-of-sample forecast"),
    ("ROBUSTNESS",       "HAC p-values confirmed, BP test p<0.001",
     "✅ Heteroskedasticity corrected — significance conclusions unchanged"),
]

row = 13
colors = [green_fill, green_fill, green_fill, red_fill,
          red_fill, green_fill, amber_fill, green_fill]
for i, (cat, metric, interp) in enumerate(findings):
    ws1.row_dimensions[row].height = 20
    c1 = ws1.cell(row=row, column=2, value=cat)
    c1.font = Font(name="Arial", bold=True, size=9, color=NAVY)
    c1.fill = colors[i]()
    c1.alignment = center()
    c1.border = thin_border()

    c2 = ws1.cell(row=row, column=3, value=metric)
    c2.font = body_font(size=9, bold=True)
    c2.fill = colors[i]()
    c2.alignment = center()
    c2.border = thin_border()

    ws1.merge_cells(start_row=row, start_column=4,
                    end_row=row,   end_column=8)
    c3 = ws1.cell(row=row, column=4, value=interp)
    c3.font = body_font(size=9)
    c3.fill = colors[i]()
    c3.alignment = left()
    c3.border = thin_border()
    row += 1

# Industry positioning statement
ws1.row_dimensions[row+1].height = 14
section_label(ws1, row+1, 2, "INDUSTRY POSITIONING STATEMENT", 8)
ws1.merge_cells(f"B{row+2}:H{row+3}")
stmt = ws1[f"B{row+2}"]
stmt.value = ("\"This project does not attempt to generate trading alpha, but evaluates whether "
              "macroeconomic and global risk indicators contribute incremental predictive information "
              "to Indian equity returns in a controlled econometric setting.\"")
stmt.font      = Font(name="Arial", italic=True, size=10, color=NAVY)
stmt.fill      = blue_fill()
stmt.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws1.row_dimensions[row+2].height = 16
ws1.row_dimensions[row+3].height = 16

# Phase checklist
row += 5
section_label(ws1, row, 2, "METHODOLOGY CHECKLIST", 8)
row += 1
checks = [
    ("✅", "No look-ahead bias",      "All macro variables lagged by 1 month before use"),
    ("✅", "Stationarity tested",     "ADF tests on all variables — USDINR differenced"),
    ("✅", "VIF checked",             "Maximum VIF = 1.58 — no multicollinearity"),
    ("✅", "Train/test respected",    "Train 2008–2022 | Test 2023–2025 — never mixed"),
    ("✅", "Global factor included",  "S&P 500 monthly return as mandatory control"),
    ("✅", "Structural break",        "COVID dummy + Pre/Post COVID split regression"),
    ("✅", "Robust standard errors",  "Newey-West HAC applied — significance confirmed"),
    ("✅", "Forecast compared",       "4 models: Naive, Baseline, Full Macro, Sig-Vars"),
    ("✅", "Limitations stated",      "Overfitting, regime sensitivity, non-normality noted"),
]
for icon, chk, note in checks:
    ws1.row_dimensions[row].height = 16
    ws1.cell(row=row, column=2, value=icon).font = body_font(bold=True)
    ws1.cell(row=row, column=2).fill = white_fill()
    ws1.cell(row=row, column=2).alignment = center()
    ws1.cell(row=row, column=2).border = thin_border()
    c = ws1.cell(row=row, column=3, value=chk)
    c.font = body_font(bold=True)
    c.fill = white_fill()
    c.alignment = left()
    c.border = thin_border()
    ws1.merge_cells(start_row=row, start_column=4, end_row=row, end_column=8)
    cn = ws1.cell(row=row, column=4, value=note)
    cn.font = body_font()
    cn.fill = white_fill()
    cn.alignment = left()
    cn.border = thin_border()
    row += 1

for col, w in zip(range(1,10), [2, 18, 28, 14, 14, 14, 14, 14, 14]):
    set_col_width(ws1, col, w)

# ════════════════════════════════════════════════════════════
# SHEET 2 — DATA OVERVIEW
# ════════════════════════════════════════════════════════════
print("Building Sheet 2: Data Overview...")
ws2 = wb.create_sheet("2. Data Overview")
ws2.sheet_view.showGridLines = False

ws2.merge_cells("B2:H2")
c = ws2["B2"]
c.value = "DATA OVERVIEW — 212 Monthly Observations  |  May 2008 – December 2025"
c.font  = Font(name="Arial", bold=True, size=13, color=WHITE)
c.fill  = navy_fill()
c.alignment = center()
ws2.row_dimensions[2].height = 22

# Data sources table
section_label(ws2, 4, 2, "DATA SOURCES", 8)
src_hdrs = ["Variable", "Description", "Source", "Frequency", "Transform"]
write_header_row(ws2, 5, range(2, 7), src_hdrs,
                 widths=[22, 28, 22, 14, 24])
sources = [
    ("NIFTY_Close",       "NIFTY 50 Index (India)",    "yfinance (^NSEI)",    "Monthly", "→ Log Return"),
    ("SP500_Close",       "S&P 500 Index (USA)",        "yfinance (^GSPC)",    "Monthly", "→ Log Return"),
    ("USDINR",            "USD/INR Exchange Rate",      "yfinance (INR=X)",    "Monthly", "→ 1st Difference"),
    ("VIX",               "India Volatility Index",     "yfinance (^INDIAVIX)","Monthly", "As-is (stationary)"),
    ("Repo_Rate",         "RBI Policy Repo Rate",       "RBI DBIE Key Rates",  "As-changed","Forward-filled → Lag"),
]
for i, row_data in enumerate(sources):
    write_data_row(ws2, 6+i, range(2,7), row_data, alt=(i%2==1))

# Descriptive statistics
section_label(ws2, 13, 2, "DESCRIPTIVE STATISTICS (Final Model Variables)", 8)
stat_hdrs = ["Variable", "Obs", "Mean", "Std Dev", "Min", "Median", "Max", "Stationarity"]
write_header_row(ws2, 14, range(2, 10), stat_hdrs,
                 widths=[22, 8, 10, 10, 10, 10, 10, 18])

stat_vars = {
    "NIFTY_Return":      processed["NIFTY_Return"],
    "SP500_Return":      processed["SP500_Return"],
    "VIX":               processed["VIX"],
    "Repo_Rate":         processed["Repo_Rate"],
    "Delta_USDINR":      processed["Delta_USDINR"],
    "Lag_NIFTY_Return":  processed["Lag_NIFTY_Return"],
    "Lag_Repo":          processed["Lag_Repo"],
    "Lag_Delta_USDINR":  processed["Lag_Delta_USDINR"],
}
stat_flags = {
    "NIFTY_Return": "✅ Stationary (p<0.01)",
    "SP500_Return": "✅ Stationary (p<0.01)",
    "VIX":          "✅ Stationary (p<0.01)",
    "Repo_Rate":    "✅ Stationary (p<0.05)",
    "Delta_USDINR": "✅ Stationary (p<0.01)",
    "Lag_NIFTY_Return": "✅ Stationary",
    "Lag_Repo":         "✅ Stationary (p<0.05)",
    "Lag_Delta_USDINR": "✅ Stationary (p<0.01)",
}

for i, (name, s) in enumerate(stat_vars.items()):
    row_vals = [name, len(s.dropna()),
                round(s.mean(), 4), round(s.std(), 4),
                round(s.min(), 4), round(s.median(), 4),
                round(s.max(), 4), stat_flags.get(name, "")]
    write_data_row(ws2, 15+i, range(2, 10), row_vals, alt=(i%2==1))
    # Color the stationarity cell
    sc = ws2.cell(row=15+i, column=9)
    sc.fill = green_fill()
    sc.font = Font(name="Arial", size=9, color="1E7145", bold=True)

# COVID dummy note
ws2.row_dimensions[25].height = 14
section_label(ws2, 25, 2, "COVID DUMMY VARIABLE", 8)
ws2.merge_cells("B26:H26")
cd = ws2["B26"]
cd.value = "COVID_Dummy = 1 for March 2020 – December 2021 (22 months) | = 0 otherwise (190 months)"
cd.font  = body_font(size=10)
cd.fill  = amber_fill()
cd.alignment = center()
cd.border = thin_border()

# ════════════════════════════════════════════════════════════
# SHEET 3 — STATIONARITY TESTS
# ════════════════════════════════════════════════════════════
print("Building Sheet 3: Stationarity Tests...")
ws3 = wb.create_sheet("3. ADF Stationarity Tests")
ws3.sheet_view.showGridLines = False

ws3.merge_cells("B2:H2")
c = ws3["B2"]
c.value = "PHASE 3 — AUGMENTED DICKEY-FULLER STATIONARITY TESTS"
c.font  = Font(name="Arial", bold=True, size=13, color=WHITE)
c.fill  = navy_fill()
c.alignment = center()
ws3.row_dimensions[2].height = 22

# Hypothesis explanation
ws3.merge_cells("B4:H5")
c = ws3["B4"]
c.value = ("H₀ (Null): Series has a unit root → Non-stationary   |   "
           "H₁ (Alternative): Series is Stationary   |   "
           "Reject H₀ if p-value < 0.05")
c.font  = body_font(size=10, bold=True)
c.fill  = blue_fill()
c.alignment = center()
c.border = thin_border()
ws3.row_dimensions[4].height = 16
ws3.row_dimensions[5].height = 16

# ADF table
section_label(ws3, 7, 2, "ADF TEST RESULTS", 8)
adf_hdrs = ["Variable", "ADF Statistic", "p-value", "Lags Used",
            "Crit. 1%", "Crit. 5%", "Stationary?", "Action Taken"]
write_header_row(ws3, 8, range(2, 10), adf_hdrs,
                 widths=[24, 16, 12, 12, 12, 12, 14, 28])

adf_actions = {
    "NIFTY_Return":          "Used as dependent variable — as-is",
    "SP500_Return":          "Used as independent variable — as-is",
    "VIX":                   "Used as independent variable — as-is",
    "Repo_Rate (level)":     "Forward-filled, lagged 1 month (Lag_Repo)",
    "USDINR (level)":        "❌ REPLACED with 1st difference (Lag_Δ_USDINR)",
    "Lag_Repo":              "Used in regression — confirmed stationary",
    "Lag_USDINR":            "❌ DROPPED — non-stationary level",
    "Δ Repo_Rate (1st diff)":"Confirmed stationary — reference check",
    "Δ USDINR (1st diff)":   "✅ Used as Lag_Delta_USDINR in model",
}

crit_vals = {
    "NIFTY_Return":          (-3.461, -2.875),
    "SP500_Return":          (-3.462, -2.875),
    "VIX":                   (-3.462, -2.875),
    "Repo_Rate (level)":     (-3.462, -2.875),
    "USDINR (level)":        (-3.461, -2.875),
    "Lag_Repo":              (-3.462, -2.875),
    "Lag_USDINR":            (-3.461, -2.875),
    "Δ Repo_Rate (1st diff)":(-3.462, -2.875),
    "Δ USDINR (1st diff)":   (-3.462, -2.875),
}

for i, row_data in adf_raw.iterrows():
    var  = row_data["Variable"]
    stat = row_data["Stationary"]
    lags_val = row_data["Lags"] if "Lags" in row_data.index else "—"
    vals = [var,
        round(row_data["ADF Statistic"], 4),
        round(row_data["p-value"], 4),
        int(lags_val) if lags_val != "—" and not pd.isna(lags_val) else "—",
        crit_vals.get(var, ("",""))[0],
        crit_vals.get(var, ("",""))[1],
        "✅ Yes" if stat == "Yes" else "❌ No",
        adf_actions.get(var, "")]
    write_data_row(ws3, 9+i, range(2, 10), vals, alt=(i%2==1))
    fill = green_fill() if stat == "Yes" else red_fill()
    sc = ws3.cell(row=9+i, column=8)
    sc.fill = fill
    sc.font = Font(name="Arial", size=9,
                   color="1E7145" if stat=="Yes" else "C0392B", bold=True)

# Key insight
r = 9 + len(adf_raw) + 2
section_label(ws3, r, 2, "KEY INSIGHT", 8)
ws3.merge_cells(f"B{r+1}:H{r+2}")
ki = ws3[f"B{r+1}"]
ki.value = ("USDINR price levels are non-stationary (random walk — p=0.82). "
            "Using levels in regression would produce spurious results. "
            "Solution: Replace Lag_USDINR with Lag_Δ_USDINR (monthly change in exchange rate), "
            "which is stationary at p<0.01. All other variables confirmed stationary.")
ki.font  = body_font(size=10)
ki.fill  = blue_fill()
ki.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
ki.border = thin_border()
ws3.row_dimensions[r+1].height = 18
ws3.row_dimensions[r+2].height = 18

# ════════════════════════════════════════════════════════════
# SHEET 4 — VIF CHECK
# ════════════════════════════════════════════════════════════
print("Building Sheet 4: VIF Check...")
ws4 = wb.create_sheet("4. VIF Multicollinearity")
ws4.sheet_view.showGridLines = False

ws4.merge_cells("B2:G2")
c = ws4["B2"]
c.value = "PHASE 4 — VIF MULTICOLLINEARITY CHECK"
c.font  = Font(name="Arial", bold=True, size=13, color=WHITE)
c.fill  = navy_fill()
c.alignment = center()
ws4.row_dimensions[2].height = 22

section_label(ws4, 4, 2, "VIF THRESHOLDS", 6)
thresh = [("VIF = 1",    "No correlation with other variables"),
          ("VIF 1 – 5",  "✅ Acceptable — proceed"),
          ("VIF 5 – 10", "⚠️  Investigate further"),
          ("VIF > 10",   "❌ Serious — must remove variable")]
for i, (t, d) in enumerate(thresh):
    c1 = ws4.cell(row=5+i, column=2, value=t)
    c1.font = body_font(bold=True)
    c1.fill = white_fill()
    c1.alignment = center()
    c1.border = thin_border()
    ws4.merge_cells(start_row=5+i, start_column=3, end_row=5+i, end_column=6)
    c2 = ws4.cell(row=5+i, column=3, value=d)
    c2.font = body_font()
    c2.fill = white_fill()
    c2.alignment = left()
    c2.border = thin_border()

section_label(ws4, 10, 2, "VIF RESULTS", 6)
vif_hdrs = ["Variable", "VIF Score", "Status", "Interpretation"]
write_header_row(ws4, 11, range(2, 6), vif_hdrs, widths=[26, 14, 16, 36])

for i, row_data in vif_raw.iterrows():
    vif = row_data["VIF"]
    status = "✅ Excellent" if vif < 2 else "✅ Acceptable" if vif < 5 else "⚠️ Investigate"
    interp = f"Contributes independent information — no redundancy detected"
    vals = [row_data["Variable"], round(vif, 4), status, interp]
    write_data_row(ws4, 12+i, range(2, 6), vals, alt=(i%2==1))
    ws4.cell(row=12+i, column=4).fill = green_fill()
    ws4.cell(row=12+i, column=4).font = Font(name="Arial", size=9,
                                              color="1E7145", bold=True)

r = 12 + len(vif_raw) + 2
section_label(ws4, r, 2, "VERDICT", 6)
ws4.merge_cells(f"B{r+1}:F{r+1}")
vd = ws4[f"B{r+1}"]
vd.value = ("Maximum VIF = 1.58 (Lag_NIFTY_Return) — well below the 5.0 threshold. "
            "All 6 independent variables contribute independent information. "
            "No variables need to be removed. Model specification is clean.")
vd.font  = Font(name="Arial", bold=True, size=10, color="1E7145")
vd.fill  = green_fill()
vd.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
vd.border = thin_border()
ws4.row_dimensions[r+1].height = 20

for col, w in zip(range(1,8), [2, 26, 14, 16, 36, 14, 14]):
    set_col_width(ws4, col, w)

# ════════════════════════════════════════════════════════════
# SHEET 5 — REGRESSION RESULTS
# ════════════════════════════════════════════════════════════
print("Building Sheet 5: Regression Results...")
ws5 = wb.create_sheet("5. Regression Results")
ws5.sheet_view.showGridLines = False

ws5.merge_cells("B2:J2")
c = ws5["B2"]
c.value = "PHASE 5 — REGRESSION ANALYSIS RESULTS"
c.font  = Font(name="Arial", bold=True, size=13, color=WHITE)
c.fill  = navy_fill()
c.alignment = center()
ws5.row_dimensions[2].height = 22

# Step B — Autocorrelation
section_label(ws5, 4, 2, "STEP B — AUTOCORRELATION TEST  |  Dependent: NIFTY_Return  |  Independent: Lag_NIFTY_Return only", 10)
auto_stats = [
    ("Lag_NIFTY_Return Coefficient", "+0.0362"),
    ("p-value",                       "0.632  ❌ Not significant"),
    ("R²",                            "0.0013  (0.13% of variation explained)"),
    ("Adjusted R²",                   "-0.004  (negative — no predictive power)"),
    ("Durbin-Watson",                 "1.965  ✅ No residual autocorrelation"),
    ("F-statistic",                   "0.230  (p=0.632)"),
]
write_header_row(ws5, 5, [2,3], ["Metric","Result"], widths=[32, 40])
for i, (m, v) in enumerate(auto_stats):
    write_data_row(ws5, 6+i, [2, 3], [m, v], alt=(i%2==1))
    if "❌" in v:
        ws5.cell(row=6+i, column=3).fill = red_fill()

ws5.merge_cells("B13:J13")
interp = ws5["B13"]
interp.value = ("📌 INTERPRETATION: Past NIFTY returns have virtually zero predictive power for future returns. "
                "This supports the Weak-Form Efficient Market Hypothesis for Indian equity markets.")
interp.font  = Font(name="Arial", italic=True, size=10, color=NAVY)
interp.fill  = blue_fill()
interp.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
interp.border = thin_border()
ws5.row_dimensions[13].height = 20

# Step A — Full Macro
section_label(ws5, 15, 2, "STEP A — FULL MACRO REGRESSION  |  Train: May 2008 – Dec 2022  |  N=176", 10)
model_stats = [
    ("R-squared",         "0.4554"),
    ("Adjusted R²",       "0.4357  ✅ Strong for monthly equity return models"),
    ("F-statistic",       "23.52  (p < 0.0001)  ✅ Model jointly significant"),
    ("Durbin-Watson",     "1.975  ✅ No serial correlation in residuals"),
    ("Observations",      "176 monthly observations (May 2008 – Dec 2022)"),
    ("Breusch-Pagan",     "31.09 (p<0.001)  ⚠️  Heteroskedasticity → HAC SEs applied"),
]
write_header_row(ws5, 16, [2,3], ["Model Statistic","Value"], widths=[32, 42])
for i, (m, v) in enumerate(model_stats):
    write_data_row(ws5, 17+i, [2, 3], [m, v], alt=(i%2==1))
    if "✅" in v: ws5.cell(row=17+i, column=3).fill = green_fill()
    if "⚠️" in v: ws5.cell(row=17+i, column=3).fill = amber_fill()

# Coefficient table
section_label(ws5, 24, 2, "COEFFICIENT TABLE (OLS + HAC Robust p-values)", 10)
coef_hdrs = ["Variable", "Coefficient", "Std Error", "t-stat",
             "OLS p-value", "HAC p-value", "Significant?", "Economic Interpretation"]
write_header_row(ws5, 25, range(2, 10), coef_hdrs,
                 widths=[24, 13, 12, 10, 13, 13, 14, 38])

hac_dict = dict(zip(hac_raw["Variable"], hac_raw["HAC_pval"]))

interps = {
    "const":             "Baseline monthly return when all vars = 0",
    "Lag_NIFTY_Return":  "Past returns don't predict future returns (EMH)",
    "SP500_Return":      "✅ Every 1% US gain → ~0.82% NIFTY gain (same month)",
    "VIX":               "Subsumed by SP500 — risk already captured globally",
    "Lag_Repo":          "✅ Higher rates last month → lower return this month",
    "Lag_Delta_USDINR":  "Lagged FX change — weak short-term predictor",
    "COVID_Dummy":       "COVID effect captured by SP500 & VIX jointly",
}
sig_colors = {"Yes": green_fill, "No": red_fill}

for i, row_data in reg_raw.iterrows():
    var = row_data["Variable"]
    hac_p = hac_dict.get(var, "")
    sig_txt = "✅ Yes (OLS+HAC)" if row_data["Significant"]=="Yes" else "❌ No"
    vals = [var,
            round(row_data["Coefficient"], 5),
            round(row_data["Std_Error"], 5),
            round(row_data["t_stat"], 4),
            round(row_data["p_value"], 4),
            round(hac_p, 4) if hac_p != "" else "",
            sig_txt,
            interps.get(var, "")]
    write_data_row(ws5, 26+i, range(2, 10), vals, alt=(i%2==1))
    sig_fill = green_fill() if row_data["Significant"]=="Yes" else red_fill()
    sc = ws5.cell(row=26+i, column=8)
    sc.fill = sig_fill
    sc.font = Font(name="Arial", size=9,
                   color="1E7145" if row_data["Significant"]=="Yes" else "C0392B",
                   bold=True)

for col, w in zip(range(1,11), [2,24,13,12,10,13,13,14,38,14]):
    set_col_width(ws5, col, w)

# ════════════════════════════════════════════════════════════
# SHEET 6 — FORECAST COMPARISON
# ════════════════════════════════════════════════════════════
print("Building Sheet 6: Forecast Comparison...")
ws6 = wb.create_sheet("6. Forecast Comparison")
ws6.sheet_view.showGridLines = False

ws6.merge_cells("B2:J2")
c = ws6["B2"]
c.value = "PHASE 5 — OUT-OF-SAMPLE FORECAST COMPARISON  |  Test Period: Jan 2023 – Dec 2025"
c.font  = Font(name="Arial", bold=True, size=13, color=WHITE)
c.fill  = navy_fill()
c.alignment = center()
ws6.row_dimensions[2].height = 22

# Model comparison summary
section_label(ws6, 4, 2, "MODEL PERFORMANCE SUMMARY", 9)
fc_hdrs = ["Model", "Variables Used", "RMSE", "MAE",
           "vs Baseline RMSE", "vs Naive RMSE", "Verdict"]
write_header_row(ws6, 5, range(2, 9), fc_hdrs,
                 widths=[30, 30, 10, 10, 18, 18, 24])

model_rows = [
    ("Naive (Benchmark)",          "Always predict historical mean",
     0.0321, 0.0257, "—",      "baseline", "📌 Benchmark only"),
    ("Model 1: Baseline",          "Lag_NIFTY_Return only",
     0.0320, 0.0254, "baseline","−0.3%",   "✅ Best forecast model"),
    ("Model 2: Full Macro (6 var)","All 6 macro variables",
     0.0382, 0.0315, "−19.4%", "−19.1%",  "❌ Overfits training data"),
    ("Model 3: Sig. Vars only",    "SP500_Return + Lag_Repo",
     0.0375, 0.0299, "−17.2%", "−16.8%",  "⚠️  Parsimonious but still worse"),
]
fc_fills = [white_fill, green_fill, red_fill, amber_fill]
for i, (model, vars_, rmse, mae, vs_m1, vs_naive, verdict) in enumerate(model_rows):
    vals = [model, vars_, rmse, mae, vs_m1, vs_naive, verdict]
    for j, col in enumerate(range(2, 9)):
        c = ws6.cell(row=6+i, column=col, value=vals[j])
        c.font   = body_font(bold=(i==1))
        c.fill   = fc_fills[i]()
        c.alignment = center()
        c.border = thin_border()

# Explanation
r = 11
section_label(ws6, r, 2, "WHY DOES MODEL 2 PERFORM WORSE OUT-OF-SAMPLE?", 9)
ws6.merge_cells(f"B{r+1}:J{r+3}")
exp = ws6[f"B{r+1}"]
exp.value = ("This is a well-documented phenomenon in financial econometrics called OVERFITTING. "
             "Model 2 has 6 parameters — it learns patterns specific to 2008–2022 "
             "(GFC, demonetisation, IL&FS, COVID) that do not repeat in 2023–2025. "
             "The extra variables add noise instead of signal in the test period. "
             "The near-zero gap between Model 1 and Naive is consistent with the Efficient Market Hypothesis: "
             "it is very difficult to forecast equity returns out-of-sample.")
exp.font  = body_font(size=10)
exp.fill  = blue_fill()
exp.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
exp.border = thin_border()
for rr in range(r+1, r+4):
    ws6.row_dimensions[rr].height = 16

# Actual vs predicted table
r += 5
section_label(ws6, r, 2, "ACTUAL vs PREDICTED RETURNS — TEST PERIOD (2023–2025)", 9)
fc_tbl_hdrs = ["Date", "Actual Return", "Model 1 Pred", "Model 2 Pred",
               "M1 Error", "M2 Error", "M1 Better?"]
write_header_row(ws6, r+1, range(2, 9), fc_tbl_hdrs,
                 widths=[14, 16, 16, 16, 14, 14, 14])
fc_raw_reset = fc_raw.reset_index()
for i, row_data in fc_raw_reset.iterrows():
    dt  = row_data.iloc[0].strftime("%b-%Y") if hasattr(row_data.iloc[0],'strftime') else str(row_data.iloc[0])[:10]
    act = round(row_data["Actual"], 4)
    m1  = round(row_data["Model1_Pred"], 4)
    m2  = round(row_data["Model2_Pred"], 4)
    e1  = round(row_data["M1_Error"], 4)
    e2  = round(row_data["M2_Error"], 4)
    m1_better = "✅ Yes" if abs(e1) < abs(e2) else "❌ No"
    vals = [dt, act, m1, m2, e1, e2, m1_better]
    write_data_row(ws6, r+2+i, range(2, 9), vals, alt=(i%2==1))
    bc = ws6.cell(row=r+2+i, column=8)
    bc.fill = green_fill() if abs(e1) < abs(e2) else red_fill()
    bc.font = Font(name="Arial", size=9,
                   color="1E7145" if abs(e1)<abs(e2) else "C0392B", bold=True)

# ════════════════════════════════════════════════════════════
# SHEET 7 — ROBUSTNESS CHECKS
# ════════════════════════════════════════════════════════════
print("Building Sheet 7: Robustness Checks...")
ws7 = wb.create_sheet("7. Robustness Checks")
ws7.sheet_view.showGridLines = False

ws7.merge_cells("B2:J2")
c = ws7["B2"]
c.value = "PHASE 6 — ROBUSTNESS CHECKS"
c.font  = Font(name="Arial", bold=True, size=13, color=WHITE)
c.fill  = navy_fill()
c.alignment = center()
ws7.row_dimensions[2].height = 22

# Check 1 — HAC
section_label(ws7, 4, 2,
    "CHECK 1 — ROBUST STANDARD ERRORS (Newey-West HAC)  |  "
    "Purpose: Correct for heteroskedasticity & autocorrelation in residuals", 9)
hac_hdrs = ["Variable", "OLS Coef", "OLS p-value", "HAC p-value",
            "Sig (OLS)?", "Sig (HAC)?", "Robustness Verdict"]
write_header_row(ws7, 5, range(2, 9), hac_hdrs,
                 widths=[24, 12, 14, 14, 14, 14, 28])

for i, row_data in hac_raw.iterrows():
    sig_ols = row_data["Sig_OLS"] == "Yes"
    sig_hac = row_data["Sig_HAC"] == "Yes"
    if sig_ols and sig_hac:
        verdict = "✅ Confirmed significant"
        vfill   = green_fill
    elif not sig_ols and not sig_hac:
        verdict = "➡️ Confirmed insignificant"
        vfill   = white_fill
    elif sig_ols and not sig_hac:
        verdict = "⚠️ Lost significance — fragile"
        vfill   = amber_fill
    else:
        verdict = "📈 Gained significance"
        vfill   = green_fill
    vals = [row_data["Variable"],
            round(row_data["OLS_coef"], 5),
            round(row_data["OLS_pval"], 4),
            round(row_data["HAC_pval"], 4),
            "✅ Yes" if sig_ols else "❌ No",
            "✅ Yes" if sig_hac else "❌ No",
            verdict]
    write_data_row(ws7, 6+i, range(2, 9), vals, alt=(i%2==1))
    ws7.cell(row=6+i, column=8).fill = vfill()
    ws7.cell(row=6+i, column=8).font = Font(
        name="Arial", size=9, bold=True,
        color="1E7145" if (sig_ols and sig_hac) else DARK_GREY)

ws7.merge_cells("B13:J13")
h1c = ws7["B13"]
h1c.value = ("✅ KEY RESULT: SP500 and Lag_Repo remain significant after HAC correction. "
             "Heteroskedasticity (BP=31.09, p<0.001) is confirmed — HAC is justified and applied.")
h1c.font  = Font(name="Arial", bold=True, size=10, color="1E7145")
h1c.fill  = green_fill()
h1c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
h1c.border = thin_border()
ws7.row_dimensions[13].height = 20

# Check 2 — Pre/Post COVID
section_label(ws7, 15, 2,
    "CHECK 2 — STRUCTURAL BREAK  |  Pre-COVID (May 2008–Feb 2020, N=142) vs Post-COVID (Mar 2020–Dec 2022, N=34)", 9)
split_hdrs = ["Variable", "Pre-COVID Coef", "Pre-COVID p",
              "Post-COVID Coef", "Post-COVID p",
              "Coefficient Change", "Stable?"]
write_header_row(ws7, 16, range(2, 9), split_hdrs,
                 widths=[22, 18, 16, 18, 16, 20, 16])

for i, row_data in split_raw.iterrows():
    chg   = round(row_data["Post_coef"] - row_data["Pre_coef"], 5)
    pct   = f"{chg/abs(row_data['Pre_coef'])*100:+.1f}%" if row_data["Pre_coef"] != 0 else "N/A"
    stable = "✅ Stable" if abs(chg) < 0.05 else "⚠️ Changed"
    vals = [row_data["Variable"],
            round(row_data["Pre_coef"], 5),
            round(row_data["Pre_pval"], 4),
            round(row_data["Post_coef"], 5),
            round(row_data["Post_pval"], 4),
            f"{chg:+.5f} ({pct})",
            stable]
    write_data_row(ws7, 17+i, range(2, 9), vals, alt=(i%2==1))
    sc = ws7.cell(row=17+i, column=8)
    sc.fill = green_fill() if stable.startswith("✅") else amber_fill()

ws7.merge_cells("B21:J22")
c2i = ws7["B21"]
c2i.value = ("SP500 remains significant in BOTH periods (pre: p<0.001, post: p<0.001) — "
             "strongest evidence of robustness in the study. "
             "Adj R² rises from 0.40 pre-COVID to 0.50 post-COVID, suggesting markets became "
             "more globally synchronised during and after the pandemic.")
c2i.font  = body_font(size=10)
c2i.fill  = blue_fill()
c2i.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
c2i.border = thin_border()
ws7.row_dimensions[21].height = 16
ws7.row_dimensions[22].height = 16

# Check 3 — Rolling window
section_label(ws7, 24, 2,
    "CHECK 3 — ROLLING WINDOW REGRESSION (60-month window)  |  Tracking SP500 & Repo coefficients", 9)
roll_hdrs = ["Date", "SP500 Coef (60m)", "Lag_Repo Coef (60m)", "SP500 Signal"]
write_header_row(ws7, 25, range(2, 6), roll_hdrs,
                 widths=[14, 20, 22, 24])

roll_sample = roll_raw.iloc[::6].reset_index(drop=True)
for i, row_data in roll_sample.iterrows():
    sp_coef  = round(row_data["SP500_coef"], 4)
    rp_coef  = round(row_data["Lag_Repo_coef"], 4)
    if sp_coef > 0.85:    signal = "🔴 High coupling"
    elif sp_coef > 0.70:  signal = "🟡 Moderate coupling"
    else:                  signal = "🟢 Lower coupling"
    vals = [row_data["Date"].strftime("%b-%Y"), sp_coef, rp_coef, signal]
    write_data_row(ws7, 26+i, range(2, 6), vals, alt=(i%2==1))

# Summary stats for rolling
r_end = 26 + len(roll_sample) + 1
ws7.merge_cells(f"B{r_end}:J{r_end}")
rs = ws7[f"B{r_end}"]
rs.value = (f"Rolling SP500 Coef — Mean: {roll_raw['SP500_coef'].mean():.4f}  |  "
            f"Std: {roll_raw['SP500_coef'].std():.4f}  |  "
            f"Min: {roll_raw['SP500_coef'].min():.4f} (Nov-2019)  |  "
            f"Max: {roll_raw['SP500_coef'].max():.4f} (May-2013)  |  "
            f"Coeff. of Variation: {roll_raw['SP500_coef'].std()/roll_raw['SP500_coef'].mean()*100:.1f}%  ✅ Relatively stable")
rs.font  = Font(name="Arial", bold=True, size=10, color=NAVY)
rs.fill  = blue_fill()
rs.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
rs.border = thin_border()
ws7.row_dimensions[r_end].height = 20

for col, w in zip(range(1, 11), [2, 24, 14, 14, 18, 14, 14, 24, 14, 14]):
    set_col_width(ws7, col, w)

# ════════════════════════════════════════════════════════════
# SAVE
# ════════════════════════════════════════════════════════════
OUTPUT = "NIFTY50_TSA_Dashboard.xlsx"
wb.save(OUTPUT)
print(f"\n✅ Dashboard saved: {OUTPUT}")
print("   7 sheets: Project Summary | Data Overview | ADF Tests |")
print("             VIF Check | Regression | Forecast | Robustness")
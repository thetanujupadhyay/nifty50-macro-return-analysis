import pandas as pd
import numpy as np

INPUT  = "nifty_macro_raw.xlsx"
OUTPUT = "nifty_macro_processed.xlsx"

# ── LOAD RAW DATA ────────────────────────────────────────
df = pd.read_excel(INPUT, sheet_name="Raw_Data", index_col=0)
df.index = pd.to_datetime(df.index)
df = df.sort_index()

# ════════════════════════════════════════════════════════
# STEP 1 — LOG RETURNS
# Formula: ln(Pt / Pt-1)
# ════════════════════════════════════════════════════════
df["NIFTY_Return"]  = np.log(df["NIFTY_Close"]  / df["NIFTY_Close"].shift(1))
df["SP500_Return"]  = np.log(df["SP500_Close"]   / df["SP500_Close"].shift(1))

# ════════════════════════════════════════════════════════
# STEP 2 — LAG VARIABLES (previous month's value)
# These avoid look-ahead bias — macro data known BEFORE
# the return month begins
# ════════════════════════════════════════════════════════
df["Lag_NIFTY_Return"] = df["NIFTY_Return"].shift(1)
df["Lag_Repo"]         = df["Repo_Rate"].shift(1)
df["Lag_USDINR"]       = df["USDINR"].shift(1)
# Note: VIX is contemporaneous (available same month, end-of-day)
# We include it as-is per project spec

# ════════════════════════════════════════════════════════
# STEP 3 — COVID DUMMY
# 1 = March 2020 to December 2021, else 0
# ════════════════════════════════════════════════════════
df["COVID_Dummy"] = 0
df.loc[(df.index >= "2020-03-01") & (df.index <= "2021-12-31"), "COVID_Dummy"] = 1

# ════════════════════════════════════════════════════════
# STEP 4 — SELECT FINAL COLUMNS FOR ANALYSIS
# ════════════════════════════════════════════════════════
processed = df[[
    "NIFTY_Close",
    "NIFTY_Return",
    "Lag_NIFTY_Return",
    "SP500_Return",
    "VIX",
    "Repo_Rate",
    "Lag_Repo",
    "USDINR",
    "Lag_USDINR",
    "COVID_Dummy"
]].copy()

# Drop first row (NaN from log return) and second row (NaN from lag)
processed = processed.dropna(subset=["Lag_NIFTY_Return"])

print("── Processed Data Preview ──")
print(processed.head(10).to_string())
print(f"\nShape: {processed.shape}")
print(f"Date range: {processed.index[0].date()} → {processed.index[-1].date()}")
print("\n── Missing Values ──")
print(processed.isnull().sum())
print("\n── COVID Dummy Check ──")
print(processed["COVID_Dummy"].value_counts())

# ════════════════════════════════════════════════════════
# STEP 5 — CORRELATION MATRIX
# Only on the regression variables
# ════════════════════════════════════════════════════════
corr_cols = [
    "NIFTY_Return",
    "Lag_NIFTY_Return",
    "SP500_Return",
    "VIX",
    "Lag_Repo",
    "Lag_USDINR",
    "COVID_Dummy"
]
corr_matrix = processed[corr_cols].corr().round(3)

print("\n── Correlation Matrix ──")
print(corr_matrix.to_string())

# Flag high correlations
print("\n── Pairs with |correlation| > 0.75 (multicollinearity risk) ──")
flagged = False
cols = corr_matrix.columns
for i in range(len(cols)):
    for j in range(i+1, len(cols)):
        val = corr_matrix.iloc[i,j]
        if abs(val) > 0.75:
            print(f"  ⚠️  {cols[i]} ↔ {cols[j]} : {val}")
            flagged = True
if not flagged:
    print("  ✅ No multicollinearity concerns above 0.75 threshold")

# ════════════════════════════════════════════════════════
# STEP 6 — EXPORT TO EXCEL (multi-sheet)
# ════════════════════════════════════════════════════════
with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
    # Sheet 1: Raw data
    df.to_excel(writer, sheet_name="Raw_Data", index=True, index_label="Date")
    # Sheet 2: Processed data (all analytical columns)
    processed.to_excel(writer, sheet_name="Processed_Data", index=True, index_label="Date")
    # Sheet 3: Correlation matrix
    corr_matrix.to_excel(writer, sheet_name="Correlation_Matrix", index=True)

print(f"\n✅ Done! File saved: {OUTPUT}")
print("   Sheets: Raw_Data | Processed_Data | Correlation_Matrix")
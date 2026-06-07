import pandas as pd

INPUT  = "nifty_macro_processed.xlsx"
OUTPUT = "nifty_macro_final.xlsx"  

# ── LOAD ─────────────────────────────────────────────────
df = pd.read_excel(INPUT, sheet_name="Processed_Data", index_col=0)
df.index = pd.to_datetime(df.index)

# ── ADD ΔUSDINR (monthly change in exchange rate) ─────────
# Step 1: compute the change from raw USDINR
df["Delta_USDINR"]     = df["USDINR"].diff()          # change this month
df["Lag_Delta_USDINR"] = df["Delta_USDINR"].shift(1)  # lag by 1 (no look-ahead)

# ── DROP OLD NON-STATIONARY COLUMN ───────────────────────
df = df.drop(columns=["Lag_USDINR"])

# ── REORDER COLUMNS CLEANLY ───────────────────────────────
df = df[[
    "NIFTY_Close",
    "NIFTY_Return",
    "Lag_NIFTY_Return",
    "SP500_Return",
    "VIX",
    "Repo_Rate",
    "Lag_Repo",
    "USDINR",
    "Delta_USDINR",
    "Lag_Delta_USDINR",
    "COVID_Dummy"
]]

# Drop rows where Lag_Delta_USDINR is NaN (first 2 rows)
df = df.dropna(subset=["Lag_Delta_USDINR"])

# ── VERIFY ────────────────────────────────────────────────
print("=== FINAL MODEL VARIABLES ===")
print(f"Rows: {df.shape[0]} | Date: {df.index[0].date()} → {df.index[-1].date()}")
print("\n── Missing Values ──")
print(df.isnull().sum())
print("\n── Preview ──")
print(df[["NIFTY_Return","SP500_Return","VIX","Lag_Repo","Lag_Delta_USDINR","COVID_Dummy"]].head(8))

print("\n── Δ USDINR Sanity Check (positive = rupee weakened) ──")
print(df[["USDINR","Delta_USDINR","Lag_Delta_USDINR"]].head(8))

# ── UPDATED CORRELATION MATRIX ────────────────────────────
corr_cols = [
    "NIFTY_Return","Lag_NIFTY_Return","SP500_Return",
    "VIX","Lag_Repo","Lag_Delta_USDINR","COVID_Dummy"
]
corr = df[corr_cols].corr().round(3)
print("\n── Updated Correlation Matrix ──")
print(corr.to_string())

# Check multicollinearity threshold
print("\n── Pairs with |correlation| > 0.75 ──")
flagged = False
cols = corr.columns
for i in range(len(cols)):
    for j in range(i+1, len(cols)):
        val = corr.iloc[i,j]
        if abs(val) > 0.75:
            print(f"  ⚠️  {cols[i]} ↔ {cols[j]} : {val}")
            flagged = True
if not flagged:
    print("  ✅ No multicollinearity concerns")

# ── SAVE ──────────────────────────────────────────────────
raw = pd.read_excel(INPUT, sheet_name="Raw_Data", index_col=0)
adf = pd.read_excel(INPUT, sheet_name="Correlation_Matrix", index_col=0)

with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
    raw.to_excel(writer, sheet_name="Raw_Data", index=True, index_label="Date")
    df.to_excel(writer, sheet_name="Processed_Data", index=True, index_label="Date")
    corr.to_excel(writer, sheet_name="Correlation_Matrix", index=True)

print(f"\n✅ File updated: {OUTPUT}")
print("   Lag_USDINR replaced with Lag_Delta_USDINR (stationary)")
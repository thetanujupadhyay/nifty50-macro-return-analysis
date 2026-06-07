import pandas as pd
import numpy as np
from statsmodels.tsa.stattools import adfuller

# ── LOAD DATA ────────────────────────────────────────────
df = pd.read_excel("nifty_macro_processed.xlsx",
                   sheet_name="Processed_Data", index_col=0)
df.index = pd.to_datetime(df.index)

print("=" * 65)
print("  PHASE 3 — AUGMENTED DICKEY-FULLER STATIONARITY TESTS")
print("=" * 65)
print()
print("  H0 (Null Hypothesis)  : Series HAS a unit root → Non-stationary")
print("  H1 (Alternative)      : Series is Stationary")
print("  Decision rule         : Reject H0 if p-value < 0.05")
print("=" * 65)


# ── ADF TEST FUNCTION ────────────────────────────────────
def run_adf(series, name):
    clean = series.dropna()
    result = adfuller(clean, autolag="AIC")
    adf_stat  = result[0]
    p_value   = result[1]
    n_lags    = result[2]
    crit_1pct = result[4]["1%"]
    crit_5pct = result[4]["5%"]

    if p_value < 0.01:
        decision = "✅ STATIONARY (p<0.01)"
    elif p_value < 0.05:
        decision = "✅ STATIONARY (p<0.05)"
    elif p_value < 0.10:
        decision = "⚠️  BORDERLINE  (p<0.10)"
    else:
        decision = "❌ NON-STATIONARY"

    print(f"\n  Variable     : {name}")
    print(f"  ADF Statistic: {adf_stat:.4f}")
    print(f"  p-value      : {p_value:.4f}")
    print(f"  Lags used    : {n_lags}")
    print(f"  Crit. Values : 1% = {crit_1pct:.3f} | 5% = {crit_5pct:.3f}")
    print(f"  Decision     : {decision}")
    print("  " + "-" * 60)

    return {
        "Variable": name,
        "ADF Statistic": round(adf_stat, 4),
        "p-value": round(p_value, 4),
        "Lags": n_lags,
        "Stationary": "Yes" if p_value < 0.05 else "No",
        "Decision": decision
    }


# ── RUN TESTS ON ALL VARIABLES ───────────────────────────
variables = {
    "NIFTY_Return"     : df["NIFTY_Return"],
    "SP500_Return"     : df["SP500_Return"],
    "VIX"              : df["VIX"],
    "Repo_Rate (level)": df["Repo_Rate"],
    "USDINR (level)"   : df["USDINR"],
    "Lag_Repo"         : df["Lag_Repo"],
    "Lag_USDINR"       : df["Lag_USDINR"],
}

results = []
for name, series in variables.items():
    res = run_adf(series, name)
    results.append(res)


# ── IF NON-STATIONARY: TEST FIRST DIFFERENCE ─────────────
print("\n" + "=" * 65)
print("  FIRST DIFFERENCES — for any non-stationary level variables")
print("=" * 65)

diff_vars = {
    "Δ Repo_Rate (1st diff)" : df["Repo_Rate"].diff().dropna(),
    "Δ USDINR (1st diff)"    : df["USDINR"].diff().dropna(),
}

for name, series in diff_vars.items():
    res = run_adf(series, name)
    results.append(res)


# ── SUMMARY TABLE ─────────────────────────────────────────
print("\n" + "=" * 65)
print("  SUMMARY TABLE")
print("=" * 65)
summary = pd.DataFrame(results)[["Variable","ADF Statistic","p-value","Stationary","Decision"]]
print(summary.to_string(index=False))


# ── EXPORT RESULTS ────────────────────────────────────────
summary.to_excel("phase3_adf_results.xlsx", index=False,
                 sheet_name="ADF_Test_Results")
print("\n✅ Results saved to: phase3_adf_results.xlsx")
print()
print("=" * 65)
print("  NEXT STEP: Review results above before proceeding to")
print("  Phase 4 — VIF Multicollinearity Check")
print("=" * 65)

import pandas as pd
import numpy as np
from statsmodels.stats.outliers_influence import variance_inflation_factor

# ── LOAD FINAL DATA ──────────────────────────────────────
df = pd.read_excel("nifty_macro_final.xlsx",
                   sheet_name="Processed_Data", index_col=0)
df.index = pd.to_datetime(df.index)

print("=" * 60)
print("  PHASE 4 — VIF MULTICOLLINEARITY CHECK")
print("=" * 60)
print()
print("  Rule: VIF = 1       → No correlation with others")
print("        VIF = 1–5     → Acceptable")
print("        VIF = 5–10    → Investigate")
print("        VIF > 10      → Serious multicollinearity → Remove")
print("=" * 60)

# ── DEFINE INDEPENDENT VARIABLES ─────────────────────────
# These are exactly the variables going into regression
indep_vars = [
    "Lag_NIFTY_Return",
    "SP500_Return",
    "VIX",
    "Lag_Repo",
    "Lag_Delta_USDINR",
    "COVID_Dummy"
]

X = df[indep_vars].dropna()

# Add constant (intercept) for VIF calculation
X_with_const = X.copy()
X_with_const.insert(0, "Constant", 1)

# ── COMPUTE VIF ───────────────────────────────────────────
vif_data = []
for i, col in enumerate(X_with_const.columns):
    vif = variance_inflation_factor(X_with_const.values, i)
    vif_data.append({"Variable": col, "VIF": round(vif, 4)})

vif_df = pd.DataFrame(vif_data)

# ── DISPLAY RESULTS ───────────────────────────────────────
print("\n── VIF Results ──")
for _, row in vif_df.iterrows():
    var = row["Variable"]
    v   = row["VIF"]
    if var == "Constant":
        continue
    if v > 10:
        flag = "❌ SERIOUS — consider removing"
    elif v > 5:
        flag = "⚠️  INVESTIGATE"
    else:
        flag = "✅ Acceptable"
    print(f"  {var:<25} VIF = {v:<8.4f}  {flag}")

print()

# ── SUMMARY ───────────────────────────────────────────────
model_vars = vif_df[vif_df["Variable"] != "Constant"]
max_vif = model_vars["VIF"].max()
max_var = model_vars.loc[model_vars["VIF"].idxmax(), "Variable"]

print("=" * 60)
print(f"  Highest VIF : {max_vif:.4f} ({max_var})")
if max_vif < 5:
    print("  Verdict     : ✅ All variables pass — no multicollinearity")
    print("                Model specification is clean.")
    print("                Proceed to Phase 5 — Regression & Forecasting")
elif max_vif < 10:
    print("  Verdict     : ⚠️  Some concern — review flagged variables")
else:
    print("  Verdict     : ❌ Multicollinearity detected — action needed")
print("=" * 60)

# ── EXPORT ────────────────────────────────────────────────
model_vars.to_excel("phase4_vif_results.xlsx",
                    sheet_name="VIF_Results", index=False)
print(f"\n✅ Saved: phase4_vif_results.xlsx")
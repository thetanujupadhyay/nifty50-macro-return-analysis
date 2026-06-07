import pandas as pd
import numpy as np
import statsmodels.api as sm
from statsmodels.stats.diagnostic import het_breuschpagan
from statsmodels.stats.stattools import durbin_watson
import warnings
warnings.filterwarnings("ignore")

df = pd.read_excel("nifty_macro_final.xlsx",
                   sheet_name="Processed_Data", index_col=0)
df.index = pd.to_datetime(df.index)

train = df[df.index <= "2022-12-31"]
Y     = train["NIFTY_Return"]
X_vars = ["Lag_NIFTY_Return","SP500_Return","VIX",
          "Lag_Repo","Lag_Delta_USDINR","COVID_Dummy"]
X = sm.add_constant(train[X_vars])

print("=" * 65)
print("  PHASE 6 — ROBUSTNESS CHECKS")
print("=" * 65)


# ════════════════════════════════════════════════════════
# CHECK 1 — ROBUST STANDARD ERRORS (Newey-West HAC)
# Fixes non-normal residuals & potential heteroskedasticity
# ════════════════════════════════════════════════════════
print("\n── CHECK 1: Robust Standard Errors (Newey-West HAC) ──")
print("   Comparing OLS vs HAC — do significance levels change?")

ols   = sm.OLS(Y, X).fit()
hac   = sm.OLS(Y, X).fit(cov_type="HAC", cov_kwds={"maxlags": 4})

print(f"\n  {'Variable':<25} {'OLS p':>8}  {'HAC p':>8}  {'Change?':>12}")
print("  " + "─" * 57)
for var in X_vars:
    p_ols = ols.pvalues[var]
    p_hac = hac.pvalues[var]
    sig_ols = p_ols < 0.05
    sig_hac = p_hac < 0.05
    if sig_ols and not sig_hac:
        change = "⚠️  LOST significance"
    elif not sig_ols and sig_hac:
        change = "📈 GAINED significance"
    elif sig_ols and sig_hac:
        change = "✅ Stays significant"
    else:
        change = "➡️  Stays insignificant"
    print(f"  {var:<25} {p_ols:>8.4f}  {p_hac:>8.4f}  {change:>12}")

print(f"\n  OLS  Adj-R²: {ols.rsquared_adj:.4f}")
print(f"  HAC  Adj-R²: {hac.rsquared_adj:.4f}  (same — only SEs change)")


# ════════════════════════════════════════════════════════
# CHECK 2 — HETEROSKEDASTICITY TEST (Breusch-Pagan)
# ════════════════════════════════════════════════════════
print("\n\n── CHECK 2: Breusch-Pagan Heteroskedasticity Test ──")
print("   H0: Residuals have constant variance (homoskedastic)")
print("   H1: Variance changes over time (heteroskedastic)")

bp_test = het_breuschpagan(ols.resid, ols.model.exog)
bp_stat, bp_pval = bp_test[0], bp_test[1]
print(f"\n  BP Statistic : {bp_stat:.4f}")
print(f"  p-value      : {bp_pval:.4f}")
if bp_pval < 0.05:
    print("  Result       : ⚠️  Heteroskedasticity detected → HAC SEs are justified")
else:
    print("  Result       : ✅ Homoskedastic — OLS standard errors are valid")


# ════════════════════════════════════════════════════════
# CHECK 3 — STRUCTURAL BREAK (Chow Test: Pre vs Post COVID)
# Split: Pre-COVID (before Mar 2020) vs Post-COVID (after)
# ════════════════════════════════════════════════════════
print("\n\n── CHECK 3: Structural Break — Pre vs Post COVID ──")
print("   Running separate regressions for each period")

core_vars = ["SP500_Return", "Lag_Repo"]   # only significant vars

pre  = train[train.index < "2020-03-01"]
post = train[train.index >= "2020-03-01"]

reg_pre  = sm.OLS(pre["NIFTY_Return"],
                  sm.add_constant(pre[core_vars])).fit()
reg_post = sm.OLS(post["NIFTY_Return"],
                  sm.add_constant(post[core_vars])).fit()

print(f"\n  Period        : {'PRE-COVID':>20}  {'POST-COVID':>20}")
print(f"  Date range    : {'May08–Feb20':>20}  {'Mar20–Dec22':>20}")
print(f"  Observations  : {len(pre):>20}  {len(post):>20}")
print(f"  Adj R²        : {reg_pre.rsquared_adj:>20.4f}  {reg_post.rsquared_adj:>20.4f}")
print()
for var in ["const"] + core_vars:
    cp = reg_pre.params[var]
    pp = reg_pre.pvalues[var]
    ca = reg_post.params[var]
    pa = reg_post.pvalues[var]
    print(f"  {var:<15} coef: {cp:>+8.4f} (p={pp:.3f})   coef: {ca:>+8.4f} (p={pa:.3f})")

print("\n  Interpretation:")
sp_change = reg_post.params["SP500_Return"] - reg_pre.params["SP500_Return"]
print(f"  SP500 coefficient change: {sp_change:+.4f} "
      f"({'stronger' if sp_change > 0 else 'weaker'} post-COVID)")


# ════════════════════════════════════════════════════════
# CHECK 4 — ROLLING WINDOW REGRESSION (60-month window)
# Tests if SP500 coefficient is stable over time
# ════════════════════════════════════════════════════════
print("\n\n── CHECK 4: Rolling Window Regression (60-month window) ──")
print("   Tracking SP500_Return coefficient over time")
print("   If stable → model is robust | If drifting → regime sensitivity")

window = 60
roll_dates, roll_coefs_sp500, roll_coefs_repo = [], [], []

for end in range(window, len(train)+1):
    sub = train.iloc[end-window:end]
    try:
        mod = sm.OLS(sub["NIFTY_Return"],
                     sm.add_constant(sub[core_vars])).fit()
        roll_dates.append(sub.index[-1])
        roll_coefs_sp500.append(mod.params["SP500_Return"])
        roll_coefs_repo.append(mod.params["Lag_Repo"])
    except:
        pass

roll_df = pd.DataFrame({
    "Date":          roll_dates,
    "SP500_coef":    roll_coefs_sp500,
    "Lag_Repo_coef": roll_coefs_repo
})

print(f"\n  SP500 coefficient over rolling 60m windows:")
print(f"  Min  : {min(roll_coefs_sp500):.4f}")
print(f"  Max  : {max(roll_coefs_sp500):.4f}")
print(f"  Mean : {np.mean(roll_coefs_sp500):.4f}")
print(f"  Std  : {np.std(roll_coefs_sp500):.4f}")
print(f"  Range: {max(roll_coefs_sp500)-min(roll_coefs_sp500):.4f}")

if np.std(roll_coefs_sp500) > 0.3:
    print("  ⚠️  High variation — SP500 relationship is regime-sensitive")
else:
    print("  ✅ Relatively stable — SP500 relationship is consistent")

print(f"\n  Sample of rolling coefficients (every 12 months):")
print(f"  {'Date':<15} {'SP500 coef':>12}  {'Repo coef':>12}")
for _, row in roll_df.iloc[::12].iterrows():
    print(f"  {str(row['Date'].date()):<15} {row['SP500_coef']:>12.4f}  {row['Lag_Repo_coef']:>12.4f}")


# ════════════════════════════════════════════════════════
# EXPORT
# ════════════════════════════════════════════════════════
hac_results = pd.DataFrame({
    "Variable":   X_vars,
    "OLS_coef":   [ols.params[v] for v in X_vars],
    "OLS_pval":   [ols.pvalues[v] for v in X_vars],
    "HAC_pval":   [hac.pvalues[v] for v in X_vars],
    "Sig_OLS":    ["Yes" if ols.pvalues[v]<0.05 else "No" for v in X_vars],
    "Sig_HAC":    ["Yes" if hac.pvalues[v]<0.05 else "No" for v in X_vars],
})

split_results = pd.DataFrame({
    "Variable":   ["const"] + core_vars,
    "Pre_coef":   [reg_pre.params[v]  for v in ["const"]+core_vars],
    "Pre_pval":   [reg_pre.pvalues[v] for v in ["const"]+core_vars],
    "Post_coef":  [reg_post.params[v]  for v in ["const"]+core_vars],
    "Post_pval":  [reg_post.pvalues[v] for v in ["const"]+core_vars],
})

with pd.ExcelWriter("phase6_robustness_results.xlsx", engine="openpyxl") as w:
    hac_results.to_excel(w, sheet_name="HAC_vs_OLS",        index=False)
    split_results.to_excel(w, sheet_name="PrePost_COVID",   index=False)
    roll_df.to_excel(w, sheet_name="Rolling_Coefficients",  index=False)

print(f"\n\n✅ Saved: phase6_robustness_results.xlsx")
print("   Sheets: HAC_vs_OLS | PrePost_COVID | Rolling_Coefficients")
print("=" * 65)
print("  ROBUSTNESS CHECKS COMPLETE")
print("  After this → Final interpretation & report structure")
print("=" * 65)
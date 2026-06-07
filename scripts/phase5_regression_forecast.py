import pandas as pd
import numpy as np
import statsmodels.api as sm
from sklearn.metrics import mean_squared_error, mean_absolute_error
import warnings
warnings.filterwarnings("ignore")

# ── LOAD DATA ────────────────────────────────────────────
df = pd.read_excel("nifty_macro_final.xlsx",
                   sheet_name="Processed_Data", index_col=0)
df.index = pd.to_datetime(df.index)

# ── DEFINE VARIABLES ─────────────────────────────────────
Y_col    = "NIFTY_Return"
m1_vars  = ["Lag_NIFTY_Return"]                        # Model 1: baseline
m2_vars  = ["Lag_NIFTY_Return", "SP500_Return", "VIX", # Model 2: full macro
            "Lag_Repo", "Lag_Delta_USDINR", "COVID_Dummy"]

# ── TRAIN / TEST SPLIT ───────────────────────────────────
train = df[df.index <= "2022-12-31"]
test  = df[df.index >  "2022-12-31"]

print("=" * 65)
print("  PHASE 5 — REGRESSION & FORECASTING")
print("=" * 65)
print(f"  Train: {train.index[0].date()} → {train.index[-1].date()} ({len(train)} months)")
print(f"  Test : {test.index[0].date()}  → {test.index[-1].date()} ({len(test)} months)")
print("=" * 65)


# ════════════════════════════════════════════════════════
# STEP B — AUTOCORRELATION REGRESSION
# Q: Does last month's NIFTY return predict this month's?
# ════════════════════════════════════════════════════════
print("\n" + "─" * 65)
print("  STEP B — AUTOCORRELATION TEST")
print("  Dependent: NIFTY_Return")
print("  Independent: Lag_NIFTY_Return only")
print("─" * 65)

Y_train = train[Y_col]
X_auto  = sm.add_constant(train[["Lag_NIFTY_Return"]])
model_auto = sm.OLS(Y_train, X_auto).fit()

print(model_auto.summary())

print("\n── KEY TAKEAWAYS (Step B) ──")
coef = model_auto.params["Lag_NIFTY_Return"]
pval = model_auto.pvalues["Lag_NIFTY_Return"]
dw   = sm.stats.durbin_watson(model_auto.resid)
print(f"  Lag_NIFTY coefficient : {coef:.4f}")
print(f"  p-value               : {pval:.4f}  {'✅ Significant' if pval < 0.05 else '❌ Not significant'}")
print(f"  R²                    : {model_auto.rsquared:.4f}")
print(f"  Durbin-Watson         : {dw:.4f}  (ideal ≈ 2.0)")


# ════════════════════════════════════════════════════════
# STEP A — FULL MACRO REGRESSION
# Q: Do macro variables explain NIFTY returns?
# ════════════════════════════════════════════════════════
print("\n" + "─" * 65)
print("  STEP A — FULL MACRO REGRESSION")
print("  Dependent: NIFTY_Return")
print("  Independent: All macro + global + dummy variables")
print("─" * 65)

X_macro = sm.add_constant(train[m2_vars])
model_macro = sm.OLS(Y_train, X_macro).fit()

print(model_macro.summary())

print("\n── KEY TAKEAWAYS (Step A) ──")
dw2 = sm.stats.durbin_watson(model_macro.resid)
print(f"  Adjusted R²    : {model_macro.rsquared_adj:.4f}")
print(f"  F-statistic    : {model_macro.fvalue:.4f}  (p={model_macro.f_pvalue:.4f})")
print(f"  Durbin-Watson  : {dw2:.4f}")
print()
for var in m2_vars:
    c = model_macro.params[var]
    p = model_macro.pvalues[var]
    sig = "✅ Significant" if p < 0.05 else ("⚠️ Borderline" if p < 0.10 else "❌ Not significant")
    print(f"  {var:<25}  coef={c:+.4f}  p={p:.4f}  {sig}")


# ════════════════════════════════════════════════════════
# STEP C — OUT-OF-SAMPLE FORECASTING
# Train on 2008-2022, forecast 2023-2025
# Compare Model 1 vs Model 2
# ════════════════════════════════════════════════════════
print("\n" + "─" * 65)
print("  STEP C — OUT-OF-SAMPLE FORECAST COMPARISON")
print("  Forecasting 2023–2025 (36 months)")
print("─" * 65)

Y_test = test[Y_col]

# ── Model 1: Baseline (Lag only) ─────────────────────────
X_train_m1 = sm.add_constant(train[m1_vars])
X_test_m1  = sm.add_constant(test[m1_vars])
m1_fit     = sm.OLS(Y_train, X_train_m1).fit()
m1_pred    = m1_fit.predict(X_test_m1)

rmse_m1 = np.sqrt(mean_squared_error(Y_test, m1_pred))
mae_m1  = mean_absolute_error(Y_test, m1_pred)

# ── Model 2: Full Macro ───────────────────────────────────
X_train_m2 = sm.add_constant(train[m2_vars])
X_test_m2  = sm.add_constant(test[m2_vars])
m2_fit     = sm.OLS(Y_train, X_train_m2).fit()
m2_pred    = m2_fit.predict(X_test_m2)

rmse_m2 = np.sqrt(mean_squared_error(Y_test, m2_pred))
mae_m2  = mean_absolute_error(Y_test, m2_pred)

# ── Naive benchmark: predict mean return every month ─────
mean_pred  = pd.Series(Y_train.mean(), index=Y_test.index)
rmse_naive = np.sqrt(mean_squared_error(Y_test, mean_pred))
mae_naive  = mean_absolute_error(Y_test, mean_pred)

# ── Results table ─────────────────────────────────────────
print(f"\n  {'Model':<35} {'RMSE':>8}  {'MAE':>8}")
print(f"  {'─'*51}")
print(f"  {'Naive (predict mean always)':<35} {rmse_naive:>8.4f}  {mae_naive:>8.4f}")
print(f"  {'Model 1: Baseline (Lag only)':<35} {rmse_m1:>8.4f}  {mae_m1:>8.4f}")
print(f"  {'Model 2: Full Macro':<35} {rmse_m2:>8.4f}  {mae_m2:>8.4f}")
print()

# ── Improvement from M1 to M2 ────────────────────────────
rmse_improvement = ((rmse_m1 - rmse_m2) / rmse_m1) * 100
mae_improvement  = ((mae_m1  - mae_m2)  / mae_m1)  * 100
print(f"  Model 2 vs Model 1 improvement:")
print(f"  RMSE change : {rmse_improvement:+.2f}%  {'✅ Macro adds value' if rmse_improvement > 2 else '➡️ Marginal improvement' if rmse_improvement > 0 else '❌ No improvement'}")
print(f"  MAE  change : {mae_improvement:+.2f}%  {'✅ Macro adds value' if mae_improvement > 2 else '➡️ Marginal improvement' if mae_improvement > 0 else '❌ No improvement'}")

# ── Actual vs Predicted table (test period) ──────────────
forecast_df = pd.DataFrame({
    "Actual":    Y_test.values,
    "Model1_Pred": m1_pred.values,
    "Model2_Pred": m2_pred.values,
    "M1_Error":  (Y_test.values - m1_pred.values),
    "M2_Error":  (Y_test.values - m2_pred.values),
}, index=Y_test.index)

print(f"\n── Actual vs Predicted (2023–2025) ──")
print(forecast_df.round(4).to_string())

# ════════════════════════════════════════════════════════
# EXPORT ALL RESULTS
# ════════════════════════════════════════════════════════
summary_data = {
    "Metric": ["RMSE", "MAE"],
    "Naive":    [round(rmse_naive,4), round(mae_naive,4)],
    "Model1_Baseline": [round(rmse_m1,4), round(mae_m1,4)],
    "Model2_FullMacro": [round(rmse_m2,4), round(mae_m2,4)],
    "M2_vs_M1_improvement_%": [round(rmse_improvement,2), round(mae_improvement,2)]
}

coef_data = pd.DataFrame({
    "Variable":    ["const"] + m2_vars,
    "Coefficient": model_macro.params.values.round(4),
    "Std_Error":   model_macro.bse.values.round(4),
    "t_stat":      model_macro.tvalues.values.round(4),
    "p_value":     model_macro.pvalues.values.round(4),
    "Significant": ["Yes" if p < 0.05 else "No"
                    for p in model_macro.pvalues.values]
})

with pd.ExcelWriter("phase5_results.xlsx", engine="openpyxl") as writer:
    coef_data.to_excel(writer, sheet_name="Macro_Regression", index=False)
    forecast_df.to_excel(writer, sheet_name="Forecast_Comparison")
    pd.DataFrame(summary_data).to_excel(writer, sheet_name="RMSE_MAE_Summary", index=False)

print(f"\n✅ Saved: phase5_results.xlsx")
print("   Sheets: Macro_Regression | Forecast_Comparison | RMSE_MAE_Summary")
print("=" * 65)
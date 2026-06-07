import pandas as pd
import numpy as np
import statsmodels.api as sm
from sklearn.metrics import mean_squared_error, mean_absolute_error
import warnings
warnings.filterwarnings("ignore")

df = pd.read_excel("nifty_macro_final.xlsx",
                   sheet_name="Processed_Data", index_col=0)
df.index = pd.to_datetime(df.index)

train = df[df.index <= "2022-12-31"]
test  = df[df.index >  "2022-12-31"]
Y_train = train["NIFTY_Return"]
Y_test  = test["NIFTY_Return"]

print("=" * 65)
print("  PHASE 5b — IMPROVED FORECAST (significant vars only)")
print("=" * 65)

# ── Model 1: Baseline ─────────────────────────────────────
m1 = sm.OLS(Y_train, sm.add_constant(train[["Lag_NIFTY_Return"]])).fit()
m1_pred = m1.predict(sm.add_constant(test[["Lag_NIFTY_Return"]]))

# ── Model 2: Full Macro (original) ───────────────────────
m2_vars = ["Lag_NIFTY_Return","SP500_Return","VIX",
           "Lag_Repo","Lag_Delta_USDINR","COVID_Dummy"]
m2 = sm.OLS(Y_train, sm.add_constant(train[m2_vars])).fit()
m2_pred = m2.predict(sm.add_constant(test[m2_vars]))

# ── Model 3: Significant vars only (SP500 + Lag_Repo) ────
# Only keep variables with p < 0.05 from Step A
m3_vars = ["SP500_Return", "Lag_Repo"]
m3 = sm.OLS(Y_train, sm.add_constant(train[m3_vars])).fit()
m3_pred = m3.predict(sm.add_constant(test[m3_vars]))

print("\nModel 3 summary (significant vars only):")
print(f"  Adj R² (train): {m3.rsquared_adj:.4f}")
for v in m3_vars:
    print(f"  {v:<20} coef={m3.params[v]:+.4f}  p={m3.pvalues[v]:.4f}")

# ── Naive benchmark ───────────────────────────────────────
naive_pred = pd.Series(Y_train.mean(), index=Y_test.index)

# ── Compare all models ────────────────────────────────────
def metrics(actual, pred):
    rmse = np.sqrt(mean_squared_error(actual, pred))
    mae  = mean_absolute_error(actual, pred)
    return rmse, mae

rmse_n, mae_n   = metrics(Y_test, naive_pred)
rmse_m1, mae_m1 = metrics(Y_test, m1_pred)
rmse_m2, mae_m2 = metrics(Y_test, m2_pred)
rmse_m3, mae_m3 = metrics(Y_test, m3_pred)

print(f"\n{'Model':<40} {'RMSE':>8}  {'MAE':>8}  {'vs M1 RMSE':>12}")
print("─" * 72)
print(f"{'Naive (mean always)':<40} {rmse_n:>8.4f}  {mae_n:>8.4f}")
print(f"{'Model 1: Lag only (baseline)':<40} {rmse_m1:>8.4f}  {mae_m1:>8.4f}  {'—':>12}")
print(f"{'Model 2: Full macro (6 vars)':<40} {rmse_m2:>8.4f}  {mae_m2:>8.4f}  {((rmse_m1-rmse_m2)/rmse_m1*100):>+11.2f}%")
print(f"{'Model 3: Sig. vars only (SP500+Repo)':<40} {rmse_m3:>8.4f}  {mae_m3:>8.4f}  {((rmse_m1-rmse_m3)/rmse_m1*100):>+11.2f}%")

# ── Save ──────────────────────────────────────────────────
results = pd.DataFrame({
    "Model":    ["Naive","Model1_Lag","Model2_FullMacro","Model3_SigVars"],
    "Variables": ["None","Lag_NIFTY","6 macro vars","SP500+Lag_Repo"],
    "RMSE":     [round(rmse_n,4), round(rmse_m1,4), round(rmse_m2,4), round(rmse_m3,4)],
    "MAE":      [round(mae_n,4),  round(mae_m1,4),  round(mae_m2,4),  round(mae_m3,4)],
})
results.to_excel("phase5b_model_comparison.xlsx", index=False)
print(f"\n✅ Saved: phase5b_model_comparison.xlsx")
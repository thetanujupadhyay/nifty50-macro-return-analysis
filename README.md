# 📈 Macroeconomic Determinants and Forecastability of NIFTY 50 Returns
### A Longitudinal Time-Series Econometric Study (2008–2025)

![Python](https://img.shields.io/badge/Python-3.11+-blue?logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)
![Status](https://img.shields.io/badge/Status-Complete-brightgreen)
![Domain](https://img.shields.io/badge/Domain-Quantitative%20Finance-navy)

> **Can publicly observable macroeconomic data predict India's NIFTY 50 equity returns?**  
> This study provides a rigorous, data-driven answer — and the finding has direct implications for investment strategy.

---

## 🎯 Business Problem

Every month, fund managers and analysts scan macroeconomic signals — RBI rate decisions, rupee movements, US market trends — trying to gauge where Indian equity markets are headed. But does this actually work?

This project answers three concrete questions:
1. Do past NIFTY returns predict future returns? *(Weak-Form EMH Test)*
2. Which macroeconomic variables genuinely drive NIFTY 50 returns?
3. Does macro data improve actual out-of-sample return forecasts?

The answers carry real implications: if macro variables forecast returns, institutional strategies can be built around them. If not, it confirms market efficiency — and warns against macro-based market-timing.

---

## 🔑 Key Findings

| Finding | Result | Implication |
|---|---|---|
| Autocorrelation (Weak-Form EMH) | p = 0.632, R² = 0.001 | ✅ Weak-Form EMH **supported** — no momentum edge |
| S&P 500 Return | β = +0.820, p < 0.001 | 🌐 Dominant driver — India's global equity beta |
| RBI Repo Rate (lagged) | β = −0.007, p = 0.019 (HAC) | 🏦 Monetary policy transmits to equities |
| Out-of-Sample Forecast | Macro model 19.4% **worse** than baseline | ✅ Semi-Strong EMH **confirmed** |

> **Bottom line:** Macro variables explain historical NIFTY returns (Adj. R² = 0.436) but offer no persistent forecasting advantage — consistent with efficient markets.

---

## 📊 Results at a Glance

```
Training Period : May 2008 – Dec 2022  (176 observations)
Test Period     : Jan 2023 – Dec 2025  (36 observations)
Total Sample    : 212 monthly observations across 4 market regimes

Model B (Full Macro OLS):
  Adjusted R²     = 0.436
  F-statistic     = 23.52  (p < 0.001)
  Durbin-Watson   = 1.975  (no serial correlation)
  Breusch-Pagan   = 31.09  (p < 0.001 → HAC correction applied)

Out-of-Sample RMSE Comparison:
  Naïve Baseline  = 0.0321  ← benchmark
  Model 1         = 0.0320  ← best forecast model
  Model 3         = 0.0375  ← 17.2% worse
  Model 2 (Full)  = 0.0382  ← 19.4% worse
```

---

## 🏗️ Project Architecture

```
6-Phase Pipeline — each phase produces a verified Excel output before proceeding

Phase 1  →  data_collection.py          →  nifty_macro_raw.xlsx
Phase 2  →  phase2_processing.py        →  nifty_macro_processed.xlsx
Phase 3  →  phase3_stationarity.py      →  phase3_adf_results.xlsx
           phase3b_fix_usdinr.py        →  nifty_macro_final.xlsx  ✦ master dataset
Phase 4  →  phase4_vif.py               →  phase4_vif_results.xlsx
Phase 5  →  phase5_regression_forecast.py    →  phase5_results.xlsx
           phase5b_improved_forecast.py →  phase5b_model_comparison.xlsx
Phase 6  →  phase6_robustness.py        →  phase6_robustness_results.xlsx
           build_dashboard.py          →  NIFTY50_TSA_Dashboard.xlsx  ✦ final output
```

---

## 📁 Repository Structure

```
nifty50-macro-return-analysis/
│
├── scripts/
│   ├── data_collection.py
│   ├── phase2_processing.py
│   ├── phase3_stationarity.py
│   ├── phase3b_fix_usdinr.py
│   ├── phase4_vif.py
│   ├── phase5_regression_forecast.py
│   ├── phase5b_improved_forecast.py
│   ├── phase6_robustness.py
│   └── build_dashboard.py
│
├── data/
│   ├── nifty_macro_raw.xlsx
│   ├── nifty_macro_processed.xlsx
│   └── nifty_macro_final.xlsx          ← master analytical dataset (212 rows)
│
├── results/
│   ├── phase3_adf_results.xlsx
│   ├── phase4_vif_results.xlsx
│   ├── phase5_results.xlsx
│   ├── phase5b_model_comparison.xlsx
│   ├── phase6_robustness_results.xlsx
│   └── NIFTY50_TSA_Dashboard.xlsx      ← 7-sheet results dashboard
│
├── figures/
│   ├── fig1_1_nifty_price.svg
│   ├── fig1_2_problem_statement.svg
│   ├── fig2_1_architecture.svg
│   ├── fig3_4_forecast.svg
│   └── fig3_5_rolling.svg
│
├── requirements.txt
└── README.md
```

---

## ⚙️ Variables Studied

| Variable | Source | Transform | Role |
|---|---|---|---|
| NIFTY 50 Close | NSE via yfinance (`^NSEI`) | Log return | **Dependent variable** |
| S&P 500 Return | Yahoo Finance (`^GSPC`) | Log return | Global market proxy |
| India VIX | NSE via yfinance (`^INDIAVIX`) | As-is | Fear / risk index |
| RBI Repo Rate | RBI DBIE portal (CSV) | Forward-fill + lag | Monetary policy |
| USD/INR Rate | Yahoo Finance (`INR=X`) | 1st difference + lag | Exchange rate |
| COVID-19 Dummy | Constructed | Binary (Mar 2020–Dec 2021) | Structural break |

---

## 🚀 How to Run

### 1. Clone the repository
```bash
git clone https://github.com/YOUR_USERNAME/nifty50-macro-return-analysis.git
cd nifty50-macro-return-analysis
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the pipeline (execute in order)
```bash
python scripts/data_collection.py
python scripts/phase2_processing.py
python scripts/phase3_stationarity.py
python scripts/phase3b_fix_usdinr.py
python scripts/phase4_vif.py
python scripts/phase5_regression_forecast.py
python scripts/phase5b_improved_forecast.py
python scripts/phase6_robustness.py
python scripts/build_dashboard.py
```

> ⏱️ Total runtime: ~5 minutes on standard hardware with internet connection.  
> 📡 Phase 1 requires an active internet connection to fetch data from Yahoo Finance.

---

## 📦 Requirements

```
yfinance>=0.2.0
pandas>=2.0.0
numpy>=1.24.0
statsmodels>=0.14.0
scikit-learn>=1.3.0
openpyxl>=3.1.0
```

---

## 📐 Methodology Highlights

- **Stationarity:** ADF test on all variables — USD/INR confirmed non-stationary (p=0.818), corrected via first differencing
- **Multicollinearity:** VIF check on all predictors — max VIF = 1.58, well below threshold of 5.0
- **Robust Inference:** Newey-West HAC standard errors (validated by Breusch-Pagan, BP=31.09, p<0.001)
- **No Look-Ahead Bias:** All predictors are one-period lagged before model estimation
- **Genuine Holdout:** 36-month test set (2023–2025) strictly separated from training — never used in any estimation step
- **Structural Robustness:** Pre/post COVID sub-period analysis + 117 rolling 60-month window regressions

---

## 💡 Business & Investment Implications

**For portfolio managers:**
The S&P 500 beta of 0.82 quantifies NIFTY's global equity risk exposure. Hedging this factor is more effective than monitoring domestic macro variables for short-term risk management.

**For macro strategists:**
The RBI Repo Rate significance (lagged by one month) confirms monetary policy transmission to equity markets — a validation of the discount-rate channel relevant for rate-change event studies.

**For quant researchers:**
The in-sample vs out-of-sample divergence provides empirical evidence for Semi-Strong EMH in Indian markets — a benchmark result for strategy backtesting validity.

---

## 🔭 Future Extensions

- [ ] GARCH(1,1) modelling for time-varying volatility (BP test confirms heteroskedasticity)
- [ ] Vector Autoregression (VAR) for bidirectional macro-equity feedback
- [ ] LASSO / Random Forest ML comparison on same train-test split
- [ ] Extend variable set: FII flows, crude oil, Google Trends uncertainty index
- [ ] Sector-level analysis: Banking, IT, Energy, FMCG indices

---

## 👥 Authors

**Anushka Trivedi** — B.Tech CSE (Data Science), NIET Greater Noida  
**Tanuj Upadhyay** — B.Tech CSE (Data Science), NIET Greater Noida  

*Capstone Project | Department of Computer Science and Engineering | Session 2025-26*  
*Supervised by Dr. Ashish Chakraverti & Dr. Nidhi Sharma, NIET Greater Noida*

---

## 📄 License

This project is licensed under the MIT License — free to use, fork, and build upon with attribution.

---

<p align="center">
  <i>If this project helped you, consider giving it a ⭐ — it helps others find it.</i>
</p>

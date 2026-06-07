import yfinance as yf
import pandas as pd

# ── CONFIG ──────────────────────────────────────────────
START       = "2008-01-01"
END         = "2025-12-31"
REPO_FILE   = "repo_rate.csv"    # Only manual file needed
OUTPUT      = "nifty_macro_raw.xlsx"
# ────────────────────────────────────────────────────────


# ── 1. NIFTY 50 ──────────────────────────────────────────
print("Fetching NIFTY 50...")
nifty = yf.download("^NSEI", start=START, end=END, interval="1mo", auto_adjust=True)[["Close"]]
nifty.columns = ["NIFTY_Close"]
nifty.index = nifty.index.to_period("M").to_timestamp("M")


# ── 2. S&P 500 ───────────────────────────────────────────
print("Fetching S&P 500...")
sp = yf.download("^GSPC", start=START, end=END, interval="1mo", auto_adjust=True)[["Close"]]
sp.columns = ["SP500_Close"]
sp.index = sp.index.to_period("M").to_timestamp("M")


# ── 3. USD/INR ───────────────────────────────────────────
print("Fetching USD/INR...")
fx = yf.download("INR=X", start=START, end=END, interval="1mo", auto_adjust=True)[["Close"]]
fx.columns = ["USDINR"]
fx.index = fx.index.to_period("M").to_timestamp("M")


# ── 4. INDIA VIX ─────────────────────────────────────────
print("Fetching India VIX...")
vix = yf.download("^INDIAVIX", start=START, end=END, interval="1mo", auto_adjust=True)[["Close"]]
vix.columns = ["VIX"]
vix.index = vix.index.to_period("M").to_timestamp("M")


# ── 5. REPO RATE (from RBI CSV) ──────────────────────────
print("Reading Repo Rate from CSV...")

# Skip the 10 junk header rows; no header — assign names manually
repo_raw = pd.read_csv(
    "repo_rate.csv",
    skiprows=10,          # skip all junk rows above actual data
    header=None,
    names=["Date","Bank_Rate","Repo","Rev_Repo","SDF","MSF","CRR","SLR","Extra"],
    na_values=["-", " ", ""]
)

# Drop the last empty row and any fully null rows
repo_raw = repo_raw.dropna(subset=["Date"])
repo_raw = repo_raw[repo_raw["Date"].astype(str).str.strip() != ""]

# Parse date
repo_raw["Date"] = pd.to_datetime(repo_raw["Date"].str.strip(), dayfirst=True, errors="coerce")
repo_raw = repo_raw.dropna(subset=["Date"])

# Convert Repo to numeric
repo_raw["Repo"] = pd.to_numeric(repo_raw["Repo"], errors="coerce")

# Sort ascending by date
repo_raw = repo_raw.sort_values("Date").set_index("Date")

print("\n── Repo Rate Preview (cleaned) ──")
print(repo_raw[["Repo"]].dropna().head(10))

# Forward-fill to get a value for every month-end
full_monthly_index = pd.date_range(start=START, end=END, freq="ME")
repo_monthly = repo_raw[["Repo"]].reindex(
    repo_raw.index.union(full_monthly_index)
).ffill().reindex(full_monthly_index)
repo_monthly.columns = ["Repo_Rate"]

print(f"\n✔ Repo Rate monthly rows: {repo_monthly.shape[0]}")
print(repo_monthly.head(10))


# ── 6. MERGE ALL ON MONTH-END DATE ───────────────────────
print("\nMerging all datasets...")
df = nifty.join([sp, fx, vix, repo_monthly], how="outer")
df = df.loc["2008-01-31":"2025-12-31"]
df = df.sort_index()

# ── 7. QUALITY CHECK ─────────────────────────────────────
print("\n── Shape ──")
print(f"Rows: {df.shape[0]}, Columns: {df.shape[1]}")
print("\n── Missing Values ──")
print(df.isnull().sum())
print("\n── First 5 Rows ──")
print(df.head())
print("\n── Last 5 Rows ──")
print(df.tail())


# ── 8. EXPORT ────────────────────────────────────────────
df.to_excel(OUTPUT, sheet_name="Raw_Data", index=True, index_label="Date")
print(f"\n✅ File saved: {OUTPUT}")
print("   Ready for Phase 2 — Excel Processing!")


#Handling Missing Values
FILE = "nifty_macro_raw.xlsx"

df = pd.read_excel(FILE, sheet_name="Raw_Data", index_col=0)

# Repo rate was 7.75% from Jan 2008, raised to 8.00% in June 2008
# Fill the 5 missing months (Jan–May 2008) with 7.75
df.loc[df.index < "2008-06-01", "Repo_Rate"] = df.loc[
    df.index < "2008-06-01", "Repo_Rate"
].fillna(7.75)

print("Missing values after fix:")
print(df.isnull().sum())
print("\nFirst 8 rows:")
print(df.head(8))

# Save back
df.to_excel(FILE, sheet_name="Raw_Data", index=True, index_label="Date")
print("\n✅ File patched and saved!")
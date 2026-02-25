"""
analysis.py
===========
Analyses three bank-account Excel files:
  - cba.xlsx
  - PP25-26.xlsx
  - smartaccess.xlsx

Each file contains at minimum the columns:
    Date | Trans_Details | Trans_Dr | Trans_Cr

The script:
1. Loads each file into a separate DataFrame and adds a **TOTAL** row.
2. Groups *similar* transaction descriptions across all three sources using
   fuzzy string matching (rapidfuzz) and builds a consolidated summary
   DataFrame with totals.
3. Saves all DataFrames to CSV files in the 'output/' directory.
4. Produces a rich set of charts (bar, pie, stacked-bar, line) and saves
   them as PNG images in 'output/'.

Usage:
    python3 analysis.py

Requirements:
    pip install pandas openpyxl matplotlib rapidfuzz
"""

import os
import warnings
import pandas as pd
import matplotlib
matplotlib.use("Agg")          # non-interactive backend – safe for all envs
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from rapidfuzz import process, fuzz

warnings.filterwarnings("ignore")

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

FILES = {
    "CBA": "cba.xlsx",
    "PP25-26": "PP25-26.xlsx",
    "SmartAccess": "smartaccess.xlsx",
}

# Similarity threshold (0-100).  Descriptions scoring >= this are grouped.
SIMILARITY_THRESHOLD = 75

# Colour palette (one colour per source)
SOURCE_COLOURS = {
    "CBA": "#1f77b4",
    "PP25-26": "#ff7f0e",
    "SmartAccess": "#2ca02c",
}

# ---------------------------------------------------------------------------
# Helper utilities
# ---------------------------------------------------------------------------

def load_file(path: str, source_name: str) -> pd.DataFrame:
    """Load an Excel file and normalise the required columns."""
    df = pd.read_excel(path)
    df.columns = [c.strip() for c in df.columns]

    # Ensure mandatory columns exist
    for col in ("Date", "Trans_Details", "Trans_Dr", "Trans_Cr"):
        if col not in df.columns:
            raise ValueError(
                f"File '{path}' is missing column '{col}'. "
                f"Found columns: {list(df.columns)}"
            )

    df = df[["Date", "Trans_Details", "Trans_Dr", "Trans_Cr"]].copy()
    df["Trans_Dr"] = pd.to_numeric(df["Trans_Dr"], errors="coerce")
    df["Trans_Cr"] = pd.to_numeric(df["Trans_Cr"], errors="coerce")
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Source"] = source_name
    return df


def add_totals_row(df: pd.DataFrame) -> pd.DataFrame:
    """Append a TOTAL row that sums Trans_Dr and Trans_Cr."""
    total_dr = df["Trans_Dr"].sum()
    total_cr = df["Trans_Cr"].sum()
    total_row = pd.DataFrame(
        [{"Date": "TOTAL", "Trans_Details": "— TOTAL —",
          "Trans_Dr": round(total_dr, 2), "Trans_Cr": round(total_cr, 2)}]
    )
    return pd.concat([df.drop(columns=["Source"], errors="ignore"),
                      total_row], ignore_index=True)


def fmt_currency(val):
    """Format a float as a currency string, blank for NaN."""
    if pd.isna(val):
        return ""
    return f"${val:,.2f}"


def save_df_csv(df: pd.DataFrame, name: str):
    """Save a DataFrame (with TOTAL row) to CSV."""
    path = os.path.join(OUTPUT_DIR, f"{name}.csv")
    display_df = df.copy()
    for col in ("Trans_Dr", "Trans_Cr"):
        if col in display_df.columns:
            display_df[col] = display_df[col].apply(fmt_currency)
    display_df.to_csv(path, index=False)
    print(f"  Saved {path}")


# ---------------------------------------------------------------------------
# Step 1 – Load each source file
# ---------------------------------------------------------------------------
print("=" * 60)
print("STEP 1 – Loading source files")
print("=" * 60)

raw_frames: dict[str, pd.DataFrame] = {}
for name, filepath in FILES.items():
    if not os.path.exists(filepath):
        print(f"  WARNING: '{filepath}' not found – skipping.")
        continue
    raw_frames[name] = load_file(filepath, name)
    print(f"  Loaded {filepath}  ({len(raw_frames[name])} rows)")

if not raw_frames:
    raise SystemExit("No source files could be loaded.  Run create_sample_data.py first.")

# ---------------------------------------------------------------------------
# Step 2 – Build per-source DataFrames with TOTAL rows and print them
# ---------------------------------------------------------------------------
print("\n" + "=" * 60)
print("STEP 2 – Per-source DataFrames with totals")
print("=" * 60)

display_frames: dict[str, pd.DataFrame] = {}
for name, df in raw_frames.items():
    with_totals = add_totals_row(df)
    display_frames[name] = with_totals
    save_df_csv(with_totals, f"transactions_{name}")
    print(f"\n--- {name} ---")
    print(with_totals.to_string(index=False))

# ---------------------------------------------------------------------------
# Step 3 – Fuzzy grouping of similar Trans_Details
# ---------------------------------------------------------------------------
print("\n" + "=" * 60)
print("STEP 3 – Grouping similar transaction descriptions")
print("=" * 60)

# Combine all raw transactions
all_df = pd.concat(raw_frames.values(), ignore_index=True)

unique_descs = all_df["Trans_Details"].dropna().unique().tolist()

# Build a mapping: original description → canonical (representative) label
canonical: dict[str, str] = {}
groups: list[list[str]] = []

for desc in unique_descs:
    if desc in canonical:
        continue
    # Find all descriptions that match this one at >= threshold
    matches = process.extract(
        desc, unique_descs,
        scorer=fuzz.token_sort_ratio,
        score_cutoff=SIMILARITY_THRESHOLD,
        limit=None,
    )
    group = [m[0] for m in matches]
    # Use the shortest (most generic) name as the canonical label
    label = min(group, key=len)
    for member in group:
        canonical[member] = label
    groups.append(group)

print(f"  {len(unique_descs)} unique descriptions → {len(set(canonical.values()))} groups")

# Map canonical labels back onto the combined frame
all_df["Group"] = all_df["Trans_Details"].map(canonical)

# Aggregate by group
grouped = (
    all_df
    .groupby("Group", as_index=False)
    .agg(
        Trans_Dr=("Trans_Dr", "sum"),
        Trans_Cr=("Trans_Cr", "sum"),
        Transactions=("Trans_Details", "count"),
    )
    .sort_values("Trans_Dr", ascending=False)
    .reset_index(drop=True)
)
grouped["Trans_Dr"] = grouped["Trans_Dr"].round(2)
grouped["Trans_Cr"] = grouped["Trans_Cr"].round(2)

# Append TOTAL row
total_row = pd.DataFrame([{
    "Group": "— TOTAL —",
    "Trans_Dr": round(grouped["Trans_Dr"].sum(), 2),
    "Trans_Cr": round(grouped["Trans_Cr"].sum(), 2),
    "Transactions": grouped["Transactions"].sum(),
}])
grouped_with_total = pd.concat([grouped, total_row], ignore_index=True)

print("\n--- Grouped Summary ---")
print(grouped_with_total.to_string(index=False))
save_df_csv(grouped_with_total, "grouped_summary")

# ---------------------------------------------------------------------------
# Step 4 – Per-source monthly summaries (used for charts)
# ---------------------------------------------------------------------------
all_df["Month"] = all_df["Date"].dt.to_period("M").astype(str)

monthly_dr = (
    all_df.groupby(["Month", "Source"])["Trans_Dr"]
    .sum()
    .unstack(fill_value=0)
)
monthly_cr = (
    all_df.groupby(["Month", "Source"])["Trans_Cr"]
    .sum()
    .unstack(fill_value=0)
)

# ---------------------------------------------------------------------------
# Step 5 – Charts
# ---------------------------------------------------------------------------
print("\n" + "=" * 60)
print("STEP 4 – Generating charts")
print("=" * 60)


def savefig(name: str):
    path = os.path.join(OUTPUT_DIR, f"{name}.png")
    plt.savefig(path, bbox_inches="tight", dpi=150)
    plt.close()
    print(f"  Saved {path}")


# ── Chart 1: Total Debits per source (bar chart) ──────────────────────────
fig, ax = plt.subplots(figsize=(8, 5))
sources = list(raw_frames.keys())
totals_dr = [raw_frames[s]["Trans_Dr"].sum() for s in sources]
bars = ax.bar(sources, totals_dr,
              color=[SOURCE_COLOURS[s] for s in sources], width=0.5)
ax.bar_label(bars, fmt="$%.0f", padding=4, fontsize=9)
ax.set_title("Total Debits by Account", fontsize=14, fontweight="bold")
ax.set_ylabel("Amount ($)")
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:,.0f}"))
ax.set_ylim(0, max(totals_dr) * 1.15)
ax.grid(axis="y", alpha=0.3)
plt.tight_layout()
savefig("chart1_total_debits_by_account")

# ── Chart 2: Total Credits per source (bar chart) ─────────────────────────
fig, ax = plt.subplots(figsize=(8, 5))
totals_cr = [raw_frames[s]["Trans_Cr"].sum() for s in sources]
bars = ax.bar(sources, totals_cr,
              color=[SOURCE_COLOURS[s] for s in sources], width=0.5)
ax.bar_label(bars, fmt="$%.0f", padding=4, fontsize=9)
ax.set_title("Total Credits by Account", fontsize=14, fontweight="bold")
ax.set_ylabel("Amount ($)")
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:,.0f}"))
ax.set_ylim(0, max(totals_cr) * 1.15)
ax.grid(axis="y", alpha=0.3)
plt.tight_layout()
savefig("chart2_total_credits_by_account")

# ── Chart 3: Debits vs Credits per source (grouped bar) ───────────────────
x = range(len(sources))
width = 0.35
fig, ax = plt.subplots(figsize=(9, 5))
bars_dr = ax.bar([i - width / 2 for i in x], totals_dr, width,
                 label="Debits", color="#d62728")
bars_cr = ax.bar([i + width / 2 for i in x], totals_cr, width,
                 label="Credits", color="#2ca02c")
ax.bar_label(bars_dr, fmt="$%.0f", padding=3, fontsize=8)
ax.bar_label(bars_cr, fmt="$%.0f", padding=3, fontsize=8)
ax.set_title("Debits vs Credits by Account", fontsize=14, fontweight="bold")
ax.set_xticks(list(x))
ax.set_xticklabels(sources)
ax.set_ylabel("Amount ($)")
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:,.0f}"))
ax.legend()
ax.grid(axis="y", alpha=0.3)
plt.tight_layout()
savefig("chart3_debits_vs_credits_by_account")

# ── Chart 4: Pie chart – share of total debits by account ─────────────────
fig, ax = plt.subplots(figsize=(7, 7))
ax.pie(
    totals_dr,
    labels=sources,
    autopct="%1.1f%%",
    colors=[SOURCE_COLOURS[s] for s in sources],
    startangle=140,
    wedgeprops={"edgecolor": "white", "linewidth": 1.5},
)
ax.set_title("Share of Total Debits by Account", fontsize=14, fontweight="bold")
plt.tight_layout()
savefig("chart4_debit_share_pie")

# ── Chart 5: Top 10 spending categories (horizontal bar, grouped summary) ─
top10 = grouped.nlargest(10, "Trans_Dr")
fig, ax = plt.subplots(figsize=(10, 6))
bars = ax.barh(top10["Group"], top10["Trans_Dr"], color="#1f77b4")
ax.bar_label(bars, fmt="$%.0f", padding=4, fontsize=9)
ax.set_title("Top 10 Spending Categories (Grouped)", fontsize=14, fontweight="bold")
ax.set_xlabel("Total Debit ($)")
ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:,.0f}"))
ax.invert_yaxis()
ax.grid(axis="x", alpha=0.3)
plt.tight_layout()
savefig("chart5_top10_categories")

# ── Chart 6: Monthly debits stacked by source ─────────────────────────────
if not monthly_dr.empty:
    fig, ax = plt.subplots(figsize=(12, 6))
    monthly_dr.plot(kind="bar", stacked=True, ax=ax,
                    color=[SOURCE_COLOURS.get(c, "#333") for c in monthly_dr.columns])
    ax.set_title("Monthly Debits by Account (Stacked)", fontsize=14, fontweight="bold")
    ax.set_xlabel("Month")
    ax.set_ylabel("Amount ($)")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:,.0f}"))
    ax.legend(title="Account", bbox_to_anchor=(1.01, 1), loc="upper left")
    plt.xticks(rotation=45, ha="right")
    ax.grid(axis="y", alpha=0.3)
    plt.tight_layout()
    savefig("chart6_monthly_debits_stacked")

# ── Chart 7: Monthly credits line chart ───────────────────────────────────
if not monthly_cr.empty:
    fig, ax = plt.subplots(figsize=(12, 5))
    for col in monthly_cr.columns:
        ax.plot(monthly_cr.index, monthly_cr[col],
                marker="o", label=col, color=SOURCE_COLOURS.get(col, "#333"))
    ax.set_title("Monthly Credits by Account", fontsize=14, fontweight="bold")
    ax.set_xlabel("Month")
    ax.set_ylabel("Amount ($)")
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:,.0f}"))
    ax.legend(title="Account")
    plt.xticks(rotation=45, ha="right")
    ax.grid(alpha=0.3)
    plt.tight_layout()
    savefig("chart7_monthly_credits_line")

# ── Chart 8: Pie chart – grouped spending categories ─────────────────────
fig, ax = plt.subplots(figsize=(9, 9))
plot_data = grouped[grouped["Trans_Dr"] > 0].copy()
# Merge small slices (<3%) into "Other"
threshold = plot_data["Trans_Dr"].sum() * 0.03
other_sum = plot_data.loc[plot_data["Trans_Dr"] < threshold, "Trans_Dr"].sum()
plot_data = plot_data[plot_data["Trans_Dr"] >= threshold].copy()
if other_sum > 0:
    plot_data = pd.concat(
        [plot_data,
         pd.DataFrame([{"Group": "Other", "Trans_Dr": other_sum}])],
        ignore_index=True,
    )
ax.pie(
    plot_data["Trans_Dr"],
    labels=plot_data["Group"],
    autopct="%1.1f%%",
    startangle=140,
    wedgeprops={"edgecolor": "white", "linewidth": 1},
    textprops={"fontsize": 8},
)
ax.set_title("Spending by Category (All Accounts)", fontsize=14, fontweight="bold")
plt.tight_layout()
savefig("chart8_spending_categories_pie")

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------
print("\n" + "=" * 60)
print("Analysis complete.  All outputs saved to:", os.path.abspath(OUTPUT_DIR))
print("=" * 60)

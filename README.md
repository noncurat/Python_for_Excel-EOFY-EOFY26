# Python_for_Excel-EOFY-EOFY26
Analysing CBA, PP25-26 and SmartAccess transaction data for EOFY spending.

## Files

| File | Description |
|------|-------------|
| `cba.xlsx` | Commonwealth Bank transaction data (Date, Trans_Details, Trans_Dr, Trans_Cr) |
| `PP25-26.xlsx` | PayPal / payment platform transaction data |
| `smartaccess.xlsx` | SmartAccess savings/transaction account data |
| `create_sample_data.py` | Generates the three sample Excel files above |
| `analysis.py` | Main analysis script – loads data, groups transactions, produces charts |

## Setup

```bash
pip install pandas openpyxl matplotlib rapidfuzz
```

## Usage

### 1. Generate sample data (if Excel files don't exist yet)

```bash
python3 create_sample_data.py
```

### 2. Run the analysis

```bash
python3 analysis.py
```

All outputs are written to the `output/` directory.

## Outputs

### CSV tables (`output/`)

| File | Description |
|------|-------------|
| `transactions_CBA.csv` | CBA transactions with TOTAL row |
| `transactions_PP25-26.csv` | PP25-26 transactions with TOTAL row |
| `transactions_SmartAccess.csv` | SmartAccess transactions with TOTAL row |
| `grouped_summary.csv` | Consolidated table grouping similar Trans_Details across all accounts |

### Charts (`output/`)

| File | Description |
|------|-------------|
| `chart1_total_debits_by_account.png` | Bar chart – total debits per account |
| `chart2_total_credits_by_account.png` | Bar chart – total credits per account |
| `chart3_debits_vs_credits_by_account.png` | Grouped bar – debits vs credits side-by-side per account |
| `chart4_debit_share_pie.png` | Pie chart – each account's share of total debits |
| `chart5_top10_categories.png` | Horizontal bar – top 10 spending categories (fuzzy-grouped) |
| `chart6_monthly_debits_stacked.png` | Stacked bar – monthly debits broken down by account |
| `chart7_monthly_credits_line.png` | Line chart – monthly credits per account |
| `chart8_spending_categories_pie.png` | Pie chart – spending share by category across all accounts |

## How grouping works

Transaction descriptions that are ≥ 75 % similar (using `rapidfuzz`
`token_sort_ratio`) are merged under a single canonical label (the shortest
matching description).  This catches common variations such as:

- `"Woolworths Supermarket"` / `"Woolworths"`
- `"Petrol Station BP"` / `"BP Petrol"`

The similarity threshold can be tuned via the `SIMILARITY_THRESHOLD`
constant at the top of `analysis.py`.

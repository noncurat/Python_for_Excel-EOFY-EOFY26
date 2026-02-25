"""
create_sample_data.py
Generates sample transaction Excel files for cba.xlsx, PP25-26.xlsx and
smartaccess.xlsx with columns: Date, Trans_Details, Trans_Dr, Trans_Cr.
Run this once to create the source data files before running analysis.py.
"""

import pandas as pd
from datetime import date


# ---------------------------------------------------------------------------
# CBA (Commonwealth Bank) transactions
# ---------------------------------------------------------------------------
cba_data = {
    "Date": [
        date(2025, 7, 1), date(2025, 7, 3), date(2025, 7, 5), date(2025, 7, 8),
        date(2025, 7, 10), date(2025, 7, 12), date(2025, 7, 15), date(2025, 7, 18),
        date(2025, 7, 20), date(2025, 7, 22), date(2025, 7, 25), date(2025, 7, 28),
        date(2025, 8, 1), date(2025, 8, 3), date(2025, 8, 7), date(2025, 8, 10),
        date(2025, 8, 12), date(2025, 8, 15), date(2025, 8, 18), date(2025, 8, 22),
        date(2025, 8, 25), date(2025, 8, 28), date(2025, 9, 2), date(2025, 9, 5),
        date(2025, 9, 8), date(2025, 9, 12), date(2025, 9, 15), date(2025, 9, 18),
        date(2025, 9, 22), date(2025, 9, 25),
    ],
    "Trans_Details": [
        "Salary Payment - Employer Co",
        "Woolworths Supermarket",
        "Netflix Subscription",
        "Coles Supermarket",
        "AGL Energy Bill",
        "Bunnings Warehouse",
        "Woolworths Supermarket",
        "Salary Payment - Employer Co",
        "Coles Supermarket",
        "Telstra Phone Bill",
        "Woolworths Supermarket",
        "Petrol Station BP",
        "Salary Payment - Employer Co",
        "Coles Supermarket",
        "Netflix Subscription",
        "Woolworths Supermarket",
        "AGL Energy Bill",
        "Petrol Station BP",
        "Salary Payment - Employer Co",
        "Bunnings Warehouse",
        "Coles Supermarket",
        "Woolworths Supermarket",
        "Salary Payment - Employer Co",
        "Telstra Phone Bill",
        "Coles Supermarket",
        "Woolworths Supermarket",
        "Petrol Station BP",
        "Netflix Subscription",
        "AGL Energy Bill",
        "Bunnings Warehouse",
    ],
    "Trans_Dr": [
        None, 187.50, 15.99, 143.20, 210.00, 89.95, 156.30, None,
        122.80, 95.00, 201.45, 78.60, None, 135.70, 15.99, 178.90,
        215.00, 82.40, None, 95.50, 148.60, 193.25, None, 95.00,
        131.40, 167.80, 86.20, 15.99, 220.00, 102.75,
    ],
    "Trans_Cr": [
        4500.00, None, None, None, None, None, None, 4500.00,
        None, None, None, None, 4500.00, None, None, None,
        None, None, 4500.00, None, None, None, 4500.00, None,
        None, None, None, None, None, None,
    ],
}

df_cba = pd.DataFrame(cba_data)
df_cba.to_excel("cba.xlsx", index=False)
print("Created cba.xlsx")


# ---------------------------------------------------------------------------
# PP25-26 (PayPal / Payment Platform) transactions
# ---------------------------------------------------------------------------
pp_data = {
    "Date": [
        date(2025, 7, 2), date(2025, 7, 6), date(2025, 7, 9), date(2025, 7, 14),
        date(2025, 7, 19), date(2025, 7, 23), date(2025, 7, 29), date(2025, 8, 4),
        date(2025, 8, 8), date(2025, 8, 13), date(2025, 8, 17), date(2025, 8, 21),
        date(2025, 8, 26), date(2025, 9, 1), date(2025, 9, 6), date(2025, 9, 10),
        date(2025, 9, 14), date(2025, 9, 19), date(2025, 9, 23), date(2025, 9, 27),
    ],
    "Trans_Details": [
        "Amazon Online Purchase",
        "Uber Eats Food Delivery",
        "Amazon Online Purchase",
        "Ebay Purchase",
        "Uber Eats Food Delivery",
        "Amazon Online Purchase",
        "Spotify Music Subscription",
        "Ebay Purchase",
        "Uber Eats Food Delivery",
        "Amazon Online Purchase",
        "Spotify Music Subscription",
        "Ebay Purchase",
        "Amazon Online Purchase",
        "Uber Eats Food Delivery",
        "PayPal Payment Received",
        "Amazon Online Purchase",
        "Spotify Music Subscription",
        "Uber Eats Food Delivery",
        "Ebay Purchase",
        "PayPal Payment Received",
    ],
    "Trans_Dr": [
        89.99, 45.50, 134.75, 67.80, 38.90, 210.00, 9.99, 54.20,
        52.30, 178.45, 9.99, 43.60, 95.00, 61.75, None, 145.90,
        9.99, 48.20, 72.40, None,
    ],
    "Trans_Cr": [
        None, None, None, None, None, None, None, None,
        None, None, None, None, None, None, 320.00, None,
        None, None, None, 185.00,
    ],
}

df_pp = pd.DataFrame(pp_data)
df_pp.to_excel("PP25-26.xlsx", index=False)
print("Created PP25-26.xlsx")


# ---------------------------------------------------------------------------
# SmartAccess (savings / transaction account) transactions
# ---------------------------------------------------------------------------
sa_data = {
    "Date": [
        date(2025, 7, 1), date(2025, 7, 4), date(2025, 7, 7), date(2025, 7, 11),
        date(2025, 7, 16), date(2025, 7, 21), date(2025, 7, 26), date(2025, 8, 2),
        date(2025, 8, 6), date(2025, 8, 11), date(2025, 8, 16), date(2025, 8, 20),
        date(2025, 8, 24), date(2025, 8, 29), date(2025, 9, 3), date(2025, 9, 9),
        date(2025, 9, 13), date(2025, 9, 17), date(2025, 9, 21), date(2025, 9, 26),
        date(2025, 10, 1), date(2025, 10, 5), date(2025, 10, 9), date(2025, 10, 14),
        date(2025, 10, 19),
    ],
    "Trans_Details": [
        "Transfer from CBA",
        "Gym Membership Fee",
        "Medical Centre GP Visit",
        "Chemist Warehouse",
        "Transfer from CBA",
        "Restaurant Dining Out",
        "Gym Membership Fee",
        "Transfer from CBA",
        "Medical Centre GP Visit",
        "Restaurant Dining Out",
        "Chemist Warehouse",
        "Transfer from CBA",
        "Gym Membership Fee",
        "Restaurant Dining Out",
        "Transfer from CBA",
        "Medical Centre GP Visit",
        "Restaurant Dining Out",
        "Gym Membership Fee",
        "Chemist Warehouse",
        "Transfer from CBA",
        "Interest Earned",
        "Restaurant Dining Out",
        "Medical Centre GP Visit",
        "Gym Membership Fee",
        "Transfer from CBA",
    ],
    "Trans_Dr": [
        None, 65.00, 85.00, 42.90, None, 120.50, 65.00, None,
        85.00, 95.30, 38.75, None, 65.00, 110.20, None, 85.00,
        88.40, 65.00, 45.60, None, None, 105.75, 85.00, 65.00, None,
    ],
    "Trans_Cr": [
        500.00, None, None, None, 500.00, None, None, 500.00,
        None, None, None, 500.00, None, None, 500.00, None,
        None, None, None, 500.00, 12.50, None, None, None, 500.00,
    ],
}

df_sa = pd.DataFrame(sa_data)
df_sa.to_excel("smartaccess.xlsx", index=False)
print("Created smartaccess.xlsx")

print("\nAll sample data files created successfully.")

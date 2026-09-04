"""
mappings.py

This file acts as a central configuration for cleaning and bucketing Bloomberg data 
so it displays perfectly on the PowerPoint slides. 

You can import these into your presentation builder like this:
from mappings import MUNI_ISSUE_MAP, MUNI_SOURCE_BUCKETS, RATINGS_ORDER

And apply them to your pandas DataFrames like this:
df['MUNI_ISSUE_TYP'] = df['MUNI_ISSUE_TYP'].replace(MUNI_ISSUE_MAP)
"""

# ==========================================
# 1. ISSUE TYPE MAPPINGS
# Cleans up variations of General Obligation, Revenue, etc.
# ==========================================
MUNI_ISSUE_MAP = {
    "General Obligation Ltd": "General Obligation",
    "General Obligation Unltd": "General Obligation",
    "G.O. Debt Certificates": "General Obligation",
    "Revenue Bonds": "Revenue",
    "Revenue Notes": "Revenue",
    "Certificate Participation": "Certificate of Participation",
    "Prerefunded/ETM": "Prerefunded / ETM",
    "Special Assessment": "Special Assessment",
    "Special Tax": "Special Tax",
    "Tax Allocation": "Tax Allocation"
}

# ==========================================
# 2. MUNI SOURCE / REVENUE BUCKETS
# ==========================================
# These are the "Core" buckets you want on your pie charts. 
# You can use this set to check if a row's source should be grouped into "Miscellaneous Revenue".
# Example in Pandas:
# df.loc[~df['MUNI SOURCE'].isin(CORE_INCOME_BUCKETS), 'MUNI SOURCE'] = 'Miscellaneous Revenue'
CORE_INCOME_BUCKETS = {
    "Ad Valorem Property Tax",
    "College & Univ. Rev.",
    "Elec. Pwr. & Lt. Revs.",
    "Highway Revenue Tolls",
    "Hlth, Hosp, Nurshome Rev.",
    "Lease Rev.",
    "Prt, Airprt & Marina Rev.",
    "Sewer Revenue",
    "Water Revenue",
    "Miscellaneous Revenue"
}

# Optional: Clean up Bloomberg's raw MUNI SOURCE strings if they are too long for the slides
MUNI_SOURCE_MAP = {
    "Hlth, Hosp, Nurshome Rev.": "Healthcare & Hospital",
    "Prt, Airprt & Marina Rev.": "Port & Airport",
    "Elec. Pwr. & Lt. Revs.": "Electric Power",
    "Highway Revenue Tolls": "Toll Road Revenue"
}

# ==========================================
# 3. CALLABLE / REFUNDED MAP
# ==========================================
CALLABLE_MAP = {
    "Callable": "Callable",
    "Non-Callable": "Non-Callable",
    "Refunded": "Refunded / Prerefunded"
}

# ==========================================
# 4. RATINGS SORTING ORDERS
# ==========================================
# When creating bar charts or tables, you want the ratings to appear in financial order, 
# not alphabetical order. Use these lists to sort your Pandas Categoricals.
# Example: 
# df['RATINGS'] = pd.Categorical(df['RATINGS'], categories=CONSOLIDATED_RATINGS_ORDER, ordered=True)

CONSOLIDATED_RATINGS_ORDER = [
    "AAA", "AA", "A", "BBB", "BB", "B", "CCC", "CC", "C", "D", "NR", "WR", "SP-1"
]

MOODYS_ORDER = [
    "Aaa", "Aa1", "Aa2", "Aa3", "A1", "A2", "A3", 
    "Baa1", "Baa2", "Baa3", "Ba1", "Ba2", "Ba3", 
    "B1", "B2", "B3", "Caa1", "Caa2", "Caa3", "Ca", "C",
    "MIG1", "NR", "WR", "A2/WR", "WR/WR", "#Aaa"
]

SP_ORDER = [
    "AAA", "AA+", "AA", "AA-", "A+", "A", "A-", 
    "BBB+", "BBB", "BBB-", "BB+", "BB", "BB-", 
    "B+", "B", "B-", "CCC+", "CCC", "CCC-", "CC", "C", "D",
    "A-/A-2", "SP-1", "NR", "WR"
]

# ==========================================
# 5. DURATION & MATURITY BUCKET ORDERS
# ==========================================
TIME_HORIZON_ORDER = [
    "0-1 Years",
    "1-3 Years",
    "3-5 Years",
    "5-7 Years",
    "7-10 Years",
    "10-15 Years",
    "15+ Years"
]

COUPON_RANGE_ORDER = [
    "0%-3%",
    "3%-4%",
    "4%-5%",
    "5%+"
]

# ==========================================
# 6. CREDIT ENHANCEMENTS
# ==========================================
# Maps obscure enhancements to cleaner labels
ENHANCEMENTS_MAP = {
    "St Aid": "State Aid",
    "AG +  St Aid": "Assured Guaranty + State Aid",
    "CA MTG INS": "CA Mortgage Insurance",
    "PSF-GTD": "Permanent School Fund Gtd",
    "N/A": "Uninsured"
}

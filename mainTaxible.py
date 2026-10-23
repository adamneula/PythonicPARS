"""
mainTaxible.py
--------------
Entry point for the taxable Corporate / Treasury / Agency portfolio pipeline.

Mirrors main.py (muni) but:
  - Queries Bloomberg with "Corp" as the yellow-key (or None for auto-detect).
  - Uses TaxableBond instead of Bond, so all Corp-relevant fields are captured.
  - Computes portfolio_weight for every holding after the full portfolio is loaded.
  - Exports a richer Portfolio_Output_Taxable_<timestamp>.xlsx with spread fields,
    seniority, and security classification columns ready for presentationTaxible.py.

Usage:
    python mainTaxible.py
"""

import datetime
import pandas as pd

from bloomberg_client import BloombergClient
from bond_taxable import TaxableBond


# ── Configuration ──────────────────────────────────────────────────────────────

# Path to the input Excel file (two columns: CUSIP, Face Value)
EXCEL_PATH = r"H:\_INSTITUTIONAL DIVISION\INTERN FOLDER\Adam Neulander\PythonicPARS\JacobOct21.xlsx"

# Column indices in the input sheet (0-based)
CUSIP_COL  = 0
FACEVALUE_COL = 1

# Bloomberg yellow-key market sector.
#   "Corp"  — look up all CUSIPs in the Corporate universe (recommended for pure corp portfolios)
#   None    — let Bloomberg auto-detect (safe for mixed Corp + Govt + Agency)
SECURITY_TYPE = "Corp"

# Set to True to run against realistic mock data without a live Bloomberg connection.
# Flip to False on a machine with an active Bloomberg terminal session.
USE_MOCK = False

# Output path (timestamp prevents PermissionError from overwriting an open file)
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
OUTPUT_PATH = rf"H:\_INSTITUTIONAL DIVISION\INTERN FOLDER\Adam Neulander\PythonicPARS\Portfolio_Output_Taxable_{timestamp}.xlsx"


# ── Pipeline functions ─────────────────────────────────────────────────────────

def load_taxable_portfolio(
    file_path: str,
    cusip_col_index: int = 0,
    face_value_col_index: int = 1,
    security_type: str = "Corp",
    use_mock: bool = False,
) -> list[TaxableBond]:
    """
    Reads an Excel file, fetches Bloomberg reference data in a single batch
    request, and returns a list of TaxableBond objects.

    After all bonds are loaded the function computes portfolio_weight for each
    holding (market value / total portfolio market value).

    Args:
        file_path: Path to the input Excel file.
        cusip_col_index: 0-based column index of the CUSIP column.
        face_value_col_index: 0-based column index of the Face Value column.
        security_type: Bloomberg yellow-key suffix ("Corp", "Govt", None, etc.)
    """
    print(f"Reading Excel file: {file_path}")
    df = pd.read_excel(file_path)

    cusips      = df.iloc[:, cusip_col_index].astype(str).tolist()
    face_values = df.iloc[:, face_value_col_index].tolist()

    print(f"Fetching Bloomberg data for {len(cusips)} CUSIPs (yellow-key: '{security_type}')...")

    # Single batch Bloomberg request — no per-CUSIP round trips
    with BloombergClient(use_mock=use_mock, security_type=security_type, is_taxable=True) as bbg:
        all_bbg_data = bbg.fetch_data(cusips)

    # Build portfolio
    portfolio: list[TaxableBond] = []
    for cusip, face_value in zip(cusips, face_values):
        bond_data = all_bbg_data.get(cusip)
        bond = TaxableBond(cusip, face_value, bbg_data=bond_data)
        portfolio.append(bond)

    # Compute portfolio weights now that we have all market values
    total_market_value = sum(b.market_value for b in portfolio)
    if total_market_value > 0:
        for bond in portfolio:
            bond.portfolio_weight = bond.market_value / total_market_value

    return portfolio


def export_taxable_portfolio_to_excel(portfolio: list[TaxableBond], output_path: str):
    """
    Exports the list of TaxableBond objects to a clean Excel spreadsheet.
    Uses vars() so every attribute on the object becomes a column — no manual
    mapping required; adding a new field to TaxableBond automatically shows up here.
    """
    bond_dicts = [vars(b) for b in portfolio]
    df = pd.DataFrame(bond_dicts)
    df.to_excel(output_path, index=False)
    print(f"\nSuccessfully exported all bond data to: {output_path}")


# ── Entry point ────────────────────────────────────────────────────────────────

if __name__ == "__main__":

    # 1. Load portfolio from Bloomberg
    my_portfolio = load_taxable_portfolio(
        EXCEL_PATH,
        cusip_col_index=CUSIP_COL,
        face_value_col_index=FACEVALUE_COL,
        security_type=SECURITY_TYPE,
        use_mock=USE_MOCK,
    )

    # 2. Print summary to console
    print(f"\n--- TAXABLE PORTFOLIO LOADED ({len(my_portfolio)} Holdings) ---")
    total_market_value = sum(b.market_value for b in my_portfolio)

    for b in my_portfolio:
        oas_str = f"OAS: {b.oas_spread:.0f}bps" if b.oas_spread is not None else "OAS: N/A"
        print(
            f"CUSIP: {b.cusip:<10} | {b.security_name:<28} | "
            f"Par: ${b.face_value:>10,.0f} | "
            f"MV: ${b.market_value:>10,.0f} | "
            f"YTW: {b.yield_to_worst or 0:.2f}% | "
            f"{oas_str} | "
            f"Sector: {b.industry_sector:<12} | "
            f"Type: {b.security_type:<15} | "
            f"Rank: {b.payment_rank} | "
            f"Wt: {b.portfolio_weight:.1%}"
        )

    print("-" * 100)
    print(f"TOTAL PORTFOLIO MARKET VALUE: ${total_market_value:,.2f}")

    # 3. Export enriched data to Excel
    export_taxable_portfolio_to_excel(my_portfolio, OUTPUT_PATH)

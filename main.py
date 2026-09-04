import pandas as pd
from bloomberg_client import BloombergClient
from bond import Bond

def load_portfolio_from_excel(file_path: str, cusip_col_index: int = 0, face_value_col_index: int = 1) -> list[Bond]:
    """
    Reads an Excel file, fetches Bloomberg data in batch, and returns a List of Bond objects.
    """
    print(f"Reading Excel file: {file_path}")
    df = pd.read_excel(file_path)
    
    # Extract CUSIPs and Face Values as lists
    cusips = df.iloc[:, cusip_col_index].astype(str).tolist()
    face_values = df.iloc[:, face_value_col_index].tolist()

    print(f"Fetching Bloomberg data for {len(cusips)} CUSIPs...")
    
    # 1. Fetch ALL data from Bloomberg at once
    with BloombergClient(use_mock=False) as bbg:
        all_bbg_data = bbg.fetch_data(cusips)

    # 2. Build the portfolio list
    portfolio = []
    for cusip, face_value in zip(cusips, face_values):
        
        # Grab just the dictionary for this specific CUSIP
        bond_data = all_bbg_data.get(cusip) 
        
        # Create the bond object and add it to the list
        bond = Bond(cusip, face_value, bbg_data=bond_data)
        portfolio.append(bond)

    return portfolio

def export_portfolio_to_excel(portfolio: list[Bond], output_path: str):
    """
    Exports the list of Bond objects back to a clean Excel spreadsheet.
    """
    # Convert the list of Bond objects into a list of dictionaries
    bond_dicts = [vars(b) for b in portfolio]
    
    # Create a DataFrame and export to Excel
    df = pd.DataFrame(bond_dicts)
    df.to_excel(output_path, index=False)
    print(f"\nSuccessfully exported all bond data to: {output_path}")

if __name__ == "__main__":
    import datetime
    
    # The path to your actual Excel sheet
    EXCEL_PATH = r"H:\_INSTITUTIONAL DIVISION\INTERN FOLDER\Adam Neulander\PythonicPARS\Book1.xlsx"
    
    # Dynamically generate the output name with a timestamp to avoid PermissionErrors
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    OUTPUT_PATH = rf"H:\_INSTITUTIONAL DIVISION\INTERN FOLDER\Adam Neulander\PythonicPARS\Portfolio_Output_{timestamp}.xlsx"
    
    # Run the function
    my_portfolio = load_portfolio_from_excel(EXCEL_PATH, cusip_col_index=0, face_value_col_index=1)
    
    # Print out the results to verify it worked!
    print(f"\n--- PORTFOLIO SUCCESSFULLY LOADED ({len(my_portfolio)} Bonds) ---")
    
    total_market_value = 0.0
    for b in my_portfolio:
        print(f"CUSIP: {b.cusip:<10} | {b.security_name:<25} | Par: ${b.face_value:,.2f} | Market Value: ${b.market_value:,.2f} | YTW: {b.yield_to_worst}%")
        total_market_value += b.market_value
        
    print("-" * 75)
    print(f"TOTAL PORTFOLIO MARKET VALUE: ${total_market_value:,.2f}")
    
    # Export all the rich Bloomberg data back to a new Excel sheet!
    export_portfolio_to_excel(my_portfolio, OUTPUT_PATH)


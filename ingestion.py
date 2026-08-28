import pandas as pd
from bloombergCalls import bloomberg
from bondObjects import bondObjects
    
def ingest_cusips(sheet: str, cusip_index: int = 0, face_value_index: int = 1) -> None:
    '''
    Ingests cusips and face values from a spreadsheet to be processed through bloomberg api
    '''
    df = pd.read_excel(sheet)
    cusips = df.iloc[:, cusip_index].tolist()
    face_values = df.iloc[:, face_value_index].tolist()

    with bloomberg() as bbg:
        raw_metadata = bbg.get_bond_metadata(cusips)
    for r in raw_metadata:
        cusip = r.get("cusip")
        face_value = face_values[cusips.index(cusip)]
        # Pass the Bloomberg metadata dictionary 'r' to populate the bond attributes
        bond = bondObjects(cusip, face_value, bbg_data=r)
        print(f"Bond CUSIP: {bond.cusip}, Face Value: {bond.face_value}, Current Price: {bond.current_price}, Maturity Date: {bond.maturity_date}, Next Call Date: {bond.next_call_date}, Yield to Worst: {bond.yield_to_worst}, Yield to Maturity: {bond.yield_to_maturity}, Duration: {bond.duration}, Convexity: {bond.convexity}, S&P Rating: {bond.sp_rating}, Moody's Rating: {bond.moodys_rating}, Industry Sector: {bond.industry_sector}, State Code: {bond.state_code}, Municipal Tax Provision: {bond.municipal_tax_provision}")


cusip_index = input("Enter the column number for the cusip column (zero-indexed): ")
face_value_index = input("Enter the column number for the face value column (zero-indexed): ")
ingest_cusips(r"H:\_INSTITUTIONAL DIVISION\INTERN FOLDER\Adam Neulander\PythonicPARS\Book1.xlsx", int(cusip_index), int(face_value_index))

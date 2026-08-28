import pandas as pd
import blpapi

client_holdings = {}

def ingest_cusips(sheet: str, cusip_index: int = 0, face_value_index: int = 1) -> None:
    '''
    Ingests cusips and face values from a spreadsheet to be processed through bloomberg api
    '''
    df = pd.read_excel(sheet)
    
    
    cusips = df.iloc[:, cusip_index].tolist()
    face_values = df.iloc[:, face_value_index].tolist()

    for c, fv in zip(cusips, face_values):
        print(f"CUSIP: {c}, Face Value: {fv}")
        client_holdings[c] = fv


cusip_index = input("Enter the column number for the cusip column (zero-indexed): ")
face_value_index = input("Enter the column number for the face value column (zero-indexed): ")
ingest_cusips("H:\_INSTITUTIONAL DIVISION\INTERN FOLDER\Adam Neulander\PythonicPARS\Book1.xlsx", int(cusip_index), int(face_value_index))

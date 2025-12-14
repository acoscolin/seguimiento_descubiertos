
import pandas as pd
import json
import traceback

FILE_PATH = "SEGUIMIENTO DESCUBIERTOS TRAB-RUTA.xlsx"

def inspect_excel():
    try:
        # Load the Excel file
        print(f"Loading {FILE_PATH}...")
        # attempt to load with default engine (openpyxl usually)
        df = pd.read_excel(FILE_PATH)
        
        print("\n--- Basic Info ---")
        df.info()
        
        print("\n--- Column Headers ---")
        print(list(df.columns))
        
        print("\n--- First 5 Rows ---")
        print(df.head().to_string())
        
        print("\n--- Sample JSON Record ---")
        # Convert first row to JSON to see structure
        if not df.empty:
            print(df.iloc[0].to_json(date_format='iso'))
            
        print("\n--- Missing Values ---")
        print(df.isnull().sum())
        
    except Exception as e:
        print(f"Error inspecting Excel: {e}")
        traceback.print_exc()

if __name__ == "__main__":
    inspect_excel()

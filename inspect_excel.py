import pandas as pd
import json

def inspect_file(filepath):
    print(f"--- Inspecting {filepath} ---")
    try:
        xls = pd.ExcelFile(filepath)
        print(f"Sheet names: {xls.sheet_names}")
        for sheet in xls.sheet_names:
            print(f"\nSheet: {sheet}")
            df = pd.read_excel(filepath, sheet_name=sheet, nrows=5)
            print("Columns:", list(df.columns))
            print("First row values:")
            if not df.empty:
                print(df.iloc[0].to_dict())
            else:
                print("Empty sheet")
    except Exception as e:
        print(f"Error reading {filepath}: {e}")

if __name__ == "__main__":
    inspect_file("RECA.04.xlsx")
    inspect_file("RECA.05.xlsx")
    inspect_file("TABLA ACUMULADOS.xlsx")

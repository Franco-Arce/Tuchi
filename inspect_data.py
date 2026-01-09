
import pandas as pd
import os

files = ["Extracto banco 01 2026.xlsx", "Libro Banco 01 2026.xlsx"]
cwd = os.getcwd()

for file in files:
    path = os.path.join(cwd, file)
    print(f"--- Inspecting: {file} ---")
    try:
        # Load first few rows to inspect headers and data types
        df = pd.read_excel(path, nrows=5)
        print("Columns:", df.columns.tolist())
        print("\nHead:")
        print(df.head())
        print("\nTypes:")
        print(df.dtypes)
    except Exception as e:
        print(f"Error reading {file}: {e}")
    print("\n")


import pandas as pd
import os

file = "Libro Banco 01 2026.xlsx"
print(f"--- Inspecting Headers of: {file} ---")
try:
    df = pd.read_excel(file, header=None, nrows=3)
    # Print row 1 (index 1) which seems to be the header
    print("Row 1 content:")
    print(df.iloc[1].tolist())
    
    # Also print row 2 to see data examples corresponding to headers
    print("Row 2 content:")
    print(df.iloc[2].tolist())

except Exception as e:
    print(f"Error reading {file}: {e}")

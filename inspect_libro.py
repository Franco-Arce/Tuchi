
import pandas as pd
import os

file = "Libro Banco 01 2026.xlsx"
print(f"--- Inspecting: {file} with header=1 ---")
try:
    # Try creating headers from row 1 (index 1), assuming row 0 is garbage?
    # Or maybe row 0 is the header.
    # Let's read with header=None and print top 5 rows to see where the headers are.
    df = pd.read_excel(file, header=None, nrows=10)
    print(df)
except Exception as e:
    print(f"Error reading {file}: {e}")

import pandas as pd
from logic import process_reconciliation

# Test reconciliation with actual data
print("Testing Bank Reconciliation...")
print("=" * 60)

output, summary = process_reconciliation('Libro Banco 01 2026.xlsx', 'Extracto banco 01 2026.xlsx')

# Save output to file for inspection
with open('test_output.xlsx', 'wb') as f:
    f.write(output.getvalue())

print("\n[OK] SUMMARY:")
print(f"  Saldo Final Banco: ${summary['saldo_final_banco']:,.2f}")
print(f"  Saldo Final Libro: ${summary['saldo_final_libro']:,.2f}")
print(f"  Diferencia Total: ${summary['diferencia_total']:,.2f}")
print(f"\n  Items Coincidentes: {summary['items_coinciden']}")
print(f"  Diferencias Temporales: {summary['diferencias_temporales_count']} (${summary['diferencias_temporales_monto']:,.2f})")
print(f"  Diferencias Permanentes: {summary['diferencias_permanentes_count']} (${summary['diferencias_permanentes_monto']:,.2f})")

print("\n[*] Excel Output Sheets:")
xls = pd.ExcelFile('test_output.xlsx')
for sheet in xls.sheet_names:
    df = pd.read_excel('test_output.xlsx', sheet_name=sheet)
    print(f"\n  {sheet}: {len(df)} rows, {len(df.columns)} columns")
    if len(df) > 0 and len(df) <= 10:
        print(df.head())

print("\n[OK] Test completed! Check 'test_output.xlsx' for detailed output.")

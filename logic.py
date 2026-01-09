
import pandas as pd
import re
import io

def clean_amount(val):
    """
    Cleans amount strings like '1.050.000,00' to float 1050000.00.
    Handles floats/ints gracefully.
    """
    if pd.isna(val):
        return 0.0
    if isinstance(val, (float, int)):
        return float(val)
    if isinstance(val, str):
        # Remove dots (thousands), replace comma with dot (decimal)
        val = val.replace('.', '').replace(',', '.')
        try:
            return float(val)
        except ValueError:
            return 0.0
    return 0.0

def extract_check_numbers(text):
    """
    Extracts all sequences of digits inside parentheses from the text.
    Example: 'Cheques de terceros (36142161)...' -> ['36142161']
    """
    if not isinstance(text, str):
        return []
    # Regex to find digits inside ()
    matches = re.findall(r'\((\d+)\)', text)
    return matches

def load_data(libro_file, extracto_file):
    """
    Loads and pre-processes the two excel files.
    """
    # --- Load Libro ---
    # Using header=1 as inspected
    df_libro = pd.read_excel(libro_file, header=1)
    
    # Standardize columns
    # We expect 'Concepto' and 'Ingreso' (or 'Egreso' depending on operation type, but user said 'Ingreso')
    # Let's handle both just in case, or focus on Ingreso as per "Cheques de terceros" usually being deposits.
    
    # Basic cleaning
    df_libro['Monto_Libro'] = df_libro['Ingreso'].apply(clean_amount)
    # If Ingreso is 0/NaN, maybe check Egreso? User example showed 'Ingreso' for the big amounts.
    # Keep it simple for now based on inspection.
    
    df_libro['Fecha Pago '] = pd.to_datetime(df_libro['Fecha Pago '], errors='coerce')
    
    # Filter only rows with amounts? Or keep all?
    # Better to keep all for completeness, but reconciliation focuses on those with checks.
    
    # --- Load Extracto ---
    df_extracto = pd.read_excel(extracto_file)
    # Columns expected: 'Numero de Comprobante', 'Crditos' or 'Créditos' (encoding depending)
    # In inspection it was 'Crditos'. Let's normalize column names if needed or use index.
    # The inspection showed 'Crditos'.
    
    # Rename columns to avoid encoding issues
    # Find column that looks like Credit
    col_map = {}
    for col in df_extracto.columns:
        col_lower = str(col).lower()
        if 'crditos' in col_lower or 'créditos' in col_lower or 'creditos' in col_lower or 'crditos' in col_lower:
            col_map[col] = 'Creditos'
        elif 'numero de comprobante' in col_lower or 'número de comprobante' in col_lower:
            col_map[col] = 'Numero de Comprobante'
        elif 'fecha' in col_lower and 'fecha' not in col_map.values(): 
             col_map[col] = 'Fecha'
        elif 'descripci' in col_lower:
             col_map[col] = 'Descripcion'
             
    if 'Creditos' not in col_map.values():
        pass

    df_extracto.rename(columns=col_map, inplace=True)

    if 'Creditos' not in df_extracto.columns:
         raise KeyError(f"No se encontró la columna 'Creditos'. Columnas detectadas: {df_extracto.columns.tolist()}")

    df_extracto['Monto_Banco'] = df_extracto['Creditos'].apply(clean_amount)
    df_extracto['Fecha'] = pd.to_datetime(df_extracto['Fecha'], errors='coerce')
    
    return df_libro, df_extracto

def process_reconciliation(libro_file, extracto_file):
    """
    Main processing function.
    Returns:
        output_file (BytesIO): The Excel file content.
        summary (dict): Stats and explanation for the UI.
    """
    df_libro, df_extracto = load_data(libro_file, extracto_file)
    
    # --- 1. Extraction & Explosion ---
    # Add a unique ID to original Libro rows to group back later
    df_libro['Libro_ID'] = df_libro.index
    
    # Extract check numbers
    df_libro['Cheques_Extraidos'] = df_libro['Concepto'].apply(extract_check_numbers)
    
    # Rows with no checks extracted
    df_libro_sin_cheques = df_libro[df_libro['Cheques_Extraidos'].apply(len) == 0].copy()
    
    # Explode
    df_exploded = df_libro.explode('Cheques_Extraidos')
    
    # Convert check numbers to string for merging (ensure matching types)
    df_exploded['check_match_id'] = df_exploded['Cheques_Extraidos'].astype(str).str.strip()
    df_extracto['check_match_id'] = df_extracto['Numero de Comprobante'].astype(str).str.strip()
    
    # --- 2. Matching ---
    # Merge Libro-Exploded with Extracto
    # Note: 'Fecha' in Extracto will NOT have suffix _B unless Libro also has 'Fecha'.
    # Libro has 'Fecha Pago '. So Extracto 'Fecha' stays 'Fecha'.
    
    merged = pd.merge(
        df_exploded, 
        df_extracto, 
        on='check_match_id', 
        how='left', 
        indicator=True,
        suffixes=('_L', '_B')
    )
    
    # --- 3. Group Back & Validate ---
    # Group by original Libro Row to verify totals
    # We sum 'Monto_Banco' found for each check in the exploded line
    
    # Correct column name for Bank Date is likely 'Fecha' if no collision
    date_col = 'Fecha_B' if 'Fecha_B' in merged.columns else 'Fecha'
    
    agg_funcs = {
        'Monto_Banco': 'sum',
        'check_match_id': lambda x: list(x.dropna()), # List of matched checks
        date_col: lambda x: sorted(list(x.dropna().unique())), # Dates found in bank
        '_merge': lambda x: list(x)
    }
    
    # We need to keep original columns from Libro
    # Grouping key
    group_cols = ['Libro_ID', 'Fecha Pago ', 'Concepto', 'Monto_Libro']
    
    # It's easier to aggregate the numeric data first, then join back to original libro
    grouped_matches = merged.groupby('Libro_ID')['Monto_Banco'].sum().reset_index()
    grouped_checks = merged.groupby('Libro_ID')['check_match_id'].apply(list).reset_index(name='Cheques_Encontrados_Detalle')
    
    # Join back to original Libro
    final_df = pd.merge(df_libro, grouped_matches, on='Libro_ID', how='left')
    final_df = pd.merge(final_df, grouped_checks, on='Libro_ID', how='left')
    
    # For rows that didn't have any matches or checks, Monto_Banco might be 0 or NaN.
    final_df['Monto_Banco'] = final_df['Monto_Banco'].fillna(0)
    
    # Calculate Difference
    final_df['Diferencia'] = final_df['Monto_Libro'] - final_df['Monto_Banco']
    
    # If 'Cheques_Extraidos' was empty, we shouldn't perhaps expect a match in this logic?
    # The user use case is primarily for the bundled checks.
    # Let's mark status.
    
    def get_status(row):
        matches = len(row['Cheques_Extraidos'])
        if matches == 0:
            return "Sin Cheques Identificados"
        if abs(row['Diferencia']) < 1.0: # Tolerance for float math
            return "Conciliado OK"
        if row['Monto_Banco'] == 0:
            return "No Encontrado en Banco"
        return "Diferencia de Monto"

    final_df['Estado'] = final_df.apply(get_status, axis=1)
    
    # --- Generate Summary ---
    total_monto_libro = final_df[final_df['Cheques_Extraidos'].apply(len) > 0]['Monto_Libro'].sum()
    total_conciliado = final_df[final_df['Estado'] == 'Conciliado OK']['Monto_Banco'].sum()
    
    summary = {
        "total_registros_procesados": len(final_df),
        "registros_con_cheques": len(final_df[final_df['Cheques_Extraidos'].apply(len) > 0]),
        "registros_conciliados_ok": len(final_df[final_df['Estado'] == 'Conciliado OK']),
        "monto_total_libro_analizado": total_monto_libro,
        "monto_total_conciliado": total_conciliado,
        "diferencia_global": total_monto_libro - total_conciliado
    }
    
    # --- Export to Excel ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, sheet_name='Conciliacion', index=False)
        
        # Add a sheet for the raw exploded view for debug? 
        # User asked for "un solo archivo xlsx explicando lo que se hizo".
        # Maybe a "Detalle Cruce" sheet is useful.
        
        # Ensure 'Descripcion' exists or use fallback
        desc_col = 'Descripcion' if 'Descripcion' in merged.columns else merged.columns[1] # fallback
        cols_to_export = ['Libro_ID', 'check_match_id', 'Monto_Banco', date_col, desc_col, 'Numero de Comprobante']
        # Filter only existing columns just in case
        cols_to_export = [c for c in cols_to_export if c in merged.columns]
        
        merged_subset = merged[cols_to_export]
        merged_subset.to_excel(writer, sheet_name='Detalle_Cheques', index=False)
        
        # Add Summary Sheet
        summary_df = pd.DataFrame([summary]).T.reset_index()
        summary_df.columns = ['Metrica', 'Valor']
        summary_df.to_excel(writer, sheet_name='Resumen_Ejecutivo', index=False)
        
    output.seek(0)
    
    return output, summary

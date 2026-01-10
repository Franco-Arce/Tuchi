
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

def categorize_difference(row, source='extracto'):
    """
    Categorizes a difference as temporary or permanent based on keywords and context.
    
    Args:
        row: DataFrame row with transaction data
        source: 'extracto' or 'libro' - where the unmatched item came from
    
    Returns:
        dict with keys: category, subcategory, requires_accounting_entry, is_temporary
    """
    descripcion = str(row.get('Descripcion', row.get('Concepto', ''))).lower()
    
    # Permanent differences keywords (require accounting entries)
    permanent_keywords = {
        'comision': 'Comisiones',
        'impuesto': 'Impuestos y percepciones',
        'imp.': 'Impuestos y percepciones',
        'percep': 'Impuestos y percepciones',
        'debito automatico': 'Débito automático',
        'debito autom': 'Débito automático',
        'sueldo': 'Sueldos y cargas sociales',
        'carga social': 'Sueldos y cargas sociales',
        'rechazo': 'Cheques rechazados',
        'devuelto': 'Cheques rechazados',
        'anulacion': 'Anulaciones',
        'ley 25413': 'Impuestos y percepciones',
        'ing. bruto': 'Impuestos y percepciones',
        'ingresos brutos': 'Impuestos y percepciones'
    }
    
    # Temporary differences keywords (adjust without accounting entries)
    temporary_keywords = {
        'acreditacion': 'Acreditaciones en tránsito',
        'transferencia': 'Transferencias transitorias',
        'transito': 'Depósitos en tránsito',
        'tránsito': 'Depósitos en tránsito',
        'en tránsito': 'Depósitos en tránsito',
        'en transito': 'Depósitos en tránsito',
        'prisma': 'Acreditaciones tarjetas',
        'tarjeta': 'Acreditaciones tarjetas'
    }
    
    # Check for permanent differences first
    for keyword, subcategory in permanent_keywords.items():
        if keyword in descripcion:
            return {
                'category': 'Permanente',
                'subcategory': subcategory,
                'requires_accounting_entry': True,
                'is_temporary': False
            }
    
    # Check for temporary differences
    for keyword, subcategory in temporary_keywords.items():
        if keyword in descripcion:
            return {
                'category': 'Temporal',
                'subcategory': subcategory,
                'requires_accounting_entry': False,
                'is_temporary': True
            }
    
    # Default categorization based on source
    if source == 'libro':
        # Item in books but not in bank = likely deposit in transit
        return {
            'category': 'Temporal',
            'subcategory': 'Depósitos en tránsito',
            'requires_accounting_entry': False,
            'is_temporary': True
        }
    else:
        # Item in bank but not in books = likely omitted entry
        return {
            'category': 'Permanente',
            'subcategory': 'Notas de débito/crédito omitidas',
            'requires_accounting_entry': True,
            'is_temporary': False
        }

def load_data(libro_file, extracto_file):
    """
    Loads and pre-processes the two excel files.
    """
    # --- Load Libro ---
    df_libro = pd.read_excel(libro_file, header=1)
    
    # Clean amounts - handle both Ingreso and Egreso
    df_libro['Ingreso_Clean'] = df_libro['Ingreso'].apply(clean_amount)
    df_libro['Egreso_Clean'] = df_libro['Egreso'].apply(clean_amount)
    
    # Net amount: Ingreso is positive, Egreso is negative
    df_libro['Monto_Libro'] = df_libro['Ingreso_Clean'] - df_libro['Egreso_Clean']
    
    df_libro['Fecha Pago '] = pd.to_datetime(df_libro['Fecha Pago '], errors='coerce')
    
    # --- Load Extracto ---
    df_extracto = pd.read_excel(extracto_file)
    
    # Rename columns to avoid encoding issues
    col_map = {}
    for col in df_extracto.columns:
        col_lower = str(col).lower()
        if 'créditos' in col_lower or 'creditos' in col_lower or 'crditos' in col_lower:
            col_map[col] = 'Creditos'
        elif 'débitos' in col_lower or 'debitos' in col_lower or 'dbitos' in col_lower:
            col_map[col] = 'Debitos'
        elif 'numero de comprobante' in col_lower or 'número de comprobante' in col_lower:
            col_map[col] = 'Numero de Comprobante'
        elif 'fecha' in col_lower and 'Fecha' not in col_map.values(): 
            col_map[col] = 'Fecha'
        elif 'descripci' in col_lower:
            col_map[col] = 'Descripcion'
        elif 'saldo' in col_lower:
            col_map[col] = 'Saldo'
              
    df_extracto.rename(columns=col_map, inplace=True)
    
    # Clean amounts
    if 'Creditos' in df_extracto.columns:
        df_extracto['Creditos_Clean'] = df_extracto['Creditos'].apply(clean_amount)
    else:
        df_extracto['Creditos_Clean'] = 0.0
        
    if 'Debitos' in df_extracto.columns:
        df_extracto['Debitos_Clean'] = df_extracto['Debitos'].apply(clean_amount)
    else:
        df_extracto['Debitos_Clean'] = 0.0
    
    # Net amount: Credits are positive, Debits are negative
    df_extracto['Monto_Banco'] = df_extracto['Creditos_Clean'] - df_extracto['Debitos_Clean']
    
    df_extracto['Fecha'] = pd.to_datetime(df_extracto['Fecha'], errors='coerce')
    
    return df_libro, df_extracto

def process_reconciliation(libro_file, extracto_file):
    """
    Main processing function for bank reconciliation.
    Returns:
        output_file (BytesIO): The Excel file content.
        summary (dict): Stats and explanation for the UI.
    """
    df_libro, df_extracto = load_data(libro_file, extracto_file)
    
    # --- Add unique IDs ---
    df_libro['Libro_ID'] = df_libro.index
    df_extracto['Banco_ID'] = df_extracto.index
    
    # --- Extract check numbers ---
    df_libro['Cheques_Extraidos'] = df_libro['Concepto'].apply(extract_check_numbers)
    
    # --- Explode libro for matching ---
    df_exploded = df_libro.explode('Cheques_Extraidos')
    df_exploded['check_match_id'] = df_exploded['Cheques_Extraidos'].astype(str).str.strip()
    df_extracto['check_match_id'] = df_extracto['Numero de Comprobante'].astype(str).str.strip()
    
    # --- Matching ---
    # First, match by check numbers
    merged = pd.merge(
        df_exploded,
        df_extracto,
        on='check_match_id',
        how='outer',
        indicator=True,
        suffixes=('_L', '_B')
    )
    
    # Mark matched items
    merged['Matched'] = merged['_merge'] == 'both'
    
    # --- Categorize unmatched items from Extracto ---
    extracto_unmatched = df_extracto[~df_extracto['Banco_ID'].isin(
        merged[merged['Matched']]['Banco_ID'].dropna()
    )].copy()
    
    # Apply categorization
    extracto_unmatched['Categoria'] = extracto_unmatched.apply(
        lambda row: categorize_difference(row, source='extracto'),
        axis=1
    )
    
    # --- Categorize unmatched items from Libro ---
    libro_matched_ids = merged[merged['Matched']]['Libro_ID'].dropna().unique()
    libro_unmatched = df_libro[~df_libro['Libro_ID'].isin(libro_matched_ids)].copy()
    
    libro_unmatched['Categoria'] = libro_unmatched.apply(
        lambda row: categorize_difference(row, source='libro'),
        axis=1
    )
    
    # --- Get bank ending balance (last saldo in extracto) ---
    if 'Saldo' in df_extracto.columns:
        # Clean saldo
        df_extracto['Saldo_Clean'] = df_extracto['Saldo'].apply(clean_amount)
        saldo_final_banco = df_extracto['Saldo_Clean'].iloc[-1]
    else:
        # Calculate from movements
        saldo_final_banco = df_extracto['Monto_Banco'].sum()
    
    # --- Get libro ending balance ---
    saldo_final_libro = df_libro['Monto_Libro'].sum()
    
    # --- Build reconciliation statement following schema format ---
    # Schema structure:
    # Col 0: Empty
    # Col 1: Signo (+/-)
    # Col 2: Concepto principal
    # Col 3-4: Empty
    # Col 5: CONCEPTO (detalle/subcategoría)
    # Col 6: Se ajusta sin asiento contable
    # Col 7: Empty  
    # Col 8: Se ajusta con asiento contable
    # Col 9: Monto
    # Col 10: Saldo Acumulado
    
    reconciliation_rows = []
    running_balance = saldo_final_banco
    
    # Row 0: Title
    reconciliation_rows.append([
        '', '', f'Conciliación Bancaria - {df_libro["Fecha Pago "].dt.strftime("%B %Y").iloc[0] if len(df_libro) > 0 else ""}',
        '', '', '', '', '', '', '', '', '', ''
    ])
    
    # Row 1: Column headers
    reconciliation_rows.append([
        '', '', '', '', '', 'CONCEPTO', 'Se ajusta sin asiento contable', 'CONCEPTO', 
        'Se ajusta con asiento contable', '', running_balance, '', ''
    ])
    
    # Row 2: Starting balance
    reconciliation_rows.append([
        '', '', 'Saldo final Banco', '', '', '', '', '', '', '', running_balance, '', ''
    ])
    
    # --- TEMPORAL DIFFERENCES (Sin asiento contable) ---
    
    # 1. Deposits in transit
    depositos_transito = libro_unmatched[
        libro_unmatched['Categoria'].apply(lambda x: x.get('is_temporary', False) and x.get('subcategory') == 'Depósitos en tránsito')
    ]
    
    if len(depositos_transito) > 0:
        reconciliation_rows.append([
            '', '+', 'Depósitos en tránsito', '', '', '', '', '', '', '', '', '', ''
        ])
        for _, row in depositos_transito.iterrows():
            monto = row['Monto_Libro']
            running_balance += monto
            concepto_corto = row['Concepto'][:80] if len(row['Concepto']) > 80 else row['Concepto']
            reconciliation_rows.append([
                '', '', '', '', '', concepto_corto, monto, '', '', monto, running_balance, '', ''
            ])
    
    # 2. Other temporary differences from libro
    otras_temp_libro = libro_unmatched[
        libro_unmatched['Categoria'].apply(lambda x: x.get('is_temporary', False) and x.get('subcategory') != 'Depósitos en tránsito')
    ]
    
    if len(otras_temp_libro) > 0:
        # Group by subcategory
        for subcategory, group in otras_temp_libro.groupby(
            otras_temp_libro['Categoria'].apply(lambda x: x.get('subcategory', 'Otros'))
        ):
            reconciliation_rows.append([
                '', '+', subcategory, '', '', '', '', '', '', '', '', '', ''
            ])
            for _, row in group.iterrows():
                monto = row['Monto_Libro']
                running_balance += monto
                concepto_corto = row['Concepto'][:80] if len(row['Concepto']) > 80 else row['Concepto']
                reconciliation_rows.append([
                    '', '', '', '', '', concepto_corto, monto, '', '', monto, running_balance, '', ''
                ])
    
    # 3. Temporary differences from extracto (credits in transit)
    temp_extracto = extracto_unmatched[
        extracto_unmatched['Categoria'].apply(lambda x: x.get('is_temporary', False))
    ]
    
    if len(temp_extracto) > 0:
        # Group by subcategory
        for subcategory, group in temp_extracto.groupby(
            temp_extracto['Categoria'].apply(lambda x: x.get('subcategory', 'Otros'))
        ):
            reconciliation_rows.append([
                '', '+', subcategory, '', '', '', '', '', '', '', '', '', ''
            ])
            for _, row in group.iterrows():
                monto = row['Monto_Banco']
                running_balance += monto
                desc_corto = row.get('Descripcion', '')[:80]
                reconciliation_rows.append([
                    '', '', '', '', '', desc_corto, monto if monto > 0 else '', '', monto if monto < 0 else '', monto, running_balance, '', ''
                ])
    
    # --- PERMANENT DIFFERENCES (Con asiento contable) ---
    
    # 4. Credits not recorded (from extracto) - Notas de crédito omitidas
    creditos_no_registrados = extracto_unmatched[
        (extracto_unmatched['Categoria'].apply(lambda x: not x.get('is_temporary', True))) &
        (extracto_unmatched['Monto_Banco'] > 0)
    ]
    
    if len(creditos_no_registrados) > 0:
        reconciliation_rows.append([
            '', '+', 'Notas de crédito omitidas por la empresa', '', '', '', '', '', '', '', '', '', ''
        ])
        
        # Group by subcategory
        for subcategory, group in creditos_no_registrados.groupby(
            creditos_no_registrados['Categoria'].apply(lambda x: x.get('subcategory', 'Otros'))
        ):
            for _, row in group.iterrows():
                monto = row['Monto_Banco']
                running_balance += monto
                desc_corto = row.get('Descripcion', '')[:80]
                reconciliation_rows.append([
                    '', '', '', '', '', desc_corto, '', '', monto, monto, running_balance, '', ''
                ])
    
    # 5. Debits not recorded (from extracto) - Notas de débito omitidas
    debitos_no_registrados = extracto_unmatched[
        (extracto_unmatched['Categoria'].apply(lambda x: not x.get('is_temporary', True))) &
        (extracto_unmatched['Monto_Banco'] < 0)
    ]
    
    if len(debitos_no_registrados) > 0:
        reconciliation_rows.append([
            '', '+', 'Notas de débito omitidas por la empresa', '', '', '', '', '', '', '', '', '', ''
        ])
        
        # Group by subcategory and aggregate
        for subcategory, group in debitos_no_registrados.groupby(
            debitos_no_registrados['Categoria'].apply(lambda x: x.get('subcategory', 'Otros'))
        ):
            total_grupo = group['Monto_Banco'].sum()
            running_balance += total_grupo
            
            reconciliation_rows.append([
                '', '', '', '', '', subcategory, '', '', total_grupo, total_grupo, running_balance, '', ''
            ])
    
    # 6. Cheques pendientes / registrados de más
    reconciliation_rows.append([
        '', '+', 'Cheques omitidos o registrados de menos', '', '', '', '', '', '', 0.00, running_balance, '', ''
    ])
    
    reconciliation_rows.append([
        '', '+', 'Depósitos registrados de más', '', '', '', '', '', '', 0.00, running_balance, '', ''
    ])
    
    reconciliation_rows.append([
        '', '-', 'Cheques pendientes', '', '', '', '', '', '', '', '', '', ''
    ])
    
    # Find pending cheques (in extracto but not matched)
    cheques_pendientes = extracto_unmatched[
        (extracto_unmatched['check_match_id'].notna()) & 
        (extracto_unmatched['check_match_id'] != '') &
        (extracto_unmatched['check_match_id'] != 'nan') &
        (extracto_unmatched['Monto_Banco'] < 0)
    ]
    
    if len(cheques_pendientes) > 0:
        for _, row in cheques_pendientes.iterrows():
            monto = abs(row['Monto_Banco'])
            running_balance -= monto
            desc = f"Cheque {row['check_match_id']}"
            reconciliation_rows.append([
                '', '', '', '', '', desc, monto, '', '', monto, running_balance, '', ''
            ])
    
    reconciliation_rows.append([
        '', '-', 'Cheques registrados de más', '', '', '', '', '', '', 0.00, running_balance, '', ''
    ])
    
    # --- FINAL BALANCE ---
    reconciliation_rows.append([
        '', '', 'Saldo final Libro', '', '', '', '', '', '', '', running_balance, '', ''
    ])
    
    # --- Calculate summary ---
    matched_count = merged['Matched'].sum()
    temporal_diff = extracto_unmatched[extracto_unmatched['Categoria'].apply(lambda x: x.get('is_temporary', False))]['Monto_Banco'].sum() + \
                    libro_unmatched[libro_unmatched['Categoria'].apply(lambda x: x.get('is_temporary', False))]['Monto_Libro'].sum()
    permanente_diff = extracto_unmatched[extracto_unmatched['Categoria'].apply(lambda x: not x.get('is_temporary', True))]['Monto_Banco'].sum()
    
    summary = {
        "saldo_final_banco": saldo_final_banco,
        "saldo_final_libro": saldo_final_libro,
        "diferencia_total": saldo_final_libro - saldo_final_banco,
        "items_coinciden": int(matched_count),
        "diferencias_temporales_count": len(libro_unmatched[libro_unmatched['Categoria'].apply(lambda x: x.get('is_temporary', False))]) + 
                                       len(extracto_unmatched[extracto_unmatched['Categoria'].apply(lambda x: x.get('is_temporary', False))]),
        "diferencias_permanentes_count": len(extracto_unmatched[extracto_unmatched['Categoria'].apply(lambda x: not x.get('is_temporary', True))]),
        "diferencias_temporales_monto": temporal_diff,
        "diferencias_permanentes_monto": permanente_diff
    }
    
    
    # --- Export to Excel ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Sheet 1: Reconciliation Statement (matching schema format)
        # Create DataFrame from rows with proper column names
        df_reconciliation = pd.DataFrame(reconciliation_rows, columns=[
            '', 'Signo', 'Concepto Principal', '', '', 'CONCEPTO', 
            'Se ajusta sin asiento contable', 'CONCEPTO', 'Se ajusta con asiento contable',
            'Monto', 'Saldo Acumulado', '', ''
        ])
        
        df_reconciliation.to_excel(writer, sheet_name='Conciliacion Bancaria', index=False, header=False)
        
        # Get workbook and worksheet to apply formatting
        workbook = writer.book
        worksheet = writer.sheets['Conciliacion Bancaria']
        
        # === DEFINE FORMATS ===
        
        # Title format (row 0)
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        # Header format (row 1)
        header_format = workbook.add_format({
            'bold': True,
            'font_size': 9,
            'bg_color': '#4472C4',
            'font_color': 'white',
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'text_wrap': True
        })
        
        # Saldo inicial/final (yellow highlight)
        saldo_format = workbook.add_format({
            'bold': True,
            'font_size': 10,
            'bg_color': '#FFFF00',
            'border': 1,
            'num_format': '$ #,##0.00',
            'align': 'right'
        })
        
        # Saldo text format
        saldo_text_format = workbook.add_format({
            'bold': True,
            'font_size': 10,
            'bg_color': '#FFFF00',
            'border': 1,
            'align': 'left'
        })
        
        # Concepto principal (temporal) - green
        temporal_concept_format = workbook.add_format({
            'bold': True,
            'bg_color': '#E2EFDA',
            'border': 1,
            'font_size': 9,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        # Concepto principal (permanente) - orange
        permanent_concept_format = workbook.add_format({
            'bold': True,
            'bg_color': '#FCE4D6',
            'border': 1,
            'font_size': 9,
            'align': 'left',
            'valign': 'vcenter'
        })
        
        # Detalle temporal (light green) - money
        temporal_detail_money = workbook.add_format({
            'bg_color': '#F4F9F4',
            'border': 1,
            'num_format': '$ #,##0.00',
            'align': 'right',
            'font_size': 9
        })
        
        # Detalle temporal (light green) - text
        temporal_detail_text = workbook.add_format({
            'bg_color': '#F4F9F4',
            'border': 1,
            'valign': 'vcenter',
            'font_size': 9,
            'text_wrap': True
        })
        
        # Detalle permanente (light orange) - money
        permanent_detail_money = workbook.add_format({
            'bg_color': '#FFF4ED',
            'border': 1,
            'num_format': '$ #,##0.00',
            'align': 'right',
            'font_size': 9
        })
        
        # Detalle permanente (light orange) - text
        permanent_detail_text = workbook.add_format({
            'bg_color': '#FFF4ED',
            'border': 1,
            'valign': 'vcenter',
            'font_size': 9,
            'text_wrap': True
        })
        
        # Money format with borders
        money_format = workbook.add_format({
            'num_format': '$ #,##0.00',
            'border': 1,
            'align': 'right'
        })
        
        # Text format with borders
        text_format = workbook.add_format({
            'border': 1,
            'valign': 'vcenter',
            'text_wrap': True
        })
        
        # Signo format
        signo_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'border': 1,
            'font_size': 11,
            'valign': 'vcenter'
        })
        
        # === APPLY COLUMN WIDTHS ===
        worksheet.set_column('A:A', 1.5)
        worksheet.set_column('B:B', 3.5)
        worksheet.set_column('C:C', 32)
        worksheet.set_column('D:E', 1.5)
        worksheet.set_column('F:F', 65)
        worksheet.set_column('G:G', 16)
        worksheet.set_column('H:H', 10)
        worksheet.set_column('I:I', 22)
        worksheet.set_column('J:J', 14)
        worksheet.set_column('K:K', 17)
        
        # === APPLY FORMATTING ROW BY ROW ===
        for row_num, row_data in enumerate(reconciliation_rows):
            signo = row_data[1]
            concepto_principal = row_data[2]
            
            # Row 0: Title
            if row_num == 0:
                worksheet.set_row(row_num, 25)
                for col in range(13):
                    worksheet.write(row_num, col, row_data[col] if row_data[col] else '', title_format)
            
            # Row 1: Headers
            elif row_num == 1:
                worksheet.set_row(row_num, 35)
                for col in range(13):
                    worksheet.write(row_num, col, row_data[col] if row_data[col] else '', header_format)
            
            # Saldo inicial/final
            elif concepto_principal and ('Saldo final' in str(concepto_principal)):
                worksheet.set_row(row_num, 22)
                for col in range(13):
                    val = row_data[col] if row_data[col] != '' else ''
                    if col == 2:  # Concepto
                        worksheet.write(row_num, col, val, saldo_text_format)
                    elif col in [6, 8, 9, 10]:  # Money columns
                        if val != '' and val != 0:
                            worksheet.write(row_num, col, val, saldo_format)
                        else:
                            worksheet.write(row_num, col, '', saldo_format)
                    else:
                        worksheet.write(row_num, col, val, saldo_text_format)
            
            # Temporal concepts (with +/- sign)
            elif signo and concepto_principal and any(term in str(concepto_principal) for term in 
                ['Depósito', 'tránsito', 'Acredita', 'Transferencia', 'tarjeta']):
                worksheet.set_row(row_num, 20)
                worksheet.write(row_num, 1, signo, signo_format)
                worksheet.write(row_num, 2, concepto_principal, temporal_concept_format)
                for col in [0, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]:
                    val = row_data[col] if row_data[col] != '' else ''
                    if col in [6, 8, 9, 10]:
                        if val != '' and val != 0:
                            worksheet.write(row_num, col, val, money_format)
                        else:
                            worksheet.write(row_num, col, '', temporal_concept_format)
                    else:
                        worksheet.write(row_num, col, val, temporal_concept_format)
            
            # Permanent concepts (with +/- sign)
            elif signo and concepto_principal and any(term in str(concepto_principal) for term in 
                ['Notas de', 'Cheques', 'omitido', 'registrado', 'pendiente']):
                worksheet.set_row(row_num, 20)
                worksheet.write(row_num, 1, signo, signo_format)
                worksheet.write(row_num, 2, concepto_principal, permanent_concept_format)
                for col in [0, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]:
                    val = row_data[col] if row_data[col] != '' else ''
                    if col in [6, 8, 9, 10]:
                        if val != '' and val != 0:
                            worksheet.write(row_num, col, val, money_format)
                        else:
                            worksheet.write(row_num, col, '', permanent_concept_format)
                    else:
                        worksheet.write(row_num, col, val, permanent_concept_format)
            
            # Detail rows (following concepts)
            elif not signo and not concepto_principal and row_data[5]:  # Has detail in column F
                # Set row height based on text length
                text_length = len(str(row_data[5]))
                if text_length > 100:
                    worksheet.set_row(row_num, 30)
                elif text_length > 60:
                    worksheet.set_row(row_num, 20)
                else:
                    worksheet.set_row(row_num, 15)
                
                # Determine if temporal or permanent based on which column has value
                is_temporal = row_data[6] != '' and row_data[6] != 0
                money_fmt = temporal_detail_money if is_temporal else permanent_detail_money
                text_fmt = temporal_detail_text if is_temporal else permanent_detail_text
                
                for col in range(13):
                    val = row_data[col]
                    if col in [6, 8, 9, 10]:  # Money columns
                        if val != '' and val != 0:
                            worksheet.write(row_num, col, val, money_fmt)
                        else:
                            worksheet.write(row_num, col, '', text_fmt)
                    else:
                        worksheet.write(row_num, col, val if val != '' else '', text_fmt)
            
            # Default format
            else:
                worksheet.set_row(row_num, 15)
                for col in range(13):
                    val = row_data[col] if row_data[col] != '' else ''
                    if col in [6, 8, 9, 10]:
                        if val != '' and val != 0:
                            worksheet.write(row_num, col, val, money_format)
                        else:
                            worksheet.write(row_num, col, '', text_format)
                    elif col == 1:
                        worksheet.write(row_num, col, val, signo_format if val else text_format)
                    else:
                        worksheet.write(row_num, col, val, text_format)
        
        # Freeze panes (freeze first 2 rows)
        worksheet.freeze_panes(2, 0)
        
        # Set print options for better printing
        worksheet.set_landscape()
        worksheet.set_paper(9)  # A4
        worksheet.fit_to_pages(1, 0)  # Fit to 1 page wide
        
        # Set zoom to 110% for better visibility
        worksheet.set_zoom(110)

        
        # === SHEET 2: MATCHED ITEMS ===
        matched_items = merged[merged['Matched']].copy()
        if len(matched_items) > 0:
            cols_matched = ['check_match_id', 'Fecha Pago ', 'Concepto', 'Monto_Libro', 
                           'Fecha', 'Descripcion', 'Monto_Banco']
            cols_matched = [c for c in cols_matched if c in matched_items.columns]
            matched_items[cols_matched].to_excel(writer, sheet_name='Items Coincidentes', index=False)
            
            # Format Sheet 2
            ws_matched = writer.sheets['Items Coincidentes']
            
            # Header format
            header_fmt = workbook.add_format({
                'bold': True,
                'bg_color': '#4472C4',
                'font_color': 'white',
                'border': 1,
                'align': 'center',
                'font_size': 9
            })
            
            # Data format
            data_fmt = workbook.add_format({'border': 1, 'font_size': 8})
            money_fmt_small = workbook.add_format({
                'border': 1,
                'num_format': '$ #,##0.00',
                'font_size': 8
            })
            
            # Apply header format
            for col_num, value in enumerate(matched_items[cols_matched].columns.values):
                ws_matched.write(0, col_num, value, header_fmt)
            
            # Auto-fit columns
            for i, col in enumerate(cols_matched):
                max_len = max(
                    matched_items[col].astype(str).map(len).max(),
                    len(str(col))
                ) + 2
                ws_matched.set_column(i, i, min(max_len, 50))
            
            
            # Apply data format
            for row in range(1, len(matched_items) + 1):
                for col in range(len(cols_matched)):
                    val = matched_items.iloc[row-1, col]
                    # Handle NaN values
                    if pd.isna(val):
                        ws_matched.write(row, col, '', data_fmt)
                    elif 'Monto' in cols_matched[col]:
                        ws_matched.write(row, col, val, money_fmt_small)
                    else:
                        ws_matched.write(row, col, val, data_fmt)
            
            ws_matched.freeze_panes(1, 0)
            ws_matched.set_zoom(110)
        
        # === SHEET 3: TEMPORARY DIFFERENCES ===
        temp_libro = libro_unmatched[libro_unmatched['Categoria'].apply(lambda x: x.get('is_temporary', False))].copy()
        temp_extracto = extracto_unmatched[extracto_unmatched['Categoria'].apply(lambda x: x.get('is_temporary', False))].copy()
        
        if len(temp_libro) > 0 or len(temp_extracto) > 0:
            temp_all = []
            if len(temp_libro) > 0:
                temp_libro['Origen'] = 'Libro'
                temp_libro['Subcategoria'] = temp_libro['Categoria'].apply(lambda x: x.get('subcategory', ''))
                temp_all.append(temp_libro[['Fecha Pago ', 'Concepto', 'Monto_Libro', 'Subcategoria', 'Origen']])
            if len(temp_extracto) > 0:
                temp_extracto['Origen'] = 'Extracto'
                temp_extracto['Subcategoria'] = temp_extracto['Categoria'].apply(lambda x: x.get('subcategory', ''))
                temp_all.append(temp_extracto[['Fecha', 'Descripcion', 'Monto_Banco', 'Subcategoria', 'Origen']].rename(
                    columns={'Fecha': 'Fecha Pago ', 'Descripcion': 'Concepto', 'Monto_Banco': 'Monto_Libro'}
                ))
            
            df_temp = pd.concat(temp_all, ignore_index=True)
            df_temp.to_excel(writer, sheet_name='Diferencias Temporales', index=False)
            
            # Format Sheet 3
            ws_temp = writer.sheets['Diferencias Temporales']
            
            # Header format (green theme)
            header_temp_fmt = workbook.add_format({
                'bold': True,
                'bg_color': '#70AD47',
                'font_color': 'white',
                'border': 1,
                'align': 'center',
                'font_size': 9
            })
            
            # Data formats
            data_temp_fmt = workbook.add_format({
                'border': 1,
                'bg_color': '#E2EFDA',
                'font_size': 8
            })
            money_temp_fmt = workbook.add_format({
                'border': 1,
                'bg_color': '#E2EFDA',
                'num_format': '$ #,##0.00',
                'font_size': 8
            })
            
            # Apply header
            for col_num, value in enumerate(df_temp.columns.values):
                ws_temp.write(0, col_num, value, header_temp_fmt)
            
            # Auto-fit columns
            for i, col in enumerate(df_temp.columns):
                if col == 'Concepto':
                    ws_temp.set_column(i, i, 60)
                elif col == 'Monto_Libro':
                    ws_temp.set_column(i, i, 15)
                else:
                    max_len = max(df_temp[col].astype(str).map(len).max(), len(str(col))) + 2
                    ws_temp.set_column(i, i, min(max_len, 25))
            
            
            # Apply data format
            for row in range(len(df_temp)):
                for col, col_name in enumerate(df_temp.columns):
                    val = df_temp.iloc[row, col]
                    if pd.isna(val):
                        ws_temp.write(row + 1, col, '', data_temp_fmt)
                    elif col_name == 'Monto_Libro':
                        ws_temp.write(row + 1, col, val, money_temp_fmt)
                    else:
                        ws_temp.write(row + 1, col, val, data_temp_fmt)
            
            ws_temp.freeze_panes(1, 0)
            ws_temp.set_zoom(110)
        
        # === SHEET 4: PERMANENT DIFFERENCES ===
        perm_extracto = extracto_unmatched[extracto_unmatched['Categoria'].apply(lambda x: not x.get('is_temporary', True))].copy()
        if len(perm_extracto) > 0:
            perm_extracto['Subcategoria'] = perm_extracto['Categoria'].apply(lambda x: x.get('subcategory', ''))
            df_perm = perm_extracto[['Fecha', 'Descripcion', 'Monto_Banco', 'Subcategoria']]
            df_perm.to_excel(writer, sheet_name='Diferencias Permanentes', index=False)
            
            # Format Sheet 4
            ws_perm = writer.sheets['Diferencias Permanentes']
            
            # Header format (orange theme)
            header_perm_fmt = workbook.add_format({
                'bold': True,
                'bg_color': '#ED7D31',
                'font_color': 'white',
                'border': 1,
                'align': 'center',
                'font_size': 9
            })
            
            # Data formats
            data_perm_fmt = workbook.add_format({
                'border': 1,
                'bg_color': '#FCE4D6',
                'font_size': 8
            })
            money_perm_fmt = workbook.add_format({
                'border': 1,
                'bg_color': '#FCE4D6',
                'num_format': '$ #,##0.00',
                'font_size': 8
            })
            
            # Apply header
            for col_num, value in enumerate(df_perm.columns.values):
                ws_perm.write(0, col_num, value, header_perm_fmt)
            
            # Auto-fit columns
            for i, col in enumerate(df_perm.columns):
                if col == 'Descripcion':
                    ws_perm.set_column(i, i, 60)
                elif col == 'Monto_Banco':
                    ws_perm.set_column(i, i, 15)
                else:
                    max_len = max(df_perm[col].astype(str).map(len).max(), len(str(col))) + 2
                    ws_perm.set_column(i, i, min(max_len, 25))
            
            
            # Apply data format
            for row in range(len(df_perm)):
                for col, col_name in enumerate(df_perm.columns):
                    val = df_perm.iloc[row, col]
                    if pd.isna(val):
                        ws_perm.write(row + 1, col, '', data_perm_fmt)
                    elif col_name == 'Monto_Banco':
                        ws_perm.write(row + 1, col, val, money_perm_fmt)
                    else:
                        ws_perm.write(row + 1, col, val, data_perm_fmt)
            
            ws_perm.freeze_panes(1, 0)
            ws_perm.set_zoom(110)
        
        # === SHEET 5: SUMMARY ===
        summary_df = pd.DataFrame([summary]).T.reset_index()
        summary_df.columns = ['Métrica', 'Valor']
        summary_df.to_excel(writer, sheet_name='Resumen', index=False)
        
        # Format Sheet 5
        ws_summary = writer.sheets['Resumen']
        
        # Header format
        header_summary_fmt = workbook.add_format({
            'bold': True,
            'bg_color': '#5B9BD5',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'font_size': 10
        })
        
        # Data formats
        metric_fmt = workbook.add_format({
            'bold': True,
            'border': 1,
            'font_size': 9,
            'bg_color': '#DDEBF7'
        })
        value_fmt = workbook.add_format({
            'border': 1,
            'num_format': '#,##0.00',
            'font_size': 9,
            'align': 'right'
        })
        
        # Apply header
        ws_summary.write(0, 0, 'Métrica', header_summary_fmt)
        ws_summary.write(0, 1, 'Valor', header_summary_fmt)
        
        # Set column widths
        ws_summary.set_column(0, 0, 35)
        ws_summary.set_column(1, 1, 20)
        
        # Apply data format
        for row in range(len(summary_df)):
            ws_summary.write(row + 1, 0, summary_df.iloc[row, 0], metric_fmt)
            ws_summary.write(row + 1, 1, summary_df.iloc[row, 1], value_fmt)
        
        ws_summary.freeze_panes(1, 0)
        ws_summary.set_zoom(110)
        
    output.seek(0)
    
    return output, summary

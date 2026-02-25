# 2. Map Columns (Updated for Merged Cells & Sub-headers)
        col_map = {}
        for row in ws.iter_rows(min_row=1, max_row=15): # Headers are usually in the top 15 rows
            for cell in row:
                if not cell.value: continue
                val = normalize_text(str(cell.value))
                
                if 'abarrotes' in val: col_map['abar'] = cell.column
                if 'agricultura' in val: col_map['agri'] = cell.column
                if 'escuela' in val or 'establecimiento' in val: col_map['escuelas'] = cell.column
                
                # --- NEW LOGIC FOR MERGED PROVEEDORES HEADER ---
                if 'proveedor' in val or 'productor' in val:
                    base_col = cell.column
                    base_row = cell.row
                    
                    # Look at the row immediately below to find exactly which column is "Total"
                    # We check the current column and the next 2 columns to cover the merged area
                    for offset in range(3): 
                        sub_cell = ws.cell(row=base_row + 1, column=base_col + offset)
                        sub_val = normalize_text(str(sub_cell.value))
                        
                        if 'total' in sub_val:
                            col_map['productores'] = sub_cell.column
                            break
                    
                    # Fallback: If 'total' wasn't explicitly found underneath, default to the main column
                    if 'productores' not in col_map:
                        col_map['productores'] = base_col

        if 'abar' not in col_map or 'agri' not in col_map:
            st.error(f"No encontré las columnas base. Columnas detectadas: {col_map}")
            st.stop()

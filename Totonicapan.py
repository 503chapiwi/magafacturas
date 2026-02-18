import pdfplumber
import openpyxl
import unicodedata
import streamlit as st
import re
import io

# --- HELPER FUNCTIONS ---
def normalize_text(text):
    if not text: return ""
    nfd = unicodedata.normalize('NFD', str(text))
    return ''.join(char for char in nfd if unicodedata.category(char) != 'Mn').lower()

def clean_currency(value):
    if not value: return 0.0
    # Handles 1.200,00 or 12:00 OCR errors
    raw = str(value).replace(':', '.').replace(',', '').strip()
    try:
        return float(raw)
    except ValueError:
        return 0.0

# --- WEB UI ---
st.title("🇬🇹 MAGA: Procesador de Facturas por la LAE")
uploaded_pdfs = st.file_uploader(label='1. Seleccione sus Facturas (PDFs)', type='pdf', accept_multiple_files=True,
                                 help='Arrastre y suelte sus facturas aquí. El límite es 200mb por archivo')
uploaded_xlsx = st.file_uploader(label='2. Seleccione su Archivo de Excel', type='xlsx',
                                 help='Arraste y suelte el archivo de Excel dónde van los totales de las facturas')

if st.button("INICIAR PROCESO") and uploaded_pdfs and uploaded_xlsx:
    try:
        input_buffer = io.BytesIO(uploaded_xlsx.read())
        wb = openpyxl.load_workbook(input_buffer)
        ws = wb.active 
        
        # 1. Setup/Load "Extra Detalles"
        if "Extra Detalles" not in wb.sheetnames:
            ws_det = wb.create_sheet("Extra Detalles")
            ws_det.append(['Nombre Emisor', 'NIT Emisor', 'NIT Receptor', 'UUID', 'Municipio', 'Alerta % Abarrotes'])
        else:
            ws_det = wb["Extra Detalles"]

        processed_uuids = set()
        for row in ws_det.iter_rows(min_row=2, min_col=4, max_col=4, values_only=True):
            if row[0]: processed_uuids.add(str(row[0]).strip())

        # 2. IMPROVED COLUMN MAPPING
        # In your file, these are in Row 6. We search for the specific MAGA wording.
        col_map = {}
        for row in ws.iter_rows(min_row=1, max_row=50): 
            for cell in row:
                if not cell.value: continue
                val = normalize_text(cell.value)
                if 'abarrotes' in val:
                    col_map['abar'] = cell.column
                if 'agricultura' in val:
                    col_map['agri'] = cell.column

        if 'abar' not in col_map or 'agri' not in col_map:
            st.error(f"No encontré las columnas. Columnas detectadas: {col_map}")
            st.stop()

        muni_map = {'totonican, totonicapan': 1, 'san cristobal': 2, 'san francisco': 3, 'san andres xecul': 4,
                    'momostenango': 5, 'santa maria chiquimula': 6, 'santa lucia la reforma': 7, 'san bartolo': 8}

        new_count = 0
        progress_bar = st.progress(0)

        # 3. Process each PDF
        for i, pdf_file in enumerate(uploaded_pdfs):
            with pdfplumber.open(pdf_file) as pdf:
                text = "".join([p.extract_text() or "" for p in pdf.pages])
                tables = []
                for p in pdf.pages:
                    t = p.extract_table()
                    if t: tables.extend(t)

                uuid_m = re.search(r'[A-F0-9]{8}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{12}', text, re.I)
                uuid_val = uuid_m.group(0) if uuid_m else pdf_file.name

                if uuid_val in processed_uuids: continue

                norm_text = normalize_text(text)
                m_id = next((v for k, v in muni_map.items() if k in norm_text), None)
                m_name = next((k for k in muni_map if k in norm_text), "N/A")

                if m_id:
                    abar_sum, agri_sum = 0, 0
                    cultivados = ['tomate', 'pina', 'banano', 'zanahoria', 'guisquil', 'cebolla', 'aguacate', 'miltomate']
                    abarrotes = ['pollo', 'tostadas']
                    
                    for row_tbl in tables:
                        if not row_tbl or len(row_tbl) < 8 or not row[3]: continue
                        desc = normalize_text(row_tbl[3])
                        val = clean_currency(row_tbl[7])
                        if any(x in desc for x in cultivados): 
                            agri_sum += val
                        if any(x in desc for x in abarrotes):
                            abar_sum += val
                    
                    # --- THE FIX: Robust Row Selection (Grid Search) ---
                    found_row = False
                    
                    # Look through the rows
                    for row_ex in ws.iter_rows(min_row=1, max_row=300, min_col=1, max_col=10):
                        # Look through every cell in the current row
                        for cell in row_ex:
                            if cell.value is None:
                                continue
                            
                            # Clean the cell value two different ways:
                            # 1. As a number (to match "7" or "7.0")
                            clean_cell_num = str(cell.value).split('.')[0].strip()
                            # 2. As normalized text (to match "santa lucia la reforma")
                            clean_cell_text = normalize_text(str(cell.value))
                            
                            target_id = str(m_id).strip()
                            target_name = str(m_name).strip()
                            
                            # MATCH LOGIC: Does the cell match the ID (7) OR the Name?
                            if clean_cell_num == target_id or target_name in clean_cell_text:
                                r_idx = cell.row
                                
                                # Get current values safely
                                try:
                                    val_abar = ws.cell(row=r_idx, column=col_map['abar']).value
                                    val_agri = ws.cell(row=r_idx, column=col_map['agri']).value
                                    
                                    current_abar = float(val_abar) if val_abar else 0.0
                                    current_agri = float(val_agri) if val_agri else 0.0
                                    
                                    # Write the NEW values
                                    ws.cell(row=r_idx, column=col_map['abar']).value = current_abar + abar_sum
                                    ws.cell(row=r_idx, column=col_map['agri']).value = current_agri + agri_sum
                                    
                                    st.write(f"✅ Fila {r_idx} actualizada para {m_name}")
                                    found_row = True
                                    break # Stop looking at other cells in this row
                                except Exception as e:
                                    st.error(f"Error escribiendo en fila {r_idx}: {e}")
                                    break
                        
                        if found_row:
                            break # Stop looking at other rows for this PDF

                    if not found_row:
                        st.warning(f"⚠️ No encontré ni el ID '{m_id}' ni el nombre '{m_name}' en el Excel.")
                    
                    # 4. Alert & Metadata
                    total_rec = abar_sum + agri_sum
                    perc_abar = (abar_sum / total_rec) if total_rec > 0 else 0
                    alert_status = "⚠️ ALERTA: >30%" if perc_abar > 0.30 else "OK"

                    nit_e = re.search(r'Emisor:\s*(\d+)', text, re.I)
                    nit_r = re.search(r'Receptor:\s*(\d+)', text, re.I)
                    name_e = re.search(r'(?:Factura|Contribuyente)\s*\n?([^\n\d]+)', text)

                    ws_det.append([
                        name_e.group(1).strip() if name_e else "N/A",
                        nit_e.group(1) if nit_e else "N/A",
                        nit_r.group(1) if nit_r else "N/A",
                        uuid_val, m_name, alert_status
                    ])
                    new_count += 1
                    processed_uuids.add(uuid_val)

            progress_bar.progress((i + 1) / len(uploaded_pdfs))

        # 5. Final Export
        output = io.BytesIO()
        wb.save(output)
        st.success(f"¡Éxito! {new_count} facturas procesadas correctamente.")
        output.seek(0)
        st.download_button("Descargar Reporte Final", data=output.getvalue(), 
                           file_name="Reporte_MAGA_Actualizado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error detectado: {e}")
                    
                    # 4. Alert & Metadata
                    total_rec = abar_sum + agri_sum
                    perc_abar = (abar_sum / total_rec) if total_rec > 0 else 0
                    alert_status = "⚠️ ALERTA: >30%" if perc_abar > 0.30 else "OK"

                    nit_e = re.search(r'Emisor:\s*(\d+)', text, re.I)
                    nit_r = re.search(r'Receptor:\s*(\d+)', text, re.I)
                    name_e = re.search(r'(?:Factura|Contribuyente)\s*\n?([^\n\d]+)', text)

                    ws_det.append([
                        name_e.group(1).strip() if name_e else "N/A",
                        nit_e.group(1) if nit_e else "N/A",
                        nit_r.group(1) if nit_r else "N/A",
                        uuid_val, m_name, alert_status
                    ])
                    new_count += 1
                    processed_uuids.add(uuid_val)

            progress_bar.progress((i + 1) / len(uploaded_pdfs))

        # 5. Final Export
        output = io.BytesIO()
        wb.save(output)
        st.success(f"¡Éxito! {new_count} facturas procesadas correctamente.")
        output.seek(0)
        st.download_button("Descargar Reporte Final", data=output.getvalue(), 
                           file_name="Reporte_MAGA_Actualizado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error detectado: {e}")

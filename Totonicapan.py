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
    raw = str(value).replace(':', '.').replace(',', '').strip()
    try:
        return float(raw)
    except ValueError:
        return 0.0

# --- WEB UI ---
st.title("🇬🇹 MAGA: Procesador de Alimentación Escolar")
uploaded_pdfs = st.file_uploader("1. Subir Facturas (PDF)", type="pdf", accept_multiple_files=True)
uploaded_xlsx = st.file_uploader("2. Subir Reporte Excel", type="xlsx")

if st.button("INICIAR PROCESO") and uploaded_pdfs and uploaded_xlsx:
    try:
        input_buffer = io.BytesIO(uploaded_xlsx.read())
        wb = openpyxl.load_workbook(input_buffer)
        ws = wb.active 
        
        # 1. Setup "Extra Detalles"
        if "Extra Detalles" not in wb.sheetnames:
            ws_det = wb.create_sheet("Extra Detalles")
            ws_det.append(['Nombre Emisor', 'NIT Emisor', 'NIT Receptor', 'UUID', 'Municipio', 'Alerta % Abarrotes'])
        else:
            ws_det = wb["Extra Detalles"]

        processed_uuids = set()
        for row in ws_det.iter_rows(min_row=2, min_col=4, max_col=4, values_only=True):
            if row[0]: processed_uuids.add(str(row[0]).strip())

        # 2. FIXED: Expanded Column Mapping (Scans rows 1-10)
        col_map = {}
        for row in ws.iter_rows(min_row=1, max_row=10): 
            for cell in row:
                val = normalize_text(cell.value)
                if 'abarrotes' in val: col_map['abar'] = cell.column
                if 'agricultura familiar' in val: col_map['agri'] = cell.column

        if 'abar' not in col_map or 'agri' not in col_map:
            st.error("No se encontraron las columnas 'Procesados abarrotes' o 'Agricultura Familiar' en el Excel.")
            st.stop()

        muni_map = {'totonican': 1, 'cristobal': 2, 'francisco': 3, 'xecul': 4,
                    'momostenango': 5, 'chiquimula': 6, 'reforma': 7, 'bartolo': 8}

        new_count = 0
        progress_bar = st.progress(0)

        # 3. Process PDFs
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
                    grown = ['tomate', 'pina', 'banano', 'zanahoria', 'guisquil', 'cebolla', 'aguacate', 'miltomate']
                    
                    for row_tbl in tables:
                        if len(row_tbl) < 8: continue
                        desc = normalize_text(row_tbl[3])
                        val = clean_currency(row_tbl[7])
                        if any(x in desc for x in grown): agri_sum += val
                        else: abar_sum += val
                    
                    # --- FIXED: Direct Row/Cell Update ---
                    found_row = False
                    for row_ex in ws.iter_rows(min_row=1, max_row=100):
                        # Ensure we check Column A (index 0) for the ID number
                        if str(row_ex[0].value).strip() == str(m_id):
                            r_idx = row_ex[0].row
                            
                            # Add values to existing ones
                            curr_abar = ws.cell(row=r_idx, column=col_map['abar']).value or 0
                            curr_agri = ws.cell(row=r_idx, column=col_map['agri']).value or 0
                            
                            ws.cell(row=r_idx, column=col_map['abar']).value = curr_abar + abar_sum
                            ws.cell(row=r_idx, column=col_map['agri']).value = curr_agri + agri_sum
                            found_row = True
                            break
                    
                    if not found_row:
                        st.warning(f"No se encontró la fila para el ID {m_id} en el Excel.")

                    # Metadata with >30% Alert
                    total_rec = abar_sum + agri_sum
                    perc_abar = (abar_sum / total_rec) if total_rec > 0 else 0
                    alert_status = "⚠️ ALERTA: >30%" if perc_abar > 0.30 else "OK"

                    nit_e = re.search(r'Emisor:\s*(\d+)', text, re.I)
                    nit_r = re.search(r'Receptor:\s*(\d+)', text, re.I)
                    name_e = re.search(r'Contribuyente\n([^\n]+)', text)

                    ws_det.append([
                        name_e.group(1).strip() if name_e else "N/A",
                        nit_e.group(1) if nit_e else "N/A",
                        nit_r.group(1) if nit_r else "N/A",
                        uuid_val, m_name, alert_status
                    ])
                    new_count += 1
                    processed_uuids.add(uuid_val)

            progress_bar.progress((i + 1) / len(uploaded_pdfs))

        # 4. Save to Buffer
        output = io.BytesIO()
        wb.save(output)
        st.success(f"Procesado: {new_count} facturas.")
        st.download_button("Descargar Reporte Actualizado", data=output.getvalue(), 
                           file_name="Reporte_Final.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error técnico: {e}")

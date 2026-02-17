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
        ws = wb.active # Main sheet (e.g., "Hoja1" or "RESUMEN")
        
        # 1. Setup/Load "Extra Detalles"
        if "Extra Detalles" not in wb.sheetnames:
            ws_det = wb.create_sheet("Extra Detalles")
            ws_det.append(['Nombre Emisor', 'NIT Emisor', 'NIT Receptor', 'UUID', 'Municipio', 'Alerta % Abarrotes'])
        else:
            ws_det = wb["Extra Detalles"]

        processed_uuids = set()
        for row in ws_det.iter_rows(min_row=2, min_col=4, max_col=4, values_only=True):
            if row[0]: processed_uuids.add(str(row[0]).strip())

        # 2. Map Columns in Main Sheet
        col_map = {}
        for row in ws.iter_rows(min_row=1, max_row=10): 
            for cell in row:
                val = normalize_text(cell.value)
                if 'abarrotes' in val: col_map['abar'] = cell.column
                if 'agricultura' in val: col_map['agri'] = cell.column

        muni_map = {'totonican': 1, 'cristobal': 2, 'francisco': 3, 'xecul': 4,
                    'momostenango': 5, 'chiquimula': 6, 'reforma': 7, 'bartolo': 8}

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
                    grown = ['tomate', 'pina', 'banano', 'zanahoria', 'guisquil', 'cebolla', 'aguacate', 'miltomate']
                    
                    for row in tables:
                        if len(row) < 8: continue
                        desc = normalize_text(row[3])
                        val = clean_currency(row[7])
                        if any(x in desc for x in grown): agri_sum += val
                        else: abar_sum += val
                    
                    # --- THE FIX: Update Main Sheet Totals ---
                    for row in ws.iter_rows(min_row=1, max_row=100):
                        # Match the ID in Column A
                        if str(row[0].value) == str(m_id):
                            r_idx = row[0].row
                            if 'abar' in col_map:
                                current = ws.cell(r_idx, col_map['abar']).value or 0
                                ws.cell(r_idx, col_map['abar']).value = current + abar_sum
                            if 'agri' in col_map:
                                current = ws.cell(r_idx, col_map['agri']).value or 0
                                ws.cell(r_idx, col_map['agri']).value = current + agri_sum

                    # --- ADD ALERT SYSTEM ---
                    total_rec = abar_sum + agri_sum
                    perc_abar = (abar_sum / total_rec) if total_rec > 0 else 0
                    alert_status = "⚠️ ALERTA: >30%" if perc_abar > 0.30 else "OK"

                    # Metadata
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

        # 4. Final Save to Buffer
        output = io.BytesIO()
        wb.save(output)
        st.success(f"¡Hecho! Se procesaron {new_count} facturas.")
        st.download_button("Descargar Reporte Actualizado", data=output.getvalue(), 
                           file_name="Reporte_Final.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error técnico: {e}")

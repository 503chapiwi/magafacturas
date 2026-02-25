import pdfplumber
import openpyxl
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter
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

# --- TRUCO CSS PARA TRADUCIR LA INTERFAZ A ESPAÑOL ---
st.markdown("""
    <style>
        .stFileUploader > div > div > div > div > span:first-child {
            display: none;
        }
        .stFileUploader > div > div > div > div::before {
            content: "Arrastre y suelte los archivos aquí";
            display: block;
            font-weight: 600;
            margin-bottom: 5px;
        }
    </style>
""", unsafe_allow_html=True)

# --- WEB UI ---
st.title("🇬🇹 MAGA: Procesador de Facturas por la LAE")
uploaded_pdfs = st.file_uploader(label='1. Seleccione sus Facturas (PDFs)', type='pdf', accept_multiple_files=True,
                                 help='Arrastre y suelte sus facturas aquí. El límite es 200mb por archivo')
uploaded_xlsx = st.file_uploader(label='2. Seleccione su Archivo de Excel', type='xlsx',
                                 help='Arrastre y suelte el archivo de Excel dónde van los totales de las facturas')

if st.button("INICIAR PROCESO") and uploaded_pdfs and uploaded_xlsx:
    try:
        input_buffer = io.BytesIO(uploaded_xlsx.read())
        wb = openpyxl.load_workbook(input_buffer)
        ws = wb.active 
        
        # 1. Setup "Extra Detalles" Sheet
        if "Extra Detalles" not in wb.sheetnames:
            ws_det = wb.create_sheet("Extra Detalles")
            ws_det.append(['Nombre Emisor', 'NIT Emisor', 'NIT Receptor', 'UUID', 'Municipio', 'Alerta % Abarrotes'])
        else:
            ws_det = wb["Extra Detalles"]

        processed_uuids = set()
        for row in ws_det.iter_rows(min_row=2, min_col=4, max_col=4, values_only=True):
            if row[0]: processed_uuids.add(str(row[0]).strip())

        # 2. Map Columns 
        col_map = {}
        for row in ws.iter_rows(min_row=1, max_row=15): 
            for cell in row:
                if not cell.value: continue
                val = normalize_text(str(cell.value))
                
                if 'abarrotes' in val: col_map['abar'] = cell.column
                if 'agricultura' in val: col_map['agri'] = cell.column
                if 'escuela' in val or 'establecimiento' in val: col_map['escuelas'] = cell.column
                
                # Logic for merged Proveedores header
                if 'proveedor' in val or 'productor' in val:
                    base_col = cell.column
                    base_row = cell.row
                    
                    for offset in range(3): 
                        sub_cell = ws.cell(row=base_row + 1, column=base_col + offset)
                        sub_val = normalize_text(str(sub_cell.value))
                        
                        if 'total' in sub_val:
                            col_map['productores'] = sub_cell.column
                            break
                    
                    if 'productores' not in col_map:
                        col_map['productores'] = base_col

        if 'abar' not in col_map or 'agri' not in col_map:
            st.error(f"No encontré las columnas base. Columnas detectadas: {col_map}")
            st.stop()

        muni_map = {'totonican totonicapan': 1, 'san cristobal totonicapan': 2, 'san francisco el alto': 3, 'san andres xecul': 4,
                    'momostenango': 5, 'santa maria chiquimula': 6, 'santa lucia la reforma': 7, 'san bartolo': 8}

        batch_totals = {m_id: {'abar': 0.0, 'agri': 0.0, 'emisores': set(), 'receptores': set()} for m_id in muni_map.values()}
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
                
                # --- THE FIX: Strip all whitespace and commas for a bulletproof municipality match ---
                squished_text = re.sub(r'[\s,]+', '', norm_text)
                
                m_id = next((v for k, v in muni_map.items() if re.sub(r'[\s,]+', '', k) in squished_text), None)
                m_name = next((k for k in muni_map if re.sub(r'[\s,]+', '', k) in squished_text), "N/A")

                if m_id:
                    abar_sum, agri_sum = 0, 0
                    cultivados = ['tomate', 'pina', 'banano', 'zanahoria', 'guisquil', 'cebolla', 'aguacate', 'miltomate']
                    abarrotes = ['pollo', 'tostadas']
                    
                    for row_tbl in tables:
                        if not row_tbl: continue
                        
                        row_text = " ".join([normalize_text(str(x)) for x in row_tbl if x])
                        
                        val = 0.0
                        if len(row_tbl) >= 8:
                            val = clean_currency(row_tbl[7])
                        elif len(row_tbl) >= 7:
                            val = clean_currency(row_tbl[6])
                        elif len(row_tbl) >= 4:
                            val = clean_currency(row_tbl[-1]) 
                            
                        if any(x in row_text for x in cultivados): agri_sum += val
                        if any(x in row_text for x in abarrotes): abar_sum += val
                    
                    nit_e_match = re.search(r'Emisor:\s*([0-9Kk\-]+)', text, re.I)
                    nit_r_match = re.search(r'Receptor:\s*([0-9Kk\-]+)', text, re.I)
                    name_e_match = re.search(r'(?:Factura(?:\s*Pequeño\s*Contribuyente)?)\s*\n+(.*?)\n+Nit\s*Emisor', text, re.IGNORECASE)
                    
                    nit_e = nit_e_match.group(1).strip() if nit_e_match else "N/A"
                    nit_r = nit_r_match.group(1).strip() if nit_r_match else "N/A"
                    
                    raw_name = name_e_match.group(1).strip() if name_e_match else "N/A"
                    name_e = re.split(r'(?i)n[úu]mero\s*de\s*autorizaci[óo]n', raw_name)[0]
                    name_e = re.split(r'(?i)\bserie\b', name_e)[0].strip()

                    batch_totals[m_id]['abar'] += abar_sum
                    batch_totals[m_id]['agri'] += agri_sum
                    if nit_e != "N/A": batch_totals[m_id]['emisores'].add(nit_e)
                    if nit_r != "N/A": batch_totals[m_id]['receptores'].add(nit_r)

                    total_rec = abar_sum + agri_sum
                    perc_abar = (abar_sum / total_rec) if total_rec > 0 else 0
                    alert_status = "⚠️ ALERTA: >30%" if perc_abar > 0.30 else "OK"

                    ws_det.append([name_e, nit_e, nit_r, uuid_val, m_name, alert_status])
                    processed_uuids.add(uuid_val)
                    new_count += 1
                else:
                    st.warning(f"No se pudo identificar el municipio en la factura: {pdf_file.name}")

            progress_bar.progress((i + 1) / len(uploaded_pdfs))

        # 4. Write accumulated data to Main Sheet
        for row_ex in ws.iter_rows(min_row=1, max_row=200):
            cell_a_val = str(row_ex[0].value).strip() if row_ex[0].value is not None else ""
            if not cell_a_val: continue

            try:
                excel_m_id = int(float(cell_a_val))
                if excel_m_id in batch_totals:
                    r_idx = row_ex[0].row
                    data = batch_totals[excel_m_id]

                    curr_abar = ws.cell(r_idx, col_map['abar']).value
                    ws.cell(r_idx, col_map['abar']).value = (float(curr_abar) if curr_abar else 0.0) + data['abar']
                    
                    curr_agri = ws.cell(r_idx, col_map['agri']).value
                    ws.cell(r_idx, col_map['agri']).value = (float(curr_agri) if curr_agri else 0.0) + data['agri']

                    if 'escuelas' in col_map:
                        curr_esc = ws.cell(r_idx, col_map['escuelas']).value
                        ws.cell(r_idx, col_map['escuelas']).value = (int(curr_esc) if curr_esc else 0) + len(data['receptores'])
                    
                    if 'productores' in col_map:
                        curr_prod = ws.cell(r_idx, col_map['productores']).value
                        ws.cell(r_idx, col_map['productores']).value = (int(curr_prod) if curr_prod else 0) + len(data['emisores'])

            except (ValueError, TypeError):
                continue

        # 5. Format "Extra Detalles" (Auto-width and Borders)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                             top=Side(style='thin'), bottom=Side(style='thin'))

        for col in ws_det.columns:
            max_length = 0
            col_letter = get_column_letter(col[0].column) 
            
            for cell in col:
                cell.border = thin_border 
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            ws_det.column_dimensions[col_letter].width = max_length + 2

        # 6. Final Export
        output = io.BytesIO()
        wb.save(output)
        st.success(f"¡Éxito! {new_count} facturas procesadas correctamente.")
        output.seek(0)
        st.download_button("Descargar Reporte Final", data=output.getvalue(), 
                           file_name="Reporte_MAGA_Actualizado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Error detectado: {e}")

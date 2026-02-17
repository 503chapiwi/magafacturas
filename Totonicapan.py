import pdfplumber
import openpyxl
import unicodedata
import streamlit as st
import re
import io
from pathlib import Path

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

# --- WEB UI ELEMENTS ---
st.title("MAGA - Procesador de Facturas v1.0")
st.write("Procesador de Datos de Alimentación Escolar")

# Instead of filedialog, we use file_uploader
uploaded_pdfs = st.file_uploader("1. Selecciona los archivos PDF", type="pdf", accept_multiple_files=True)
uploaded_xlsx = st.file_uploader("2. Selecciona el Excel de Reporte", type="xlsx")

if st.button("INICIAR PROCESO"):
    if not uploaded_pdfs or not uploaded_xlsx:
        st.error("Por favor, sube tanto los PDFs como el archivo Excel.")
    else:
        try:
            # Load Excel from the uploaded file buffer
            input_buffer = io.BytesIO(uploaded_xlsx.read())
            wb = openpyxl.load_workbook(input_buffer)
            
            processed_uuids = set()
            if "Extra Detalles" in wb.sheetnames:
                ws_det = wb["Extra Detalles"]
                for row in ws_det.iter_rows(min_row=2, min_col=4, max_col=4, values_only=True):
                    if row[0]: processed_uuids.add(str(row[0]).strip())
            else:
                ws_det = wb.create_sheet("Extra Detalles")
                ws_det.append(['Nombre Emisor', 'NIT Emisor', 'NIT Receptor', 'UUID', 'Municipio'])

            muni_map = {'totonican': 1, 'cristobal': 2, 'francisco': 3, 'xecul': 4,
                        'momostenango': 5, 'chiquimula': 6, 'reforma': 7, 'bartolo': 8}

            new_count = 0
            total_files = len(uploaded_pdfs)
            
            # Progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, pdf_file in enumerate(uploaded_pdfs):
                status_text.text(f"Procesando {i+1} de {total_files}...")
                
                with pdfplumber.open(pdf_file) as pdf:
                    text = "".join([p.extract_text() or "" for p in pdf.pages])
                    tables = []
                    for p in pdf.pages:
                        t = p.extract_table()
                        if t: tables.extend(t)

                    uuid_m = re.search(r'[A-F0-9]{8}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{12}', text, re.I)
                    uuid_val = uuid_m.group(0) if uuid_m else pdf_file.name

                    if uuid_val in processed_uuids:
                        continue

                    norm_text = normalize_text(text)
                    m_id = next((v for k, v in muni_map.items() if k in norm_text), None)
                    m_name = next((k for k in muni_map if k in norm_text), "N/A")

                    if m_id:
                        abar, agri = 0, 0
                        grown = ['tomate', 'pina', 'banano', 'zanahoria', 'guisquil']
                        for row in tables:
                            if len(row) < 8: continue
                            desc = normalize_text(row[3])
                            val = clean_currency(row[7])
                            if any(x in desc for x in grown): agri += val
                            else: abar += val

                        nit_e = re.search(r'Emisor:\s*(\d+)', text, re.I)
                        nit_r = re.search(r'Receptor:\s*(\d+)', text, re.I)
                        name_e = re.search(r'Contribuyente\n([^\n]+)', text)

                        ws_det.append([
                            name_e.group(1).strip() if name_e else "N/A",
                            nit_e.group(1) if nit_e else "N/A",
                            nit_r.group(1) if nit_r else "N/A",
                            uuid_val, m_name
                        ])
                        new_count += 1
                        processed_uuids.add(uuid_val)
                
                progress_bar.progress((i + 1) / total_files)

            # --- SAVE AND DOWNLOAD ---
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            
            st.success(f"Proceso finalizado. Facturas nuevas: {new_count}")
            
            st.download_button(
                label="Descargar Excel Procesado",
                data=output_buffer.getvalue(),
                file_name="Reporte_Finalizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Ocurrió un problema: {str(e)}")

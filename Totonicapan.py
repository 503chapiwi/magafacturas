import pdfplumber
import openpyxl
import unicodedata
import tkinter as tk
import re
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

# --- HELPER FUNCTIONS ---
def normalize_text(text):
    if not text: return ""
    nfd = unicodedata.normalize('NFD', str(text))
    return ''.join(char for char in nfd if unicodedata.category(char) != 'Mn').lower()

def clean_currency(value):
    if not value: return 0.0
    # Fixes OCR errors like 12:00 or 1,078.00 [cite: 15]
    raw = str(value).replace(':', '.').replace(',', '').strip()
    try:
        return float(raw)
    except ValueError:
        return 0.0

# --- MAIN APP CLASS ---
class MinistryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MAGA - Procesador de Facturas v1.0")
        self.root.geometry("500x300")

        # UI Elements
        self.label = tk.Label(root, text="Procesador de Datos de Alimentación Escolar", font=("Arial", 12, "bold"))
        self.label.pack(pady=10)

        self.btn_run = tk.Button(root, text="INICIAR PROCESO", command=self.run_process,
                                 bg="#2ecc71", fg="white", font=("Arial", 10, "bold"), height=2, width=20)
        self.btn_run.pack(pady=20)

        self.status_label = tk.Label(root, text="Estado: Esperando archivos...", fg="blue")
        self.status_label.pack(pady=5)

        self.progress = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
        self.progress.pack(pady=10)

    def update_status(self, text):
        self.status_label.config(text=f"Estado: {text}")
        self.root.update_idletasks()

    def run_process(self):
        # 1. Select Folders
        pdffolder = filedialog.askdirectory(title="1. Selecciona Carpeta con PDFs")
        if not pdffolder: return

        xlsx_path = filedialog.askopenfilename(title="2. Selecciona el Excel de Reporte", filetypes=[("Excel", "*.xlsx")])
        if not xlsx_path: return

        try:
            self.update_status("Escaneando archivos...")
            all_pdfs = list(Path(pdffolder).rglob('*.pdf'))
            total_files = len(all_pdfs)
            self.progress["maximum"] = total_files

            # Load Excel and existing UUIDs
            wb = openpyxl.load_workbook(xlsx_path)
            processed_uuids = set()
            if "Extra Detalles" in wb.sheetnames:
                ws_det = wb["Extra Detalles"]
                for row in ws_det.iter_rows(min_row=2, min_col=4, max_col=4, values_only=True):
                    if row[0]: processed_uuids.add(str(row[0]).strip())
            else:
                ws_det = wb.create_sheet("Extra Detalles")
                ws_det.append(['Nombre Emisor', 'NIT Emisor', 'NIT Receptor', 'UUID', 'Municipio'])

            # Dictionary mapping based on your data [cite: 1, 2]
            muni_map = {'totonican': 1, 'cristobal': 2, 'francisco': 3, 'xecul': 4,
                        'momostenango': 5, 'chiquimula': 6, 'reforma': 7, 'bartolo': 8}

            summary_data = {}
            new_count = 0

            # 2. Loop through PDFs
            for i, pdf_path in enumerate(all_pdfs):
                self.update_status(f"Procesando {i+1} de {total_files}...")
                self.progress["value"] = i + 1

                with pdfplumber.open(pdf_path) as pdf:
                    text = "".join([p.extract_text() or "" for p in pdf.pages])
                    tables = []
                    for p in pdf.pages:
                        t = p.extract_table()
                        if t: tables.extend(t)

                    # Extract UUID [cite: 12, 35]
                    uuid_m = re.search(r'[A-F0-9]{8}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{12}', text, re.I)
                    uuid_val = uuid_m.group(0) if uuid_m else pdf_path.name

                    if uuid_val in processed_uuids: continue

                    # Classification logic
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

                        # Metadata
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

            # 3. Save
            wb.save(xlsx_path)
            messagebox.showinfo("Éxito", f"Proceso finalizado.\nFacturas nuevas: {new_count}")
            self.update_status("Finalizado con éxito.")

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un problema: {str(e)}")
            self.update_status("Error en el proceso.")

if __name__ == "__main__":
    root = tk.Tk()
    app = MinistryApp(root)
    root.mainloop()

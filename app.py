import os
import re
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def procesar_pdf():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    pdf_path = filedialog.askopenfilename(
        title="Seleccioná el archivo PDF de liquidación SADAIC",
        filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")]
    )

    if not pdf_path:
        return

    data_liquidacion = []
    data_resumen = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue

                    header_idx = -1
                    col_titulo = -1
                    col_pct = -1
                    col_cant = -1
                    col_neto = -1

                    # Identificar los índices exactos de cada columna en el encabezado
                    for i, row in enumerate(table):
                        row_clean = [str(c).replace('\n', ' ').strip() if c else '' for c in row]
                        row_str = " ".join(row_clean)

                        if "Título Obra" in row_str or "Titulo Obra" in row_str:
                            header_idx = i
                            for c_idx, cell in enumerate(row_clean):
                                if "Título" in cell or "Titulo" in cell or "Obra" in cell:
                                    if col_titulo == -1:
                                        col_titulo = c_idx
                                elif "%" in cell:
                                    col_pct = c_idx
                                elif "Cant" in cell:
                                    col_cant = c_idx
                                elif "Neto" in cell:
                                    col_neto = c_idx
                            break

                    # Extraer únicamente los datos pertenecientes a cada columna
                    if header_idx != -1 and col_titulo != -1:
                        for row in table[header_idx + 1:]:
                            if not row or all(cell is None or str(cell).strip() == "" for cell in row):
                                continue

                            row_str = " ".join([str(c) for c in row if c])
                            if "Título Obra" in row_str or "Titulo Obra" in row_str or "Total" in row_str or "Concepto:" in row_str:
                                continue

                            # 1. TÍTULO OBRA: Unificar renglones dobles de la celda y limpiar saltos de línea
                            raw_titulo = str(row[col_titulo]) if col_titulo < len(row) and row[col_titulo] else ""
                            titulo_limpio = " ".join(raw_titulo.split()).strip()

                            if not titulo_limpio:
                                continue

                            # 2. PORCENTAJE (%)
                            pct_val = ""
                            if col_pct != -1 and col_pct < len(row) and row[col_pct]:
                                pct_val = str(row[col_pct]).replace('\n', ' ').strip()
                            else:
                                match_pct = re.search(r'(\d{1,3}\.\d{2})', row_str)
                                pct_val = match_pct.group(1) if match_pct else ""

                            # 3. CANTIDAD
                            cant_val = ""
                            if col_cant != -1 and col_cant < len(row) and row[col_cant]:
                                cant_val = str(row[col_cant]).replace('\n', ' ').strip()
                            else:
                                cant_val = "1"

                            # 4. NETO
                            neto_val = ""
                            if col_neto != -1 and col_neto < len(row) and row[col_neto]:
                                neto_val = str(row[col_neto]).replace('\n', ' ').strip()
                            else:
                                match_neto = re.search(r'([\d\.,\-]+)$', row_str)
                                neto_val = match_neto.group(1) if match_neto else ""

                            # Pestaña Liquidación (Valores exactos sin modificar)
                            data_liquidacion.append({
                                "Título Obra": titulo_limpio,
                                "%": pct_val,
                                "Cant.": cant_val,
                                "Neto": neto_val
                            })

                            # Pestaña Resumen (Conversión numérica para sumar)
                            try:
                                neto_num = float(neto_val.replace('.', '').replace(',', '.'))
                            except ValueError:
                                neto_num = 0.0

                            try:
                                cant_num = int(cant_val)
                            except ValueError:
                                cant_num = 1

                            data_resumen.append({
                                "Título Obra": titulo_limpio,
                                "Cant.": cant_num,
                                "Neto": neto_num
                            })

        if not data_liquidacion:
            messagebox.showwarning(
                "Aviso", 
                "No se pudieron extraer los datos. Verificá que el PDF contenga las tablas de liquidación."
            )
            return

        df_liquidacion = pd.DataFrame(data_liquidacion)
        df_res_base = pd.DataFrame(data_resumen)

        # Hoja Resumen: Agrupar por Título Obra único y sumar Totales
        df_resumen = df_res_base.groupby("Título Obra", as_index=False).agg({
            "Cant.": "sum",
            "Neto": "sum"
        }).rename(columns={
            "Cant.": "Total Cant.",
            "Neto": "Total Neto"
        })

        nombre_sugerido = f"Liquidacion_{os.path.splitext(os.path.basename(pdf_path))[0]}.xlsx"
        
        output_excel = filedialog.asksaveasfilename(
            title="¿Dónde querés guardar el archivo de Excel?",
            initialfile=nombre_sugerido,
            defaultextension=".xlsx",
            filetypes=[("Archivo de Excel", "*.xlsx")]
        )

        if not output_excel:
            return

        with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
            df_liquidacion.to_excel(writer, index=False, sheet_name="Liquidación")
            df_resumen.to_excel(writer, index=False, sheet_name="Resumen")

        messagebox.showinfo("¡Éxito!", f"El archivo Excel se generó correctamente en:\n{output_excel}")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al procesar el archivo:\n{str(e)}")

if __name__ == "__main__":
    procesar_pdf()

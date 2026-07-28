import os
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def procesar_pdf():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    pdf_path = filedialog.askopenfilename(
        title="Selecciona el archivo PDF de liquidación",
        filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")]
    )

    if not pdf_path:
        return

    data = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue
                    
                    header_idx = -1
                    col_indices = {}
                    
                    for i, row in enumerate(table):
                        row_str = " ".join([str(cell) for cell in row if cell is not None])
                        if "Titulo Obra" in row_str or "Título Obra" in row_str:
                            header_idx = i
                            for c_idx, cell in enumerate(row):
                                if cell:
                                    cell_clean = str(cell).replace("\n", " ").strip()
                                    if "Titulo" in cell_clean or "Título" in cell_clean:
                                        col_indices["titulo"] = c_idx
                                    elif "%" in cell_clean:
                                        col_indices["porcentaje"] = c_idx
                                    elif "Cant" in cell_clean:
                                        col_indices["cantidad"] = c_idx
                                    elif "Neto" in cell_clean:
                                        col_indices["neto"] = c_idx
                            break

                    if header_idx != -1:
                        for row in table[header_idx + 1:]:
                            if not row or all(cell is None or cell == "" for cell in row):
                                continue
                            
                            row_str = " ".join([str(c) for c in row if c])
                            if "Titulo Obra" in row_str or "Total" in row_str:
                                continue

                            def get_val(key):
                                idx = col_indices.get(key)
                                if idx is not None and idx < len(row) and row[idx] is not None:
                                    return str(row[idx]).strip()
                                return ""

                            data.append({
                                "Título Obra": get_val("titulo"),
                                "%": get_val("porcentaje"),
                                "Cant.": get_val("cantidad"),
                                "Neto": get_val("neto")
                            })

        if not data:
            messagebox.showwarning("Sin datos", "No se encontraron tablas compatibles en el PDF seleccionado.")
            return

        nombre_sugerido = f"Liquidacion_{os.path.splitext(os.path.basename(pdf_path))[0]}.xlsx"
        
        output_excel = filedialog.asksaveasfilename(
            title="¿Dónde querés guardar el archivo de Excel?",
            initialfile=nombre_sugerido,
            defaultextension=".xlsx",
            filetypes=[("Archivo de Excel", "*.xlsx")]
        )

        if not output_excel:
            return

        df = pd.DataFrame(data)
        with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Liquidación")

        messagebox.showinfo("¡Éxito!", f"El archivo se guardó correctamente en:\n{output_excel}")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al procesar el archivo:\n{str(e)}")

if __name__ == "__main__":
    procesar_pdf()

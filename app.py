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

    data = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                lines = text.split('\n')
                for line in lines:
                    line_clean = line.strip()

                    if not line_clean or "Título Obra" in line_clean or "Titulo Obra" in line_clean or "Concepto:" in line_clean:
                        continue

                    # Buscar Neto al final de la línea
                    neto_match = re.search(r'([\d\.,\-]+(?:\w+)?)$', line_clean)
                    # Buscar Porcentaje (%)
                    porcentaje_match = re.search(r'(\d{1,3}\.\d{2})', line_clean)

                    if neto_match and porcentaje_match:
                        neto_str = neto_match.group(1).strip()
                        porcentaje = porcentaje_match.group(1).strip()

                        # Extraer Título Obra
                        partes = re.split(r'\d{1,3}\.\d{2}', line_clean)
                        titulo_raw = partes[0] if partes else line_clean
                        titulo = re.sub(r'\s+[A-Z0-9]+$', '', titulo_raw).strip()

                        # Extraer Cantidad
                        cant_match = re.search(r'\s(\d{1,3})\s*-\s*', line_clean)
                        cantidad_str = cant_match.group(1) if cant_match else "1"

                        # Limpieza y conversión a valores numéricos para poder sumar en Excel
                        try:
                            # Quitar puntos de miles y reemplazar coma por punto si viniera formateado
                            neto_clean = re.sub(r'[^\d.-]', '', neto_str.replace('.', '').replace(',', '.'))
                            neto_val = float(neto_clean) if neto_clean else 0.0
                        except ValueError:
                            neto_val = 0.0

                        try:
                            cant_val = int(cantidad_str)
                        except ValueError:
                            cant_val = 1

                        data.append({
                            "Título Obra": titulo,
                            "%": porcentaje,
                            "Cant.": cant_val,
                            "Neto": neto_val
                        })

        if not data:
            messagebox.showwarning(
                "Aviso", 
                "No se pudieron extraer datos del PDF seleccionando esa estructura."
            )
            return

        # Crear DataFrames
        df_detalle = pd.DataFrame(data)

        # Crear Hoja de Resumen (Agrupado por Título Obra)
        df_resumen = df_detalle.groupby("Título Obra", as_index=False).agg({
            "Cant.": "sum",
            "Neto": "sum"
        }).rename(columns={
            "Cant.": "Total Cant.",
            "Neto": "Total Neto"
        })

        # Preguntar dónde guardar
        nombre_sugerido = f"Liquidacion_{os.path.splitext(os.path.basename(pdf_path))[0]}.xlsx"
        
        output_excel = filedialog.asksaveasfilename(
            title="¿Dónde querés guardar el archivo de Excel?",
            initialfile=nombre_sugerido,
            defaultextension=".xlsx",
            filetypes=[("Archivo de Excel", "*.xlsx")]
        )

        if not output_excel:
            return

        # Exportar ambas pestañas en el mismo archivo
        with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
            df_detalle.to_excel(writer, index=False, sheet_name="Liquidación")
            df_resumen.to_excel(writer, index=False, sheet_name="Resumen")

        messagebox.showinfo("¡Éxito!", f"El archivo Excel se generó correctamente en:\n{output_excel}")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al procesar el archivo:\n{str(e)}")

if __name__ == "__main__":
    procesar_pdf()

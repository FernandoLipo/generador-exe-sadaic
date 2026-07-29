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
        title="Selecciona el archivo PDF de liquidación SADAIC",
        filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")]
    )

    if not pdf_path:
        return

    data_liquidacion = []
    data_resumen = []

    # Palabras clave de autores/compositores y encabezados a ignorar para aislar el título
    palabras_corte = ["TAUZI", "RAMIREZ", "BARREIRO", "LAURIA", "CONCEPT", "TIPO DE DERECHO", "DISTRIBUCIÓN", "CAMBIO MONEDA"]

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                lines = text.split('\n')
                for line in lines:
                    line_clean = line.strip()

                    # Omitir encabezados generales o renglones de distribución/moneda
                    if not line_clean or "Título Obra" in line_clean or "Titulo Obra" in line_clean or "Concepto:" in line_clean:
                        continue
                    if "Distribución" in line_clean or "Cambio Moneda" in line_clean:
                        continue

                    # Buscar Neto y Porcentaje
                    neto_match = re.search(r'([\d\.,\-]+(?:\w+)?)$', line_clean)
                    porcentaje_match = re.search(r'(\d{1,3}\.\d{2})', line_clean)

                    if neto_match and porcentaje_match:
                        neto_original = neto_match.group(1).strip()
                        porcentaje_original = porcentaje_match.group(1).strip()

                        # Extraer la porción inicial antes del porcentaje
                        partes = re.split(r'\d{1,3}\.\d{2}', line_clean)
                        texto_previo = partes[0] if partes else line_clean

                        # Aislar el título limpio recortando cuando aparecen nombres de autores/códigos
                        titulo_limpio = texto_previo
                        for palabra in palabras_corte:
                            if palabra in titulo_limpio.upper():
                                titulo_limpio = titulo_limpio.upper().split(palabra)[0]

                        # Eliminar códigos numéricos sobrantes al final del título
                        titulo_limpio = re.sub(r'\s+\d+.*$', '', titulo_limpio).strip()

                        # Extraer Cantidad tal cual
                        cant_match = re.search(r'\s(\d{1,3})\s*-\s*', line_clean)
                        cantidad_original = cant_match.group(1) if cant_match else "1"

                        # 1. Datos intactos para la hoja "Liquidación"
                        data_liquidacion.append({
                            "Título Obra": titulo_limpio,
                            "%": porcentaje_original,
                            "Cant.": cantidad_original,
                            "Neto": neto_original
                        })

                        # 2. Conversión limpia para la hoja "Resumen"
                        try:
                            # Convertir neto a número flotante para sumar en el resumen
                            neto_num = float(neto_original.replace('.', '').replace(',', '.'))
                        except ValueError:
                            neto_num = 0.0

                        try:
                            cant_num = int(cantidad_original)
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
                "No se pudieron extraer datos válidos del PDF."
            )
            return

        # DataFrames
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

        # Diálogo para guardar
        nombre_sugerido = f"Liquidacion_{os.path.splitext(os.path.basename(pdf_path))[0]}.xlsx"
        
        output_excel = filedialog.asksaveasfilename(
            title="¿Dónde querés guardar el archivo de Excel?",
            initialfile=nombre_sugerido,
            defaultextension=".xlsx",
            filetypes=[("Archivo de Excel", "*.xlsx")]
        )

        if not output_excel:
            return

        # Exportar ambas pestañas
        with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
            df_liquidacion.to_excel(writer, index=False, sheet_name="Liquidación")
            df_resumen.to_excel(writer, index=False, sheet_name="Resumen")

        messagebox.showinfo("¡Éxito!", f"El archivo Excel se generó correctamente en:\n{output_excel}")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error al procesar el archivo:\n{str(e)}")

if __name__ == "__main__":
    procesar_pdf()

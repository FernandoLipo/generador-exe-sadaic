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
                text = page.extract_text()
                if not text:
                    continue

                lines = text.split('\n')
                for line in lines:
                    line_clean = line.strip()

                    # Omitir líneas vacías, encabezados, subtotales o líneas irrelevantes
                    if not line_clean or "Título Obra" in line_clean or "Titulo Obra" in line_clean or "Concepto:" in line_clean:
                        continue
                    if "Distribución" in line_clean or "Cambio Moneda" in line_clean or "Total" in line_clean or "LIQ." in line_clean:
                        continue

                    # 1. VERIFICAR SI TIENE PORCENTAJE DE PARTICIPACIÓN (%)
                    pct_match = re.search(r'\b(\d{1,3}\.\d{2})\b', line_clean)
                    if not pct_match:
                        continue

                    pct_val = pct_match.group(1).strip()

                    # Separar parte izquierda (Título + Autor/es) y parte derecha (Metadatos + Importes)
                    partes_pct = line_clean.split(pct_val, 1)
                    texto_izq = partes_pct[0].strip()
                    texto_der = partes_pct[1].strip() if len(partes_pct) > 1 else ""

                    # 2. EXTRAER TÍTULO DE MANERA PRECISA
                    # Buscamos la presencia del código de obra o indicador de tipo/rol ' E ' al final del bloque de autor
                    # Los nombres de autor en SADAIC suelen estar en MAYÚSCULAS tras el título de la obra.
                    
                    # Intentamos separar el título cortando donde empieza la secuencia del primer Autor
                    # Formato típico: TITULO DE LA OBRA <APELLIDO NOMBRES> E <COD_OBRA>
                    match_autor_e = re.search(r'\s+([A-ZÁÉÍÓÚÑ\s]+?)\s+E\b', texto_izq)
                    
                    if match_autor_e:
                        # Si encontramos el bloque de autor antes de la 'E', tomamos todo lo que está antes como título
                        titulo_limpio = texto_izq[:match_autor_e.start()].strip()
                    else:
                        # Si la 'E' no está en la parte izquierda, cortamos por el primer salto de doble espacio
                        partes_espacio = re.split(r'\s{2,}', texto_izq)
                        titulo_limpio = partes_espacio[0].strip()

                    # Normalización de espacios múltiples en el título
                    titulo_limpio = re.sub(r'\s+', ' ', titulo_limpio).strip()

                    if not titulo_limpio or any(k in titulo_limpio for k in ["Distribución", "Cambio", "Total Concepto"]):
                        continue

                    # 3. EXTRAER CANTIDAD
                    cant_match = re.search(r'\b(\d{1,3})\s*-\s*', texto_der) or re.search(r'\bAR\s+(\d{1,3})\b', texto_der)
                    cant_val = cant_match.group(1).strip() if cant_match else "1"

                    # 4. EXTRAER NETO
                    neto_match = re.search(r'([\d\.,\-]+)$', texto_der)
                    if neto_match:
                        neto_val = neto_match.group(1).strip()
                    else:
                        neto_alt = re.findall(r'[\d\.,\-]+', texto_der)
                        neto_val = neto_alt[-1] if neto_alt else "0.00"

                    # Formatear montos y cantidades numéricas
                    try:
                        neto_clean = neto_val.replace('.', '').replace(',', '.')
                        neto_num = float(neto_clean)
                    except ValueError:
                        neto_num = 0.0

                    try:
                        cant_num = int(cant_val)
                    except ValueError:
                        cant_num = 1

                    # Agregar registro detallado a Liquidación
                    data_liquidacion.append({
                        "Título Obra": titulo_limpio,
                        "%": pct_val,
                        "Cant.": cant_num,
                        "Neto": neto_val
                    })

                    # Agregar registro acumulativo a Resumen
                    data_resumen.append({
                        "Título Obra": titulo_limpio,
                        "Cant.": cant_num,
                        "Neto": neto_num
                    })

        if not data_liquidacion:
            messagebox.showwarning("Aviso", "No se pudieron extraer los datos del PDF.")
            return

        df_liquidacion = pd.DataFrame(data_liquidacion)
        df_res_base = pd.DataFrame(data_resumen)

        # Agrupar por Título Obra único y sumar Totales en la hoja de Resumen
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

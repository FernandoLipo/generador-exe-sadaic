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

    # Apellidos/Nombres conocidos que marcan el FIN del título de la obra
    patron_corte_autores = r'\b(TAUZI|RAMIREZ|BARREIRO|LAURIA|URBANI|LOPEZ|GARCIA|GOMEZ|PEREZ|RODRIGUEZ|GONZALEZ|FERNANDEZ|MARTINEZ|SANCHEZ|ROMERO|SOUTO|ALVAREZ)\b'

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                lines = text.split('\n')
                for line in lines:
                    line_clean = line.strip()

                    # Omitir encabezaos y subtotales
                    if not line_clean or "Título Obra" in line_clean or "Titulo Obra" in line_clean or "Concepto:" in line_clean:
                        continue
                    if "Distribución" in line_clean or "Cambio Moneda" in line_clean or "Total" in line_clean:
                        continue

                    # 1. BUSCAR PORCENTAJE (%) - ej: 12.50
                    pct_match = re.search(r'(\d{1,3}\.\d{2})', line_clean)
                    if not pct_match:
                        continue

                    pct_val = pct_match.group(1).strip()

                    # 2. SEPARAR TEXTO ANTES Y DESPUÉS DEL PORCENTAJE
                    partes_pct = line_clean.split(pct_val, 1)
                    texto_izq = partes_pct[0].strip()   # Contiene Título + Autor
                    texto_der = partes_pct[1].strip() if len(partes_pct) > 1 else "" # Contiene Fechas, Local, Cant, Neto

                    # 3. EXTRAER TÍTULO OBRA (De la parte izquierda, cortando antes del nombre del autor)
                    corte_match = re.search(patron_corte_autores, texto_izq, re.IGNORECASE)
                    if corte_match:
                        titulo_limpio = texto_izq[:corte_match.start()].strip()
                    else:
                        # Si no coincide con la lista, corta donde encuentre dos palabras seguidas en Mayúsculas/Cód
                        titulo_limpio = re.split(r'\s+[A-ZÁÉÍÓÚÑ]{3,}\s+[A-ZÁÉÍÓÚÑ]{3,}', texto_izq)[0].strip()

                    # Eliminar códigos numéricos sueltos si quedaron al final del título
                    titulo_limpio = re.sub(r'\s+\d+$', '', titulo_limpio).strip()

                    if not titulo_limpio:
                        continue

                    # 4. EXTRAER CANTIDAD (De la parte derecha)
                    cant_match = re.search(r'\b(\d{1,3})\s*-\s*', texto_der)
                    if not cant_match:
                        cant_match = re.search(r'\bAR\s+(\d{1,2})\b', texto_der)
                    
                    cant_val = cant_match.group(1).strip() if cant_match else "1"

                    # 5. EXTRAER NETO (De la parte derecha, aislando números y decimales antes de 'INCIDENTAL' u otros textos)
                    neto_match = re.search(r'([\d\.,\-]+)(?:\s*[A-Z]+)?$', texto_der)
                    if neto_match:
                        neto_val = neto_match.group(1).strip()
                    else:
                        # Buscar cualquier número decimal al final
                        neto_alt = re.findall(r'[\d\.,\-]+', texto_der)
                        neto_val = neto_alt[-1] if neto_alt else "0.00"

                    # 1. Pestaña "Liquidación" (Mantener formato tal cual)
                    data_liquidacion.append({
                        "Título Obra": titulo_limpio,
                        "%": pct_val,
                        "Cant.": cant_val,
                        "Neto": neto_val
                    })

                    # 2. Pestaña "Resumen" (Convertir a número para sumar)
                    try:
                        neto_clean = re.sub(r'[^\d.-]', '', neto_val.replace('.', '').replace(',', '.'))
                        neto_num = float(neto_clean) if neto_clean else 0.0
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
                "No se pudieron extraer los datos del PDF."
            )
            return

        df_liquidacion = pd.DataFrame(data_liquidacion)
        df_res_base = pd.DataFrame(data_resumen)

        # Agrupar por Título Obra único y sumar Totales
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

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

    # Lista de autores/socios frecuentes para usar como límite estricto de corte
    patron_autores = r'\b(TAUZI|RAMIREZ|BARREIRO|LAURIA|URBANI|LOPEZ|GARCIA|GOMEZ|PEREZ|RODRIGUEZ|GONZALEZ|FERNANDEZ|MARTINEZ|SANCHEZ|ROMERO|SOUTO|ALVAREZ)\b'

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                lines = text.split('\n')
                i = 0
                while i < len(lines):
                    line_clean = lines[i].strip()

                    # Omitir encabezados, subtotales o líneas irrelevantes
                    if not line_clean or "Título Obra" in line_clean or "Titulo Obra" in line_clean or "Concepto:" in line_clean:
                        i += 1
                        continue
                    if "Distribución" in line_clean or "Cambio Moneda" in line_clean or "Total" in line_clean:
                        i += 1
                        continue

                    # 1. VERIFICAR SI LA LÍNEA TIENE PORCENTAJE (%) Y MONTO
                    pct_match = re.search(r'(\d{1,3}\.\d{2})', line_clean)
                    if not pct_match:
                        i += 1
                        continue

                    pct_val = pct_match.group(1).strip()

                    # Separar texto a la izquierda del porcentaje (Título + Autor) y a la derecha (Cant + Neto)
                    partes_pct = line_clean.split(pct_val, 1)
                    texto_izq = partes_pct[0].strip()
                    texto_der = partes_pct[1].strip() if len(partes_pct) > 1 else ""

                    # 2. EVALUAR SI EL TÍTULO CONTINÚA EN EL SIGUIENTE RENGLÓN
                    # Miramos la línea siguiente si existe
                    if i + 1 < len(lines):
                        siguiente_linea = lines[i + 1].strip()
                        # Si la siguiente línea NO tiene porcentaje % ni es encabezado, puede ser la continuación del título
                        if siguiente_linea and not re.search(r'(\d{1,3}\.\d{2})', siguiente_linea) and not "Concepto:" in siguiente_linea:
                            # Cortamos autores si los hubiera en la segunda línea
                            corte_sig = re.search(patron_autores, siguiente_linea, re.IGNORECASE)
                            sub_titulo = siguiente_linea[:corte_sig.start()].strip() if corte_sig else siguiente_linea
                            
                            # Si hay texto válido en el renglón de abajo, lo anexamos al título
                            if len(sub_titulo) > 0:
                                texto_izq = f"{texto_izq} {sub_titulo}"

                    # 3. AISLAR EL TÍTULO LIMPIO (Cortar antes del nombre de los autores/socios)
                    corte_match = re.search(patron_autores, texto_izq, re.IGNORECASE)
                    if corte_match:
                        titulo_limpio = texto_izq[:corte_match.start()].strip()
                    else:
                        # Si no hay autor en la lista, corta donde encuentre dos palabras seguidas en Mayúsculas o código
                        titulo_limpio = re.split(r'\s+[A-ZÁÉÍÓÚÑ]{3,}\s+[A-ZÁÉÍÓÚÑ]{3,}', texto_izq)[0].strip()

                    # Limpiar códigos numéricos finales si quedaron sueltos
                    titulo_limpio = re.sub(r'\s+\d+$', '', titulo_limpio).strip()

                    if not titulo_limpio:
                        i += 1
                        continue

                    # 4. EXTRAER CANTIDAD
                    cant_match = re.search(r'\b(\d{1,3})\s*-\s*', texto_der)
                    if not cant_match:
                        cant_match = re.search(r'\bAR\s+(\d{1,2})\b', texto_der)
                    cant_val = cant_match.group(1).strip() if cant_match else "1"

                    # 5. EXTRAER NETO
                    neto_match = re.search(r'([\d\.,\-]+)(?:\s*[A-Z]+)?$', texto_der)
                    if neto_match:
                        neto_val = neto_match.group(1).strip()
                    else:
                        neto_alt = re.findall(r'[\d\.,\-]+', texto_der)
                        neto_val = neto_alt[-1] if neto_alt else "0.00"

                    # 1. Pestaña Liquidación (Tal cual en el PDF)
                    data_liquidacion.append({
                        "Título Obra": titulo_limpio,
                        "%": pct_val,
                        "Cant.": cant_val,
                        "Neto": neto_val
                    })

                    # 2. Pestaña Resumen (Conversión a número para sumar)
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

                    i += 1

        if not data_liquidacion:
            messagebox.showwarning(
                "Aviso", 
                "No se pudieron extraer datos válidos del PDF."
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

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

                # Extraer las palabras de la página con sus coordenadas espaciales
                words = page.extract_words()
                if not words:
                    continue

                # Encontrar el encabezado "Título" / "Nombre" para determinar la frontera X exacta entre columnas
                x_corte_nombre = 220.0  # Límite por defecto para el título en SADAIC
                for w in words:
                    if "Nombre" in w['text']:
                        x_corte_nombre = w['x0'] - 5  # La columna Nombre empieza aquí
                        break

                lines = text.split('\n')
                for idx, line in enumerate(lines):
                    line_clean = line.strip()

                    # Ignorar encabezados o subtotales
                    if not line_clean or "Título Obra" in line_clean or "Concepto:" in line_clean or "Total" in line_clean:
                        continue
                    if "Distribución" in line_clean or "Cambio Moneda" in line_clean:
                        continue

                    # Verificar si la línea contiene un Porcentaje (%) de liquidación
                    pct_match = re.search(r'(\d{1,3}\.\d{2})', line_clean)
                    if not pct_match:
                        continue

                    pct_val = pct_match.group(1).strip()
                    
                    # Separar la parte derecha para Cantidad y Neto
                    partes = line_clean.split(pct_val, 1)
                    texto_der = partes[1].strip() if len(partes) > 1 else ""

                    # EXTRAER TÍTULO POR COORDENADAS ESPACIALES EN EL PDF
                    # Buscamos todas las palabras en esa zona de la página que estén a la izquierda de 'Nombre'
                    # y alineadas verticalmente con esta fila
                    
                    # 1. Palabras en la primera línea del título
                    palabras_linea1 = [
                        w['text'] for w in words 
                        if w['x1'] <= x_corte_nombre and abs(w['top'] - line_y_approx(words, line_clean)) < 8
                    ]
                    
                    # Fallback si no encuentra por coordenadas exactas
                    if palabras_linea1:
                        tit_part1 = " ".join(palabras_linea1).strip()
                    else:
                        tit_part1 = partes[0].strip()

                    # 2. Verificar si hay un segundo renglón del título justo debajo
                    tit_part2 = ""
                    if idx + 1 < len(lines):
                        siguiente_linea = lines[idx + 1].strip()
                        # Si el renglón de abajo NO tiene porcentaje % ni es un concepto/encabezado
                        if siguiente_linea and not re.search(r'(\d{1,3}\.\d{2})', siguiente_linea) and not "Concepto:" in siguiente_linea:
                            # Filtramos las palabras del renglón de abajo que estén DENTRO de la columna Título
                            palabras_linea2 = [
                                w['text'] for w in words 
                                if w['x1'] <= x_corte_nombre and abs(w['top'] - line_y_approx(words, siguiente_linea)) < 8
                            ]
                            if palabras_linea2:
                                tit_part2 = " ".join(palabras_linea2).strip()

                    # Construir el Título Obra Unificado y Completo
                    if tit_part2:
                        titulo_completo = f"{tit_part1} {tit_part2}".strip()
                    else:
                        titulo_completo = tit_part1.strip()

                    # Limpieza final por si quedó alguna palabra duplicada o basura
                    titulo_limpio = " ".join(titulo_completo.split()).strip()

                    if not titulo_limpio:
                        continue

                    # EXTRAER CANTIDAD Y NETO
                    cant_match = re.search(r'\b(\d{1,3})\s*-\s*', texto_der)
                    if not cant_match:
                        cant_match = re.search(r'\bAR\s+(\d{1,2})\b', texto_der)
                    cant_val = cant_match.group(1).strip() if cant_match else "1"

                    neto_match = re.search(r'([\d\.,\-]+)(?:\s*[A-Z]+)?$', texto_der)
                    if neto_match:
                        neto_val = neto_match.group(1).strip()
                    else:
                        neto_alt = re.findall(r'[\d\.,\-]+', texto_der)
                        neto_val = neto_alt[-1] if neto_alt else "0.00"

                    # Agregar a Pestaña Liquidación
                    data_liquidacion.append({
                        "Título Obra": titulo_limpio,
                        "%": pct_val,
                        "Cant.": cant_val,
                        "Neto": neto_val
                    })

                    # Agregar a Pestaña Resumen
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

def line_y_approx(words, line_text):
    """Encuentra la posición Y aproximada de una línea de texto"""
    first_word = line_text.split()[0] if line_text.split() else ""
    for w in words:
        if w['text'] == first_word:
            return w['top']
    return 0.0

if __name__ == "__main__":
    procesar_pdf()

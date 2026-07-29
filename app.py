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
                words = page.extract_words()
                if not words:
                    continue

                # 1. DETERMINAR LÍMITES HORIZONTALES (Columnas)
                # Buscamos las coordenadas X de "Título Obra" y "Nombre"
                x_inicio_titulo = 0.0
                x_fin_titulo = 220.0  # Límite por defecto si no encuentra la columna Nombre

                for w in words:
                    if "Título" in w['text'] or "Titulo" in w['text']:
                        x_inicio_titulo = max(0.0, w['x0'] - 5)
                    if "Nombre" in w['text']:
                        x_fin_titulo = w['x0'] - 5  # La columna Nombre marca el límite derecho estricto
                        break

                # 2. ENCONTRAR LÍNEAS HORIZONTALES (Límites de filas/celdas)
                # Extraemos las líneas horizontales dibujadas en el PDF
                lines = page.lines
                y_divisores = sorted(list(set([round(l['top'], 1) for l in lines if l['width'] > 100])))

                # Si el PDF no tiene objetos línea explícitos, agrupamos por renglones de texto
                if len(y_divisores) < 2:
                    text = page.extract_text()
                    if not text:
                        continue
                    lineas_texto = text.split('\n')
                    for line in lineas_texto:
                        line_clean = line.strip()
                        if not line_clean or "Título Obra" in line_clean or "Concepto:" in line_clean or "Total" in line_clean:
                            continue
                        
                        pct_match = re.search(r'(\d{1,3}\.\d{2})', line_clean)
                        if not pct_match:
                            continue

                        pct_val = pct_match.group(1).strip()
                        partes = line_clean.split(pct_val, 1)
                        texto_izq = partes[0].strip()
                        texto_der = partes[1].strip() if len(partes) > 1 else ""

                        # Limpieza por palabras clave si falla el detector geométrico
                        tit_limpio = re.split(r'\s+[A-ZÁÉÍÓÚÑ]{3,}\s+[A-ZÁÉÍÓÚÑ]{3,}', texto_izq)[0].strip()
                        
                        cant_match = re.search(r'\b(\d{1,3})\s*-\s*', texto_der) or re.search(r'\bAR\s+(\d{1,2})\b', texto_der)
                        cant_val = cant_match.group(1).strip() if cant_match else "1"
                        
                        neto_match = re.search(r'([\d\.,\-]+)(?:\s*[A-Z]+)?$', texto_der)
                        neto_val = neto_match.group(1).strip() if neto_match else "0.00"

                        data_liquidacion.append({"Título Obra": tit_limpio, "%": pct_val, "Cant.": cant_val, "Neto": neto_val})
                        agregar_a_resumen(data_resumen, tit_limpio, cant_val, neto_val)
                    continue

                # 3. EXTRAER CELDA POR CELDA ENTRE LÍNEAS DIVISORIAS
                for idx in range(len(y_divisores) - 1):
                    y_top = y_divisores[idx]
                    y_bottom = y_divisores[idx + 1]

                    # Distancia mínima razonable de una fila de tabla
                    if (y_bottom - y_top) < 8:
                        continue

                    # Bajar la caja de palabras estrictamente dentro del recuadro
                    palabras_celda_titulo = [
                        w for w in words 
                        if w['x0'] >= x_inicio_titulo 
                        and w['x1'] <= x_fin_titulo 
                        and w['top'] >= y_top - 2 
                        and w['bottom'] <= y_bottom + 2
                    ]

                    # Si no hay título en este recuadro, pasar al siguiente
                    if not palabras_celda_titulo:
                        continue

                    # Ordenar palabras de arriba a abajo y de izquierda a derecha
                    palabras_celda_titulo.sort(key=lambda w: (round(w['top'], 1), w['x0']))
                    titulo_obra = " ".join([w['text'] for w in palabras_celda_titulo]).strip()

                    # Omitir si agarró la cabecera
                    if "Título" in titulo_obra or "Obra" in titulo_obra:
                        continue

                    # Buscar Porcentaje, Cantidad y Neto en la misma franja horizontal (y_top a y_bottom)
                    palabras_fila_completa = [
                        w for w in words 
                        if w['top'] >= y_top - 2 and w['bottom'] <= y_bottom + 2
                    ]
                    palabras_fila_completa.sort(key=lambda w: w['x0'])
                    texto_fila = " ".join([w['text'] for w in palabras_fila_completa])

                    pct_match = re.search(r'(\d{1,3}\.\d{2})', texto_fila)
                    if not pct_match:
                        continue

                    pct_val = pct_match.group(1).strip()

                    # Separar texto a la derecha del Porcentaje
                    partes = texto_fila.split(pct_val, 1)
                    texto_der = partes[1].strip() if len(partes) > 1 else ""

                    cant_match = re.search(r'\b(\d{1,3})\s*-\s*', texto_der) or re.search(r'\bAR\s+(\d{1,2})\b', texto_der)
                    cant_val = cant_match.group(1).strip() if cant_match else "1"

                    neto_match = re.search(r'([\d\.,\-]+)(?:\s*[A-Z]+)?$', texto_der)
                    if neto_match:
                        neto_val = neto_match.group(1).strip()
                    else:
                        neto_alt = re.findall(r'[\d\.,\-]+', texto_der)
                        neto_val = neto_alt[-1] if neto_alt else "0.00"

                    data_liquidacion.append({
                        "Título Obra": titulo_obra,
                        "%": pct_val,
                        "Cant.": cant_val,
                        "Neto": neto_val
                    })

                    agregar_a_resumen(data_resumen, titulo_obra, cant_val, neto_val)

        if not data_liquidacion:
            messagebox.showwarning("Aviso", "No se pudieron extraer los datos del PDF.")
            return

        df_liquidacion = pd.DataFrame(data_liquidacion)
        df_res_base = pd.DataFrame(data_resumen)

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

def agregar_a_resumen(data_resumen, titulo, cant_val, neto_val):
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
        "Título Obra": titulo,
        "Cant.": cant_num,
        "Neto": neto_num
    })

if __name__ == "__main__":
    procesar_pdf()

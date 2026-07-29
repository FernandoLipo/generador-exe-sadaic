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
                i = 0
                while i < len(lines):
                    line_clean = lines[i].strip()

                    # Omitir encabezados, subtotales o líneas irrelevantes del PDF
                    if not line_clean or "Título Obra" in line_clean or "Titulo Obra" in line_clean or "Concepto:" in line_clean:
                        i += 1
                        continue
                    if "Distribución" in line_clean or "Cambio Moneda" in line_clean or "Total" in line_clean or "LIQ." in line_clean:
                        i += 1
                        continue

                    # 1. VERIFICAR SI TIENE PORCENTAJE (%)
                    pct_match = re.search(r'(\d{1,3}\.\d{2})', line_clean)
                    if not pct_match:
                        i += 1
                        continue

                    pct_val = pct_match.group(1).strip()

                    # 2. SEPARAR PARTE IZQUIERDA (Título + Autor) Y PARTE DERECHA (Cant + Neto)
                    partes_pct = line_clean.split(pct_val, 1)
                    texto_izq = partes_pct[0].strip()
                    texto_der = partes_pct[1].strip() if len(partes_pct) > 1 else ""

                    # 3. VERIFICAR SI EL TÍTULO CONTINÚA EN EL RENGLÓN DE ABAJO
                    sub_titulo = ""
                    if i + 1 < len(lines):
                        siguiente_linea = lines[i + 1].strip()
                        # Si el renglón de abajo NO tiene porcentaje %, ni es encabezado o total
                        if siguiente_linea and not re.search(r'(\d{1,3}\.\d{2})', siguiente_linea) and not any(k in siguiente_linea for k in ["Concepto:", "Distribución", "Cambio Moneda", "Total"]):
                            # Extraemos la parte del título del renglón de abajo cortando si aparecen autores
                            sub_titulo = re.split(r'\s+[A-ZÁÉÍÓÚÑ]{3,}\s+[A-ZÁÉÍÓÚÑ]{3,}', siguiente_linea)[0].strip()

                    # 4. EXTRAER TÍTULO DE LA PRIMERA LÍNEA CORTANDO ÚNICAMENTE DONDE EMPIEZAN LOS NOMBRES DE AUTORES
                    # Se busca el patrón donde termina el título (palabra en MAYÚSCULAS/Números) y empiezan los Apellidos
                    corte_autor = re.search(r'\s+([A-ZÁÉÍÓÚÑ]{3,}\s+[A-ZÁÉÍÓÚÑ]{3,}\s+[A-ZÁÉÍÓÚÑ]{3,}\s+E\s+\d+)', texto_izq)
                    
                    if corte_autor:
                        tit_principal = texto_izq[:corte_autor.start()].strip()
                    else:
                        # Si no coincide la estructura exacta de socio, separar por doble espacio o secuencia de autor
                        partes_espacio = re.split(r'\s{2,}', texto_izq)
                        tit_principal = partes_espacio[0].strip()

                    # Unir línea 1 y línea 2 si correspondía
                    if sub_titulo and sub_titulo not in tit_principal:
                        titulo_completo = f"{tit_principal} {sub_titulo}".strip()
                    else:
                        titulo_completo = tit_principal.strip()

                    # Limpieza general
                    titulo_limpio = re.sub(r'\s+', ' ', titulo_completo).strip()

                    if not titulo_limpio or "Distribución" in titulo_limpio or "Cambio" in titulo_limpio:
                        i += 1
                        continue

                    # 5. EXTRAER CANTIDAD
                    cant_match = re.search(r'\b(\d{1,3})\s*-\s*', texto_der) or re.search(r'\bAR\s+(\d{1,2})\b', texto_der)
                    cant_val = cant_match.group(1).strip() if cant_match else "1"

                    # 6. EXTRAER NETO
                    neto_match = re.search(r'([\d\.,\-]+)(?:\s*[A-Z]+)?$', texto_der)
                    if neto_match:
                        neto_val = neto_match.group(1).strip()
                    else:
                        neto_alt = re.findall(r'[\d\.,\-]+', texto_der)
                        neto_val = neto_alt[-1] if neto_alt else "0.00"

                    # Hoja Liquidación
                    data_liquidacion.append({
                        "Título Obra": titulo_limpio,
                        "%": pct_val,
                        "Cant.": cant_val,
                        "Neto": neto_val
                    })

                    # Hoja Resumen
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
            messagebox.showwarning("Aviso", "No se pudieron extraer los datos del PDF.")
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

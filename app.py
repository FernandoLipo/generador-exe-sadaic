import os
import re
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def normalizar_texto(texto):
    if not isinstance(texto, str):
        return ""
    texto = re.sub(r'\s+', ' ', texto)
    return texto.strip().upper()

def procesar_liquidacion_con_guia():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    # 1. PEDIR EL ARCHIVO EXCEL GUÍA
    messagebox.showinfo("Paso 1", "Seleccioná el archivo Excel que contiene la lista GUÍA de temas (Titulo Obra.xlsx).")
    excel_guia_path = filedialog.askopenfilename(
        title="Seleccioná el Excel GUÍA de temas",
        filetypes=[("Archivos de Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
    )

    if not excel_guia_path:
        return

    # Leer los títulos del Excel guía (sin asumir que la 1ra fila es encabezado)
    try:
        df_guia = pd.read_excel(excel_guia_path, header=None)
        
        # Juntar todas las celdas texto del Excel en una sola lista
        titulos_brutos = []
        for col in df_guia.columns:
            titulos_brutos.extend(df_guia[col].dropna().astype(str).tolist())

        # Limpiar y ordenar de MAYOR a MENOR longitud
        lista_titulos = sorted(
            [normalizar_texto(t) for t in titulos_brutos if len(t.strip()) > 0],
            key=len,
            reverse=True
        )
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo leer el Excel guía:\n{str(e)}")
        return

    # 2. PEDIR EL ARCHIVO PDF DE LIQUIDACIÓN
    messagebox.showinfo("Paso 2", "Seleccioná el archivo PDF de Liquidación SADAIC.")
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

                    # Omitir encabezados y textos institucionales
                    if not line_clean or "Título Obra" in line_clean or "Concepto:" in line_clean:
                        continue
                    if "Distribución" in line_clean or "Cambio Moneda" in line_clean or "Total" in line_clean or "LIQ." in line_clean:
                        continue

                    # Verificar si la línea contiene un porcentaje de participación (%)
                    pct_match = re.search(r'\b(\d{1,3}\.\d{2})\b', line_clean)
                    if not pct_match:
                        continue

                    pct_val = pct_match.group(1).strip()
                    partes_pct = line_clean.split(pct_val, 1)
                    texto_izq = normalizar_texto(partes_pct[0])
                    texto_der = partes_pct[1].strip() if len(partes_pct) > 1 else ""

                    # 3. BUSCAR COINCIDENCIA CON LA LISTA GUÍA
                    titulo_encontrado = None
                    for t_guia in lista_titulos:
                        if t_guia in texto_izq:
                            titulo_encontrado = t_guia
                            break

                    # 4. SI NO ESTÁ EN LA GUÍA, EXTRAER SIN CORTAR PALABRAS DEL TÍTULO
                    if not titulo_encontrado:
                        # SADAIC pone: TITULO AUTOR E CODIGO_OBRA
                        # Buscamos el patrón " E " seguido de números de obra
                        match_e = re.search(r'\s+E\s+\d{5,8}', texto_izq)
                        if match_e:
                            # Tomamos todo antes de " E CODIGO" y quitamos las palabras del autor (que suelen ser las últimas 2 o 3)
                            bloque_titulo_autor = texto_izq[:match_e.start()].strip()
                            partes = bloque_titulo_autor.split()
                            # Si tiene varias palabras, tomamos la mayoría excepto las probables de autor
                            titulo_encontrado = bloque_titulo_autor
                        else:
                            titulo_encontrado = texto_izq

                    titulo_limpio = normalizar_texto(titulo_encontrado)

                    # 5. EXTRAER CANTIDAD
                    cant_match = re.search(r'\b(\d{1,4})\s*-\s*', texto_der) or re.search(r'\bAR\s+(\d{1,4})\b', texto_der)
                    cant_val = cant_match.group(1).strip() if cant_match else "1"

                    # 6. EXTRAER NETO
                    neto_match = re.search(r'([\d\.,\-]+)$', texto_der)
                    if neto_match:
                        neto_val = neto_match.group(1).strip()
                    else:
                        neto_alt = re.findall(r'[\d\.,\-]+', texto_der)
                        neto_val = neto_alt[-1] if neto_alt else "0.00"

                    # Formatear números
                    try:
                        neto_clean = neto_val.replace('.', '').replace(',', '.')
                        neto_num = float(neto_clean)
                    except ValueError:
                        neto_num = 0.0

                    try:
                        cant_num = int(cant_val)
                    except ValueError:
                        cant_num = 1

                    data_liquidacion.append({
                        "Título Obra": titulo_limpio,
                        "%": pct_val,
                        "Cant.": cant_num,
                        "Neto": neto_val
                    })

                    data_resumen.append({
                        "Título Obra": titulo_limpio,
                        "Cant.": cant_num,
                        "Neto": neto_num
                    })

        if not data_liquidacion:
            messagebox.showwarning("Aviso", "No se pudieron extraer datos del PDF.")
            return

        df_liquidacion = pd.DataFrame(data_liquidacion)
        df_res_base = pd.DataFrame(data_resumen)

        # Agrupar resumen por obra
        df_resumen = df_res_base.groupby("Título Obra", as_index=False).agg({
            "Cant.": "sum",
            "Neto": "sum"
        }).rename(columns={
            "Cant.": "Total Cant.",
            "Neto": "Total Neto"
        })

        # GUARDAR RESULTADO
        nombre_sugerido = f"Liquidacion_{os.path.splitext(os.path.basename(pdf_path))[0]}.xlsx"
        
        output_excel = filedialog.asksaveasfilename(
            title="¿Dónde querés guardar el archivo Excel resultante?",
            initialfile=nombre_sugerido,
            defaultextension=".xlsx",
            filetypes=[("Archivo de Excel", "*.xlsx")]
        )

        if not output_excel:
            return

        with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
            df_liquidacion.to_excel(writer, index=False, sheet_name="Liquidación")
            df_resumen.to_excel(writer, index=False, sheet_name="Resumen")

        messagebox.showinfo("¡Éxito!", f"El proceso finalizó correctamente.\nGuardado en:\n{output_excel}")

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error inesperado:\n{str(e)}")

if __name__ == "__main__":
    procesar_liquidacion_con_guia()

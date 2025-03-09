import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# ------------------- PARTE 1: PROCESAMIENTO Y GENERACIÓN DE "Certificado.xlsx" -------------------
st.title("Validador de registro de datos - densidad")

# Selección de plantilla
opciones_plantilla = {
    "ROSA LA PRIMOROSA": "PLANTILLA.xlsx",
    "MILAGROS CHAMPIRREINO": "PLANTILLA1.xlsx",
    "YONATAN CON Y": "PLANTILLA2.xlsx"
}

plantilla_seleccionada = st.selectbox("Seleccione el responsable:", list(opciones_plantilla.keys()))
plantilla_path = opciones_plantilla[plantilla_seleccionada]

# Subir archivo
archivo_cargado = st.file_uploader("Carga archivo de datos en Excel", type=["xlsx"])

def procesar_archivo(archivo_cargado, plantilla):
    plantilla_wb = openpyxl.load_workbook(plantilla)
    plantilla_ws = plantilla_wb["PECLD07792"]
    duplicado_ws = plantilla_wb["Duplicado"]
    standar_ws = plantilla_wb["STD (PECLSTDEN02)"]

    wb = openpyxl.load_workbook(archivo_cargado, data_only=True)
    if "BD_densidad_2020" not in wb.sheetnames:
        st.error("El archivo cargado no contiene la hoja 'BD_densidad_2020'")
        return None
    ws = wb["BD_densidad_2020"]

    datos = []
    for row in ws.iter_rows(min_row=10, max_col=17, values_only=True):
        if any(row):
            datos.append(row)

    start_row = 28
    for i, row in enumerate(datos, start=start_row):
        for j, value in enumerate(row, start=1):
            plantilla_ws.cell(row=i, column=j, value=value)

    max_row = plantilla_ws.max_row
    columns = 17
    filas_eliminar = []
    for i in range(start_row + len(datos), max_row + 1):
        if all(plantilla_ws.cell(row=i, column=j).value is None for j in range(1, columns + 1)):
            filas_eliminar.append(i)

    for row in reversed(filas_eliminar):
        plantilla_ws.delete_rows(row)

    from openpyxl.styles import PatternFill
    fill = PatternFill(start_color="E26B0A", end_color="E26B0A", fill_type="solid")
    
    for row in plantilla_ws.iter_rows(min_row=28, max_col=17):
        if any(cell.value in ["DSTD", "DEND"] for cell in row):
            for cell in row:
                cell.fill = fill

    dest_row = 11
    for row in plantilla_ws.iter_rows(min_row=27, min_col=15, max_col=15, values_only=False):
        if row[0].value == "DEND":
            fila_actual = row[0].row
            valor_d = plantilla_ws.cell(row=fila_actual, column=4).value
            valor_m = plantilla_ws.cell(row=fila_actual, column=13).value
            valor_m_ant = plantilla_ws.cell(row=fila_actual - 1, column=13).value
            
            duplicado_ws.cell(row=dest_row, column=1, value=valor_d)
            duplicado_ws.cell(row=dest_row, column=3, value=valor_d)
            duplicado_ws.cell(row=dest_row, column=4, value=valor_m)
            duplicado_ws.cell(row=dest_row, column=2, value=valor_m_ant)
            dest_row += 1

    nuevo_nombre = plantilla_ws.cell(row=28, column=1).value
    if nuevo_nombre:
        plantilla_ws.title = str(nuevo_nombre)

    output = BytesIO()
    plantilla_wb.save(output)
    output.seek(0)
    return output

if archivo_cargado is not None:
    output = procesar_archivo(archivo_cargado, plantilla_path)
    if output:
        st.download_button(
            label="Descargar archivo procesado",
            data=output,
            file_name="Certificado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ------------------- PARTE 2: CARGA DE "Certificado.xlsx" Y GENERACIÓN DE CSVs -------------------
st.title("Generador de Archivos CSV")

archivo_procesado = st.file_uploader("Cargar archivo procesado (Certificado.xlsx)", type=["xlsx"])

if archivo_procesado is not None:
    wb = openpyxl.load_workbook(archivo_procesado, data_only=True)
    sheetnames = wb.sheetnames
    hoja_principal = wb[sheetnames[0]]
    nombre_base = hoja_principal["A28"].value

    if not nombre_base:
        st.error("Error: No se encontró un valor en la celda A28.")
    else:
        data = []
        for row in hoja_principal.iter_rows(min_row=27, values_only=True):
            data.append(row)

        df = pd.DataFrame(data)
        if df.shape[1] < 18:
            st.error("El archivo no contiene suficientes columnas.")
        else:
            df.columns = df.iloc[0]
            df = df[1:].reset_index(drop=True)

            if "QC_Type" not in df.columns:
                st.error("No se encontró la columna 'QC_Type'.")
            else:
                responsable = plantilla_seleccionada  # Utilizamos la selección de la primera parte

                df1 = df[~df["QC_Type"].isin(["DSTD", "DEND"])]
                columnas1 = ["Holeid", "From", "To", "Sample number", "Displaced volume (g)", "Wet weight (g)",
                             "Dry weight (g)", "Coated dry weight (g)", "Weight in water (g)", "Coated weight in water (g)",
                             "Coat density", "moisture", "Determination method", "Laboratory", "Date", "Responsible", "comments"]
                df1 = df1[list(columnas1)]
                df1["Responsible"] = responsable
                df1.insert(13, "Laboratory", "")

                df2 = df[df["QC_Type"] == "DEND"]
                columnas2 = ["hole_number", "depth_from", "depth_to", "sample", "displaced_volume_g_D",
                             "dry_weight_g_D", "coated_dry_weight_g_D", "weight_water_g", "coated_wght_water_g",
                             "coat_density", "QC_type", "determination_method", "density_date", "responsible", "comments"]
                df2 = df2[list(columnas2)]
                df2["responsible"] = responsable

                df3 = df[df["QC_Type"] == "DSTD"]
                columnas3 = ["hole_number", "displaced_volume_g", "dry_weight_g", "coated_dry_weight_g",
                             "weight_water_g", "coated_wght_water_g", "coat_density", "DSTD_id",
                             "determination_method", "density_date", "responsable", "comments"]
                df3 = df3[list(columnas3)]
                df3["responsable"] = responsable

                for df, suffix in [(df1, ""), (df2, "__QC-DUP"), (df3, "__QC-STD")]:
                    output = BytesIO()
                    df.to_csv(output, index=False, encoding="utf-8-sig")
                    output.seek(0)
                    st.download_button(
                        label=f"Descargar {nombre_base}{suffix}.csv",
                        data=output,
                        file_name=f"{nombre_base}{suffix}.csv",
                        mime="text/csv"
                    )

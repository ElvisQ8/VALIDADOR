import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

def procesar_archivo(archivo_cargado, plantilla, responsable):
    # Cargar archivo de plantilla seleccionado
    plantilla_wb = openpyxl.load_workbook(plantilla)
    plantilla_ws = plantilla_wb["PECLD07792"]
    duplicado_ws = plantilla_wb["Duplicado"]
    
    # Cargar archivo CUALQUIERA.xlsx
    wb = openpyxl.load_workbook(archivo_cargado, data_only=True)
    if "BD_densidad_2020" not in wb.sheetnames:
        st.error("El archivo cargado no contiene la hoja 'BD_densidad_2020'")
        return None, None
    ws = wb["BD_densidad_2020"]
    
    # Leer los datos desde A10 hasta R en la hoja BD_densidad_2020
    datos = []
    for row in ws.iter_rows(min_row=10, max_col=18, values_only=True):
        if any(row):  # Solo tomar filas con datos
            datos.append(row)
    
    # Pegar datos en la plantilla desde A28
    start_row = 28
    for i, row in enumerate(datos, start=start_row):
        for j, value in enumerate(row, start=1):
            plantilla_ws.cell(row=i, column=j, value=value)
    
    # Eliminar filas en blanco debajo de los datos pegados
    max_row = plantilla_ws.max_row
    for i in range(i + 1, max_row + 1):
        plantilla_ws.delete_rows(i)
    
    # Filtrar datos para CSVs
    df = pd.DataFrame(plantilla_ws.values)
    df.columns = df.iloc[26]
    df = df[27:].reset_index(drop=True)
    
    df1 = df[~df["QC_Type"].isin(["DSTD", "DEND"])][["Holeid", "From", "To", "Sample number", "Displaced volume (g)", "Wet weight (g)", "Dry weight (g)", "Coated dry weight (g)", "Weight in water (g)", "Coated weight in water (g)", "Coat density", "moisture", "Determination method", "Date", "comments"]]
    df1.insert(13, "Laboratory", "")
    df1.insert(14, "Responsible", responsable)
    
    df2 = df[df["QC_Type"] == "DEND"][["hole_number", "depth_from", "depth_to", "sample", "displaced_volume_g_D", "dry_weight_g_D", "coated_dry_weight_g_D", "weight_water_g", "coated_wght_water_g", "coat_density", "QC_type", "determination_method", "density_date", "comments"]]
    df2.insert(12, "Responsible", responsable)
    
    df3 = df[df["QC_Type"] == "DSTD"][["hole_number", "displaced_volume_g", "dry_weight_g", "coated_dry_weight_g", "weight_water_g", "coated_wght_water_g", "coat_density", "DSTD_id", "determination_method", "density_date", "comments"]]
    df3.insert(10, "Responsible", responsable)
    
    return plantilla_wb, df1, df2, df3

st.title("Editor de Archivos Excel")

# Selección de plantilla
opciones_plantilla = {
    "Nombre 1": "PLANTILLA.xlsx",
    "Nombre 2": "https://raw.githubusercontent.com/tu_usuario/tu_repositorio/main/PLANTILLA1.xlsx",
    "Nombre 3": "https://raw.githubusercontent.com/tu_usuario/tu_repositorio/main/PLANTILLA2.xlsx"
}

plantilla_seleccionada = st.selectbox("Selecciona una plantilla:", list(opciones_plantilla.keys()))
plantilla_path = opciones_plantilla[plantilla_seleccionada]
responsable = plantilla_seleccionada

# Subir archivo
archivo_cargado = st.file_uploader("Carga un archivo Excel", type=["xlsx"])

if archivo_cargado is not None:
    plantilla_wb, df1, df2, df3 = procesar_archivo(archivo_cargado, plantilla_path, responsable)
    
    if plantilla_wb:
        output = BytesIO()
        plantilla_wb.save(output)
        output.seek(0)
        st.download_button(label="Descargar archivo procesado", data=output, file_name="Resultado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        csv1 = df1.to_csv(index=False).encode("utf-8")
        csv2 = df2.to_csv(index=False).encode("utf-8")
        csv3 = df3.to_csv(index=False).encode("utf-8")
        
        st.download_button("Descargar CSV 1", data=csv1, file_name=f"{df.iloc[27, 0]}.csv", mime="text/csv")
        st.download_button("Descargar CSV 2", data=csv2, file_name=f"{df.iloc[27, 0]}__QC-DUP.csv", mime="text/csv")
        st.download_button("Descargar CSV 3", data=csv3, file_name=f"{df.iloc[27, 0]}__QC-STD.csv", mime="text/csv")

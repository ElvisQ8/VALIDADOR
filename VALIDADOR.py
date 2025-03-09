import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

def generar_csvs(archivo_procesado, responsable):
    wb = openpyxl.load_workbook(archivo_procesado, data_only=True)
    ws = wb[wb.sheetnames[0]]  # Tomar la primera hoja
    
    # Obtener nombre base desde A28
    nombre_base = ws["A28"].value
    if not nombre_base:
        st.error("No se encontró un nombre válido en la celda A28.")
        return None, None, None
    
    # Obtener cabeceras desde fila 27
    cabeceras = [cell.value for cell in ws[27] if cell.value is not None]
    
    # Obtener datos desde fila 28 en adelante
    datos = []
    for row in ws.iter_rows(min_row=28, values_only=True):
        if any(row):
            datos.append(row)
    
    df = pd.DataFrame(datos, columns=cabeceras)
    
    # Filtrado para cada archivo y reordenación de columnas
    if "QC_Type" not in df.columns:
        st.error("No se encontró la columna 'QC_Type'. Asegúrate de que la estructura es correcta.")
        return None, None, None
    
    # Archivo 1: Filtrar datos que NO contengan "DSTD" o "DEND" en QC_Type
    df1 = df[~df["QC_Type"].isin(["DSTD", "DEND"])]
    df1 = df1[["Holeid", "From", "To", "Sample number", "Displaced volume (g)", "Wet weight (g)",
               "Dry weight (g)", "Coated dry weight (g)", "Weight in water (g)", "Coated weight in water (g)",
               "Coat density", "moisture", "Determination method", "Date", "comments"]]
    df1.insert(13, "Laboratory", "")  # Agregar columna vacía
    df1.insert(15, "Responsible", responsable)  # Agregar responsable
    
    # Archivo 2: Filtrar datos donde QC_Type sea "DEND"
    df2 = df[df["QC_Type"] == "DEND"]
    df2 = df2[["hole_number", "depth_from", "depth_to", "sample", "displaced_volume_g_D", "dry_weight_g_D", 
               "coated_dry_weight_g_D", "weight_water_g", "coated_wght_water_g", "coat_density", "QC_Type", 
               "determination_method", "density_date", "comments"]]
    df2.insert(13, "responsible", responsable)
    
    # Archivo 3: Filtrar datos donde QC_Type sea "DSTD"
    df3 = df[df["QC_Type"] == "DSTD"]
    df3 = df3[["hole_number", "displaced_volume_g", "dry_weight_g", "coated_dry_weight_g", "weight_water_g", 
               "coated_wght_water_g", "coat_density", "DSTD_id", "determination_method", "density_date", "comments"]]
    df3.insert(10, "responsible", responsable)
    
    # Convertir dataframes a CSV en memoria
    def convertir_a_csv(df):
        output = BytesIO()
        df.to_csv(output, index=False, encoding='utf-8')
        output.seek(0)
        return output
    
    return (convertir_a_csv(df1), convertir_a_csv(df2), convertir_a_csv(df3), nombre_base)

# Interfaz en Streamlit
st.title("Validador de registro de datos - densidad")

# Selección de plantilla
opciones_plantilla = {
    "ROSA LA PRIMOROSA": "PLANTILLA.xlsx",
    "MILAGROS CHAMPIRREINO": "PLANTILLA1.xlsx",
    "YONATAN CON Y": "PLANTILLA2.xlsx"
}

plantilla_seleccionada = st.selectbox("Seleccione el responsable:", list(opciones_plantilla.keys()))
plantilla_path = opciones_plantilla[plantilla_seleccionada]

# Subir archivo procesado
archivo_procesado = st.file_uploader("Carga archivo procesado (Certificado.xlsx)", type=["xlsx"])

if archivo_procesado is not None:
    csv1, csv2, csv3, nombre_base = generar_csvs(archivo_procesado, plantilla_seleccionada)
    
    if csv1 and csv2 and csv3:
        st.download_button(label=f"Descargar {nombre_base}.csv", data=csv1, file_name=f"{nombre_base}.csv", mime="text/csv")
        st.download_button(label=f"Descargar {nombre_base}__QC-DUP.csv", data=csv2, file_name=f"{nombre_base}__QC-DUP.csv", mime="text/csv")
        st.download_button(label=f"Descargar {nombre_base}__QC-STD.csv", data=csv3, file_name=f"{nombre_base}__QC-STD.csv", mime="text/csv")

import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

def procesar_archivo(archivo_cargado, plantilla):
    # Cargar archivo PLANTILLA.xlsx
    plantilla_wb = openpyxl.load_workbook(plantilla)
    plantilla_ws = plantilla_wb["PECLD07792"]
    duplicado_ws = plantilla_wb["Duplicado"]
    
    # Cargar archivo CUALQUIERA.xlsx
    wb = openpyxl.load_workbook(archivo_cargado, data_only=True)
    if "BD_densidad_2020" not in wb.sheetnames:
        st.error("El archivo cargado no contiene la hoja 'BD_densidad_2020'")
        return None
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
    
    # Aplicar color de relleno a las filas que contengan "DSTD" o "DEND" en alguna celda
    from openpyxl.styles import PatternFill
    fill = PatternFill(start_color="E26B0A", end_color="E26B0A", fill_type="solid")
    
    for row in plantilla_ws.iter_rows(min_row=28, max_col=18):
        if any(cell.value in ["DSTD", "DEND"] for cell in row):
            for cell in row:
                cell.fill = fill
    
    # Buscar "DEND" en la columna "O" y copiar valores a la hoja "Duplicado"
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
    
    # Cambiar el nombre de la hoja "PECLD07792" por el valor de la celda A28
    nuevo_nombre = plantilla_ws.cell(row=28, column=1).value
    if nuevo_nombre:
        plantilla_ws.title = str(nuevo_nombre)
    
    # Guardar cambios en un BytesIO para permitir la descarga
    output = BytesIO()
    plantilla_wb.save(output)
    output.seek(0)
    return output

st.title("Editor de Archivos Excel")

# Cargar plantilla en el servidor
plantilla_path = "PLANTILLA.xlsx"

# Subir archivo
archivo_cargado = st.file_uploader("Carga un archivo Excel", type=["xlsx"])

if archivo_cargado is not None:
    output = procesar_archivo(archivo_cargado, plantilla_path)
    if output:
        st.download_button(label="Descargar archivo procesado", data=output, file_name="Resultado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

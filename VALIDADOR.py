import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

def procesar_archivo(archivo_cargado, plantilla):
    # Cargar archivo de plantilla seleccionado
    plantilla_wb = openpyxl.load_workbook(plantilla)
    plantilla_ws = plantilla_wb["PECLD07792"]
    duplicado_ws = plantilla_wb["Duplicado"]
    standar_ws = plantilla_wb["STD (PECLSTDEN02)"]
    
    # Cargar archivo CUALQUIERA.xlsx
    wb = openpyxl.load_workbook(archivo_cargado, data_only=True)
    if "BD_densidad_2020" not in wb.sheetnames:
        st.error("El archivo cargado no contiene la hoja 'BD_densidad_2020'")
        return None
    ws = wb["BD_densidad_2020"]
    
    # Leer los datos desde A10 hasta R en la hoja BD_densidad_2020
    datos = []
    for row in ws.iter_rows(min_row=10, max_col=17, values_only=True):
        if any(row):  # Solo tomar filas con datos
            datos.append(row)
    
    # Pegar datos en la plantilla desde A28
    start_row = 28
    for i, row in enumerate(datos, start=start_row):
        for j, value in enumerate(row, start=1):
            plantilla_ws.cell(row=i, column=j, value=value)
    
    # Eliminar filas en blanco debajo de los datos pegados
    max_row = plantilla_ws.max_row
    columns = 17  # Número de columnas en la hoja

    filas_eliminar = []
    for i in range(start_row + len(datos), max_row + 1):
        # Verificar si la fila está completamente vacía
        if all(plantilla_ws.cell(row=i, column=j).value is None for j in range(1, columns + 1)):
            filas_eliminar.append(i)

    # Eliminar filas en orden inverso para evitar problemas con los índices
    for row in reversed(filas_eliminar):
        plantilla_ws.delete_rows(row)

    # Aplicar color de relleno a las filas que contengan "DSTD" o "DEND" en alguna celda
    from openpyxl.styles import PatternFill
    fill = PatternFill(start_color="E26B0A", end_color="E26B0A", fill_type="solid")
    
    for row in plantilla_ws.iter_rows(min_row=28, max_col=17):
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
    valor_d1 = plantilla_ws.cell(row=28, column=17).value
    valor_d2 = plantilla_ws.cell(row=28, column=13).value
    standar_ws.cell(row=11, column=2, value=valor_d1)
    standar_ws.cell(row=11, column=4, value=valor_d2)
    nuevo_nombre = plantilla_ws.cell(row=28, column=1).value
    if nuevo_nombre:
        plantilla_ws.title = str(nuevo_nombre)
    
    # Guardar cambios en un BytesIO para permitir la descarga
    output = BytesIO()
    plantilla_wb.save(output)
    output.seek(0)
    return output

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

if archivo_cargado is not None:
    output = procesar_archivo(archivo_cargado, plantilla_path)
    if output:
        st.download_button(
            label="Descargar archivo procesado",
            data=output,
            file_name="Certificado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

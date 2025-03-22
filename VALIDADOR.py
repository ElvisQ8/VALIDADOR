import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# Código del primer archivo (Validador de Registro de Datos)
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

# Código del segundo archivo (Exportador de Datos a Plantilla)
def load_data(file_path):
    return pd.read_excel(file_path, sheet_name=0, header=None, skiprows=27, usecols="A:R", nrows=101)

def clean_data(df):
    # Eliminar las celdas con la palabra 'hola'
    df_cleaned = df[df != 'hola']
    return df_cleaned

def copy_data_to_template(df, sheet_name, selected_name, template_file):
    template = pd.ExcelFile(template_file)

    with BytesIO() as output:
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Escribir solo la hoja seleccionada de la plantilla (sin sobrescribir la primera fila)
            if sheet_name in template.sheet_names:
                temp_df = template.parse(sheet_name)
                temp_df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Filtrar y copiar los datos según el tipo de hoja (O, DP, STD)
            if sheet_name == "O":
                df_filtered = df[~df[14].str.contains('DSTD|DEND', na=False)]
                df_filtered[13] = selected_name  # Colocar el nombre en la columna "N" de la hoja "O"
                # Mapeo de columnas para "O"
                columns_mapping_o = [
                    (0, 'A'), (1, 'B'), (2, 'C'), (3, 'D'), (4, 'E'),
                    (5, 'F'), (6, 'G'), (7, 'H'), (8, 'I'), (9, 'J'),
                    (10, 'K'), (11, 'L'), (13, 'P'), (16, 'O')
                ]
                for df_col, template_col in columns_mapping_o:
                    writer.sheets[sheet_name].write_column(f'{template_col}2', df_filtered[df_col].fillna('').astype(str).values)

            elif sheet_name == "DP":
                df_filtered = df[df[14].str.contains('DEND', na=False)]
                df_filtered[13] = selected_name  # Colocar el nombre en la columna "K" de la hoja "DP"
                # Mapeo de columnas para "DP"
                columns_mapping_dp = [
                    (0, 'A'), (1, 'B'), (2, 'C'), (3, 'D'), (4, 'E'),
                    (6, 'F'), (7, 'G'), (8, 'H'), (9, 'I'), (10, 'J'),
                    (14, 'K'), (13, 'L'), (16, 'N'), (17, 'O')
                ]
                for df_col, template_col in columns_mapping_dp:
                    writer.sheets[sheet_name].write_column(f'{template_col}2', df_filtered[df_col].fillna('').astype(str).values)

            elif sheet_name == "STD":
                df_filtered = df[df[14].str.contains('DSTD', na=False)]
                df_filtered[13] = selected_name  # Colocar el nombre en la columna "K" de la hoja "STD"
                # Mapeo de columnas para "STD"
                columns_mapping_std = [
                    (0, 'A'), (4, 'B'), (6, 'C'), (7, 'D'), (8, 'E'),
                    (9, 'F'), (10, 'G'), (11, 'H'), (13, 'K'), (16, 'J'),
                    (17, 'L')
                ]
                # En este caso, el valor "PECLSTDEN02" lo tratamos como un valor específico
                peclstd_value = "PECLSTDEN02"
                writer.sheets[sheet_name].write_column('H2', [peclstd_value] * len(df_filtered))  # Asignamos este valor en la columna H
                for df_col, template_col in columns_mapping_std:
                    writer.sheets[sheet_name].write_column(f'{template_col}2', df_filtered[df_col].fillna('').astype(str).values)

        # Convertir el archivo a CSV para la descarga
        output.seek(0)  # Asegurarse de que el flujo esté al principio
        df_csv = pd.read_excel(output, sheet_name=sheet_name)  # Leer el archivo modificado en el buffer
        csv_output = BytesIO()
        df_csv.to_csv(csv_output, index=False, sep=';', encoding='utf-8')  # Convertir a CSV
        csv_output.seek(0)
        return csv_output.getvalue()

# Crear la interfaz de usuario
st.title("Aplicación de Datos")

# Barra lateral con opciones de menú
pagina = st.sidebar.radio("Selecciona una página", ["Validador de Datos", "Exportador"])

# Página de "Validador de Datos"
if pagina == "Validador de Datos":
    st.subheader("Bienvenido al Validador de Registro de Datos")
    
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

# Página de "Exportador"
elif pagina == "Exportador":
    st.subheader("Bienvenido al Exportador")
    
    names = ["nombre1", "nombre2", "nombre3"]
    selected_name = st.selectbox("Selecciona un nombre", names)

    uploaded_file = st.file_uploader("Sube el archivo .xlsx", type=["xlsx"])
    template_file = "plantilla_export.xlsx"

    if uploaded_file is not None:
        df = load_data(uploaded_file)

        if st.button('Exportar Hoja O'):
            df_cleaned = clean_data(df)
            file_o = copy_data_to_template(df_cleaned, "O", selected_name, template_file)
            st.download_button("Descargar Hoja O como CSV", data=file_o, file_name="plantilla_O.csv")

        if st.button('Exportar Hoja DP'):
            df_cleaned = clean_data(df)
            file_dp = copy_data_to_template(df_cleaned, "DP", selected_name, template_file)
            st.download_button("Descargar Hoja DP como CSV", data=file_dp, file_name="plantilla_DP.csv")

        if st.button('Exportar Hoja STD'):
            df_cleaned = clean_data(df)
            file_std = copy_data_to_template(df_cleaned, "STD", selected_name, template_file)
            st.download_button("Descargar Hoja STD como CSV", data=file_std, file_name="plantilla_STD.csv")

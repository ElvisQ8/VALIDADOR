# PROYECTO ELVIS Q.
import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import plotly.graph_objects as go
from io import BytesIO

# CERTIFICADO
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
    datos = [row for row in ws.iter_rows(min_row=10, max_col=17, values_only=True) if any(row)]

    start_row = 28
    for i, row in enumerate(datos, start=start_row):
        for j, value in enumerate(row, start=1):
            plantilla_ws.cell(row=i, column=j, value=value)

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
            duplicado_ws.cell(row=dest_row, column=1, value=plantilla_ws.cell(row=fila_actual, column=4).value)
            duplicado_ws.cell(row=dest_row, column=3, value=plantilla_ws.cell(row=fila_actual, column=4).value)
            duplicado_ws.cell(row=dest_row, column=4, value=plantilla_ws.cell(row=fila_actual, column=13).value)
            duplicado_ws.cell(row=dest_row, column=2, value=plantilla_ws.cell(row=fila_actual - 1, column=13).value)
            dest_row += 1

    valor_d1 = plantilla_ws.cell(row=28, column=17).value
    valor_d2 = plantilla_ws.cell(row=28, column=13).value

    
    final_data_row = start_row + len(datos)
    plantilla_ws.delete_rows(final_data_row + 0, 80)
    standar_ws.cell(row=11, column=2, value=valor_d1)
    standar_ws.cell(row=11, column=4, value=valor_d2)

    nuevo_nombre = plantilla_ws.cell(row=28, column=1).value
    if nuevo_nombre:
        plantilla_ws.title = str(nuevo_nombre)

    output = BytesIO()
    plantilla_wb.save(output)
    output.seek(0)
    return output

# INTERFAZ 
st.title("OPTIMIZACION DE PROCESOS - DENSIDADES")
pagina = st.sidebar.radio("Selecciona un proceso", ["Generar certificado", "Exportador"])

if pagina == "Generar certificado":
    st.subheader("Generación de certificado")

    opciones_plantilla = {
        "ARTURO": "PLANTILLA.xlsx",
        "MILAGROS ": "PLANTILLA1.xlsx",
        "YONATAN": "PLANTILLA2.xlsx"
    }
    plantilla_seleccionada = st.selectbox("Seleccione el responsable:", list(opciones_plantilla.keys()))
    plantilla_path = opciones_plantilla[plantilla_seleccionada]

    archivo_cargado = st.file_uploader("Carga archivo de datos en Excel", type=["xlsx"])

    if archivo_cargado:
        # certificado
        output = procesar_archivo(archivo_cargado, plantilla_path)
        if output:
            st.download_button("Descargar archivo procesado", data=output, file_name="Certificado.xlsx")

        # ANÁLISIS DE DENSIDADES CL
        st.divider()
        st.subheader("Análisis de Densidades")

        df = pd.read_excel(archivo_cargado, sheet_name=0, header=None)
        df = df.drop(index=np.arange(8)).reset_index(drop=True)
        df.columns = df.iloc[0]
        df = df.drop(index=0).reset_index(drop=True)

        
        df['TIPO DE CONTROL QA/QC'] = df['TIPO DE CONTROL QA/QC'].fillna('ORD')
        df['MUESTRA'] = df['MUESTRA'].fillna('ESTANDAR')

        
        metodo = st.multiselect("Filtrar por MÉTODO DE ANÁLISIS", sorted(df['MÉTODO DE ANÁLISIS'].dropna().unique()))
        tipo_control = st.multiselect("Filtrar por TIPO DE CONTROL QA/QC", sorted(df['TIPO DE CONTROL QA/QC'].dropna().unique()))
        comentario = st.multiselect("Filtrar por DOMINIO", sorted(df['COMENTARIO'].dropna().unique()))

        filtrado = df.copy()
        if metodo:
            filtrado = filtrado[filtrado['MÉTODO DE ANÁLISIS'].isin(metodo)]
        if tipo_control:
            filtrado = filtrado[filtrado['TIPO DE CONTROL QA/QC'].isin(tipo_control)]
        if comentario:
            filtrado = filtrado[filtrado['COMENTARIO'].isin(comentario)]

        # LIMITES DE DATA ENTREGADA POR LOS TECNICOS ( ESTO QUEDA PENDIENTE REAFIRMAR CON MAS ENSAYOS)
        rangos_lito = {
            'D': (2.67, 2.8), 'D1': (2.71, 2.95), 'VD': (2.51, 3.26), 'VM': (2.55, 3.86),
            'SSM': (2.8, 4.2), 'SPB': (3.32, 4.94), 'SPP': (3.51, 4.9),
            'PECLSTDEN02': (2.749, 2.779), 'SSL': (2.8, 4.2), 'SOB': (3.32, 4.94),
            'SOP': (3.51, 4.9), 'VL': (2.51, 3.26)
        }

        estados, comentarios = [], []
        for idx, row in filtrado.iterrows():
            densidad = row['DENSIDAD']
            litologia = row['COMENTARIO']
            if pd.isna(densidad):
                estados.append('Sin Densidad')
                comentarios.append('')
                continue
            if pd.isna(litologia):
                if 2.749 <= densidad <= 2.779:
                    estados.append('Correcto')
                    comentarios.append('Estándar dentro del rango')
                else:
                    estados.append('Fuera de Rango')
                    comentarios.append('Estándar fuera de rango')
            elif litologia in rangos_lito:
                min_val, max_val = rangos_lito[litologia]
                estados.append('Correcto' if min_val <= densidad <= max_val else 'Fuera de Rango')
                comentarios.append('')
            else:
                estados.append('Litología desconocida')
                comentarios.append('')

        filtrado['Estado'] = estados
        filtrado['Comentario Validación'] = comentarios

        
        for idx in range(1, len(filtrado)):
            if filtrado.iloc[idx]['TIPO DE CONTROL QA/QC'] == 'DEND':
                actual = filtrado.iloc[idx]['DENSIDAD']
                anterior = filtrado.iloc[idx - 1]['DENSIDAD']
                if pd.notna(actual) and pd.notna(anterior):
                    var_pct = abs(actual - anterior) / anterior
                    if var_pct > 0.10:
                        filtrado.at[idx, 'Estado'] = 'Error Duplicado'
                        filtrado.at[idx, 'Comentario Validación'] = 'Duplicado fuera del 10%'
                    else:
                        filtrado.at[idx, 'Comentario Validación'] = 'Duplicado dentro del 10%'

        
        st.dataframe(filtrado)

        
        fig = go.Figure()
        for lit, (min_v, max_v) in rangos_lito.items():
            fig.add_shape(type="line", x0=0, x1=len(filtrado), y0=min_v, y1=min_v, line=dict(color="gray", dash="dash"))
            fig.add_shape(type="line", x0=0, x1=len(filtrado), y0=max_v, y1=max_v, line=dict(color="gray", dash="dash"))

        fig.add_trace(go.Scatter(
            x=filtrado['MUESTRA'],
            y=filtrado['DENSIDAD'],
            mode='markers',
            marker=dict(
                color=np.where(filtrado['Estado'].isin(['Fuera de Rango', 'Error Duplicado']), 'red', 'blue'),
                size=8
            ),
            text=filtrado['COMENTARIO'],
            hovertemplate='<b>Muestra:</b> %{x}<br><b>Densidad:</b> %{y}<br><b>Litología:</b> %{text}<extra></extra>'
        ))

        fig.update_layout(title='Validación de Densidades', xaxis_title='MUESTRA', yaxis_title='Densidad')
        st.plotly_chart(fig)

elif pagina == "Exportador":
    st.subheader("Exportador de datos para FUSION")

    # EXPORTAR
    def load_data(file_path):
        return pd.read_excel(file_path, sheet_name=0, header=None, skiprows=27, usecols="A:R", nrows=101)

    def clean_data(df):
        return df[df != 'hola']

    def copy_data_to_template(df, sheet_name, selected_name, template_file):
        template = pd.ExcelFile(template_file)
        with BytesIO() as output:
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                if sheet_name in template.sheet_names:
                    temp_df = template.parse(sheet_name)
                    temp_df.to_excel(writer, sheet_name=sheet_name, index=False)

                if sheet_name == "O":
                    df_filtered = df[~df[14].astype(str).str.contains('DSTD|DEND', na=False)]
                    df_filtered[13] = selected_name
                    columns_mapping_o = [(0, 'A'), (1, 'B'), (2, 'C'), (3, 'D'), (4, 'E'),
                                         (5, 'F'), (6, 'G'), (7, 'H'), (8, 'I'), (9, 'J'),
                                         (10, 'K'), (11, 'L'), (13, 'P'), (16, 'O')]
                    for df_col, template_col in columns_mapping_o:
                        writer.sheets[sheet_name].write_column(f'{template_col}2', df_filtered[df_col].fillna('').astype(str).values)

                elif sheet_name == "DP":
                    df_filtered = df[df[14].astype(str).str.contains('DEND', na=False)]
                    df_filtered[13] = selected_name
                    columns_mapping_dp = [(0, 'A'), (1, 'B'), (2, 'C'), (3, 'D'), (4, 'E'),
                                          (6, 'F'), (7, 'G'), (8, 'H'), (9, 'I'), (10, 'J'),
                                          (14, 'K'), (13, 'N'), (16, 'M'), (17, 'O')]
                    for df_col, template_col in columns_mapping_dp:
                        writer.sheets[sheet_name].write_column(f'{template_col}2', df_filtered[df_col].fillna('').astype(str).values)

                elif sheet_name == "STD":
                    df_filtered = df[df[14].astype(str).str.contains('DSTD', na=False)]
                    df_filtered[13] = selected_name
                    columns_mapping_std = [(0, 'A'), (4, 'B'), (6, 'C'), (7, 'D'), (8, 'E'),
                                           (9, 'F'), (10, 'G'), (11, 'H'), (13, 'K'), (16, 'J'), (17, 'L')]
                    peclstd_value = "PECLSTDEN02"
                    writer.sheets[sheet_name].write_column('H2', [peclstd_value] * len(df_filtered))
                    for df_col, template_col in columns_mapping_std:
                        writer.sheets[sheet_name].write_column(f'{template_col}2', df_filtered[df_col].fillna('').astype(str).values)

            output.seek(0)
            df_csv = pd.read_excel(output, sheet_name=sheet_name)
            csv_output = BytesIO()
            df_csv.to_csv(csv_output, index=False, sep=';', encoding='utf-8')
            csv_output.seek(0)
            return csv_output.getvalue()

    names = ["AJGU", "MIAP", "RYSA"]
    selected_name = st.selectbox("Selecciona un usuario", names)
    uploaded_file = st.file_uploader("Cargar el certificado .xlsx", type=["xlsx"])
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

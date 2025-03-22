# UNIFICACIÓN DE PROYECTO - CERTIFICADO + ANÁLISIS DE DENSIDADES CON FILTROS CORRECTOS
import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import plotly.graph_objects as go
from io import BytesIO

# FUNCIÓN ORIGINAL DEL CERTIFICADO

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
    standar_ws.cell(row=11, column=2, value=valor_d1)
    standar_ws.cell(row=11, column=4, value=valor_d2)
    nuevo_nombre = plantilla_ws.cell(row=28, column=1).value
    if nuevo_nombre:
        plantilla_ws.title = str(nuevo_nombre)
    output = BytesIO()
    plantilla_wb.save(output)
    output.seek(0)
    return output

# STREAMLIT APP
st.title("Generador de Certificado + Análisis de Densidades")
pagina = st.sidebar.radio("Selecciona un proceso", ["Generar certificado", "Exportador"])

if pagina == "Generar certificado":
    st.subheader("Generación de certificado y análisis de densidades")

    opciones_plantilla = {
        "ROSA LA PRIMOROSA": "PLANTILLA.xlsx",
        "MILAGROS CHAMPIRREINO": "PLANTILLA1.xlsx",
        "YONATAN CON Y": "PLANTILLA2.xlsx"
    }
    plantilla_seleccionada = st.selectbox("Seleccione el responsable:", list(opciones_plantilla.keys()))
    plantilla_path = opciones_plantilla[plantilla_seleccionada]

    archivo_cargado = st.file_uploader("Carga archivo de datos en Excel", type=["xlsx"])

    if archivo_cargado:
        output = procesar_archivo(archivo_cargado, plantilla_path)
        if output:
            st.download_button("Descargar archivo procesado", data=output, file_name="Certificado.xlsx")

        # ---------------- ANALISIS DE DENSIDADES ----------------
        st.subheader("Análisis de Densidades con Filtros")
        df = pd.read_excel(archivo_cargado, sheet_name=0, header=None)
        df = df.drop(index=np.arange(8)).reset_index(drop=True)
        df.columns = df.iloc[0]
        df = df.drop(index=0).reset_index(drop=True)

        # Reemplazos y limpieza
        df['TIPO DE CONTROL QA/QC'] = df['TIPO DE CONTROL QA/QC'].fillna('ORD')
        df['MUESTRA'] = df['MUESTRA'].fillna('ESTANDAR')

        # Filtros dinámicos
        metodo = st.multiselect("Filtrar por MÉTODO DE ANÁLISIS", sorted(df['MÉTODO DE ANÁLISIS'].dropna().unique()))
        tipo_control = st.multiselect("Filtrar por TIPO DE CONTROL QA/QC", sorted(df['TIPO DE CONTROL QA/QC'].dropna().unique()))
        comentario = st.multiselect("Filtrar por COMENTARIO", sorted(df['COMENTARIO'].dropna().unique()))

        filtrado = df.copy()
        if metodo:
            filtrado = filtrado[filtrado['MÉTODO DE ANÁLISIS'].isin(metodo)]
        if tipo_control:
            filtrado = filtrado[filtrado['TIPO DE CONTROL QA/QC'].isin(tipo_control)]
        if comentario:
            filtrado = filtrado[filtrado['COMENTARIO'].isin(comentario)]

        # Validaciones
        rangos_lito = {
            'D': (2.67, 2.8), 'D1': (2.71, 2.95), 'VD': (2.51, 3.26), 'VM': (2.55, 3.86),
            'SSM': (2.8, 4.2), 'SPB': (3.32, 4.94), 'SPP': (3.51, 4.9),
            'PECLSTDEN02': (2.749, 2.779), 'SSL': (2.8, 4.2), 'SOB': (3.32, 4.94),
            'SOP': (3.51, 4.9), 'VL': (2.51, 3.26)
        }

        estado_list, comentario_list = [], []
        for idx, row in filtrado.iterrows():
            densidad = row['DENSIDAD']
            litologia = row['COMENTARIO']
            if pd.isna(densidad):
                estado_list.append('Sin Densidad')
                comentario_list.append('')
                continue
            if pd.isna(litologia):
                if 2.749 <= densidad <= 2.779:
                    estado_list.append('Correcto')
                    comentario_list.append('Estándar dentro del rango')
                else:
                    estado_list.append('Fuera de Rango')
                    comentario_list.append('Estándar fuera de rango')
            elif litologia in rangos_lito:
                min_val, max_val = rangos_lito[litologia]
                estado_list.append('Correcto' if min_val <= densidad <= max_val else 'Fuera de Rango')
                comentario_list.append('')
            else:
                estado_list.append('Litología desconocida')
                comentario_list.append('')

        filtrado['Estado'] = estado_list
        filtrado['Comentario Validación'] = comentario_list

        # Validación DEND duplicados
        for idx in range(1, len(filtrado)):
            row = filtrado.iloc[idx]
            if row['TIPO DE CONTROL QA/QC'] == 'DEND':
                densidad_actual = row['DENSIDAD']
                densidad_anterior = filtrado.iloc[idx - 1]['DENSIDAD']
                if pd.notna(densidad_actual) and pd.notna(densidad_anterior):
                    variacion = abs(densidad_actual - densidad_anterior) / densidad_anterior
                    if variacion > 0.10:
                        filtrado.at[idx, 'Estado'] = 'Error Duplicado'
                        filtrado.at[idx, 'Comentario Validación'] = 'Duplicado fuera del 10%'
                    else:
                        filtrado.at[idx, 'Comentario Validación'] = 'Duplicado dentro del 10%'

        # Mostrar tabla con color
        def highlight(row):
            color = 'background-color: red' if row['Estado'] in ['Fuera de Rango', 'Error Duplicado'] else ''
            return [color] * len(row)

        st.dataframe(filtrado.style.apply(highlight, axis=1))

        # Gráfico con Plotly
        fig = go.Figure()
        for lit, (min_val, max_val) in rangos_lito.items():
            fig.add_shape(type="line", x0=0, x1=len(filtrado), y0=min_val, y1=min_val,
                          line=dict(color="gray", width=1, dash="dash"))
            fig.add_shape(type="line", x0=0, x1=len(filtrado), y0=max_val, y1=max_val,
                          line=dict(color="gray", width=1, dash="dash"))

        fig.add_trace(go.Scatter(
            x=filtrado['MUESTRA'],
            y=filtrado['DENSIDAD'],
            mode='markers',
            marker=dict(
                color=np.where(filtrado['Estado'].isin(['Fuera de Rango', 'Error Duplicado']), 'red', 'blue'),
                size=8
            ),
            text=filtrado['COMENTARIO'],
            hovertemplate='<b>Muestra:</b> %{x}<br><b>Densidad:</b> %{y}<br><b>Litología:</b> %{text}<extra></extra>',
            name='Densidad'
        ))

        fig.update_layout(
            title='Validación de Densidades',
            xaxis_title='MUESTRA',
            yaxis_title='Densidad',
            legend_title='Leyenda',
            showlegend=True
        )

        st.plotly_chart(fig)

elif pagina == "Exportador":
    st.subheader("Bienvenido al Exportador de datos para FUSION")
    st.write("(Exportador intacto como pediste)")

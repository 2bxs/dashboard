import streamlit as st
import pandas as pd
from docx import Document
from docx2pdf import convert
import os
import zipfile
import io
import datetime

# Configuración de la página
st.set_page_config(page_title="Generador de Notificaciones de Reexpedición", layout="wide")
st.title("Generador Masivo de Notificaciones (Reexpediciones)")

# 1. Carga de Archivos
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("Sube el Archivo 1 (Excel de registros)", type=["xlsx"])
with col2:
    template_file = st.file_uploader("Sube el Archivo 2 (Plantilla base en .docx)", type=["docx"])

if excel_file and template_file:
    # 2. Procesamiento de Datos con Pandas
    df = pd.read_excel(excel_file)
    
    # Limpiar espacios en blanco al inicio y al final de los nombres de las columnas
    df.columns = df.columns.str.strip()
    
    # Lógica para el Teléfono (Fono 1 -> Fono 2 -> Fono 3)
    if 'FONO 1' in df.columns and 'FONO 2' in df.columns and 'FONO 3' in df.columns:
        df['TELE'] = df['FONO 1'].fillna(df['FONO 2']).fillna(df['FONO 3'])
    else:
        df['TELE'] = "SIN FONO"

    # Agrupar por RUT para crear la tabla resumen
    resumen = df.groupby('RUT').agg(
        Razon_Social=('Razon Social', 'first'),
        Cantidad=('RUT', 'count') # Cuenta cuántas reexpediciones tiene ese RUT
    ).reset_index()

    # Renombrar las columnas para que se vean bien en la tabla interactiva
    resumen = resumen.rename(columns={
        'Razon_Social': 'Nombre de Empresa', 
        'Cantidad': 'Cantidad de Reexpediciones'
    })

    st.divider()
    st.subheader("Configuración de Datos Automáticos")
    
    # Cuadros para llenado automático
    col_a, col_b = st.columns(2)
    with col_a:
        inicio_mr = st.number_input("Inicio de MR (NRM Inicial)", min_value=0, value=0, step=1, help="Ingresa un número mayor a 0 para autocompletar NRM de forma ascendente.")
    with col_b:
        iniciales = st.text_input("Iniciales de quien genera", help="Se autocompletará en la columna GENE.")

    # Generar la fecha actual en formato español "DD de mes de YYYY"
    meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    hoy = datetime.datetime.now()
    fecha_actual_texto = f"{hoy.day:02d} de {meses[hoy.month - 1]} de {hoy.year}"

    # Aplicar los datos automáticos al dataframe antes de mostrarlo
    if inicio_mr > 0:
        resumen['NRM'] = [str(inicio_mr + i) for i in range(len(resumen))]
    else:
        resumen['NRM'] = ""
        
    resumen['FECHA_GEN'] = fecha_actual_texto
    resumen['GENE'] = iniciales

    st.divider()
    st.subheader("Registros Pendientes (Tabla Editable)")
    st.caption("Los datos generados automáticamente pueden ser editados manualmente haciendo clic en las celdas.")

    # 3. Interfaz Frontend Dinámica (Tabla de datos editable)
    # FECHA_GEN pasa a ser TextColumn en lugar de DateColumn para soportar el formato "09 de junio de 2026"
    edited_df = st.data_editor(
        resumen,
        column_config={
            "Nombre de Empresa": st.column_config.TextColumn(disabled=True),
            "RUT": st.column_config.TextColumn(disabled=True),
            "Cantidad de Reexpediciones": st.column_config.NumberColumn(disabled=True),
            "NRM": st.column_config.TextColumn("NRM ✎"),
            "FECHA_GEN": st.column_config.TextColumn("FECHA_GEN ✎"), 
            "GENE": st.column_config.TextColumn("GENE ✎"),
        },
        hide_index=True,
        use_container_width=True

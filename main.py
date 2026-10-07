import streamlit as st
import pandas as pd
import io

# Configuración de página
st.set_page_config(page_title="Reexpedicion Por Cumplir", layout="wide")
st.title("Reexpedicion Por Cumplir")

# 1. Zona de carga de archivos
col1, col2 = st.columns(2)
with col1:
    Zofri = st.file_uploader("Archivo Zofri", type=['xlsx', 'xls', 'csv'])
with col2:
    archivos_2 = st.file_uploader("Archivos Sirote", type=['xlsm', 'xlsx', 'csv'], accept_multiple_files=True)

# Función auxiliar robusta para leer CSV sin errores
def leer_csv_robusto(archivo):
    try:
        archivo.seek(0)
        return pd.read_csv(archivo, sep=None, engine='python')
    except UnicodeDecodeError:
        archivo.seek(0)
        return pd.read_csv(archivo, encoding='latin-1', sep=None, engine='python')

# Función auxiliar para convertir DataFrame a Excel en memoria
def generar_excel(df, nombre_hoja='Para cumplir'):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=nombre_hoja)
    return output.getvalue()

if Zofri and archivos_2:
    try:
        # 2. Carga y consolidación de datos
        if Zofri.name.endswith('.csv'):
            df1 = leer_csv_robusto(Zofri)
        else:
            df1 = pd.read_excel(Zofri)

        lista_df2 = []

        # Procesar todos los archivos subidos en el cuadro 2
        for file in archivos_2:
            nombre_archivo = file.name.lower()

            # --- Lógica especial para el archivo de Punta Arenas ---
            if 'puntaarenas' in nombre_archivo and nombre_archivo.endswith('.xlsx'):
                df_temp = pd.read_excel(file, header=7, engine='openpyxl')

                if 'Numero' in df_temp.columns:
                    mapping = {
                        'Numero': 'Reexpediciones',
                        'Fecha Ingreso': 'Fecha Cierre',
                        'Num_Int': 'N° MIC'
                    }
                    df

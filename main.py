import streamlit as st
import pandas as pd
import io

# Configuración de página
st.set_page_config(page_title="Reexpedicion Por Cumplir", layout="wide")

# CSS personalizado para los bordes redondeados y naranjo candy en las métricas
st.markdown("""
<style>
/* Apuntar al contenedor de las métricas en Streamlit */
div[data-testid="stMetric"] {
    border: 2px solid #FF8C00; /* Color naranjo candy */
    border-radius: 15px;       /* Esquinas redondeadas */
    padding: 15px;             /* Espacio interior para que no quede pegado al borde */
   
}
</style>
""", unsafe_allow_html=True)

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
                    df_temp = df_temp.rename(columns=mapping)

                    # Forzar formato de fecha corta (DD-MM-YYYY) sin horas
                    if 'Fecha Cierre' in df_temp.columns:
                        df_temp['Fecha Cierre'] = pd.to_datetime(df_temp['Fecha Cierre'], errors='coerce').dt.strftime('%d-%m-%Y').fillna(df_temp['Fecha Cierre'])

                    df_temp['Aduana Destino'] = 'Punta Arenas'
                    if 'Patente Tracto' not in df_temp.columns:
                        df_temp['Patente Tracto'] = pd.NA
                else:
                    st.warning(f"El archivo {file.name} se detectó como Punta Arenas, pero no tiene la columna 'Numero' en la fila 8.")

                lista_df2.append(df_temp)

            # Procesar el resto de archivos normalmente
            else:
                if nombre_archivo.endswith('.csv'):
                    df_temp = leer_csv_robusto(file)
                else:
                    df_temp = pd.read_excel(file, engine='openpyxl')
                lista_df2.append(df_temp)

        # Consolidar todo el universo de reexpediciones
        df2 = pd.concat(lista_df2, ignore_index=True)

        # Validaciones de columnas maestras
        if 'documento_salida' not in df1.columns:
            st.error("El Archivo 1 no contiene la columna 'documento_salida'.")
        elif 'Reexpediciones' not in df2.columns:
            st.error("Los archivos de comparación no contienen la columna 'Reexpediciones' (o 'Numero' en el caso de Punta Arenas).")
        else:
            # 3. Limpieza y Comparación
            df1['doc_clean'] = df1['documento_salida'].astype(str).str.replace('-', '', regex=False).str.strip()
            df2['reexp_clean'] = df2['Reexpediciones'].astype(str).str.replace('-', '', regex=False).str.strip()

            df_match = pd.merge(df1, df2, left_on='doc_clean', right_on='reexp_clean', how='inner')

            # Preparar DataFrames para exportación antes de mostrar las métricas
            columnas_reporte = ['N° MIC', 'Aduana Destino', 'Fecha Cierre', 'Patente Tracto', 'Reexpediciones']
            columnas_finales = [col for col in columnas_reporte if col in df_match.columns]

            # DF 1: Todas las coincidencias
            df_todas = df_match[columnas_finales].copy()
            if 'Reexpediciones' in df_todas.columns:
                df_todas['Reexpediciones'] = df_todas['Reexpediciones'].astype(str).str.replace('-', '', regex=False).str.strip()

            # DF 2: Solo con Fecha de Cierre válida
            df_con_fecha_export = df_todas.copy()
            if 'Fecha Cierre' in df_con_fecha_export.columns:
                df_con_fecha_export = df_con_fecha_export.dropna(subset=['Fecha Cierre'])
                df_con_fecha_export = df_con_fecha_export[df_con_fecha_export['Fecha Cierre'].astype(str).str.strip() != '']
                df_con_fecha_export = df_con_fecha_export[df_con_fecha_export['Fecha Cierre'].astype(str).str.strip().str.lower() != 'nan']
                df_con_fecha_export = df_con_fecha_export[df_con_fecha_export['Fecha Cierre'].astype(str).str.strip().str.lower() != 'nat']

            # Calcular el número para la métrica
            if 'Fecha Cierre' in df_match.columns:
                s_str = df_match['Fecha Cierre'].astype(str).str.strip()
                s_low = s_str.str.lower()
                n_con_fecha = int((df_match['Fecha Cierre'].notna() & (s_str != '') & (~s_low.isin(['nan', 'nat', 'none']))).sum())
            else:
                n_con_fecha = 0

            # 4. Métricas y Botones de Descarga
            st.divider()
            m1, m2, m3 = st.columns(3)
            
            with m1:
                st.metric("Universo de Comparación", len(df2))
                
            with m2:
                st.metric("Coincidencias", len(df_match))
                st.write("") # Pequeño espacio para separar la métrica del botón
                
                # Subcolumnas [1, 1] hacen que el botón ocupe la mitad izquierda
                btn_col1, _ = st.columns([1, 1])
                with btn_col1:
                    st.download_button(
                        label="📥 Generar",
                        data=generar_excel(df_todas),
                        file_name="Coincidencias_Totales.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
            with m3:
                st.metric("Reexpediciones (con Fecha de Cierre)", n_con_fecha)
                st.write("") # Pequeño espacio para separar la métrica del botón
                
                # Subcolumnas [1, 1] hacen que el botón ocupe la mitad izquierda
                btn_col2, _ = st.columns([1, 1])
                with btn_col2:
                    st.download_button(
                        label="📥 Genera documento para cumplir",
                        data=generar_excel(df_con_fecha_export),
                        file_name="Reexpediciones para cumplir.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )

    except Exception as e:
        st.error(f"Error procesando los datos: {e}")
else:
    st.info("Sube el Archivo 1 y los Archivos 2 de comparación para iniciar el análisis.")

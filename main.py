import streamlit as st
import pandas as pd
import io

# Configuración de página
st.set_page_config(page_title="Dashboard de Cruce Aduanero", layout="wide")
st.title("Dashboard de Cruce: Documentos de Salida vs Múltiples Reexpediciones")

# 1. Zona de carga de archivos
col1, col2 = st.columns(2)
with col1:
    archivo_1 = st.file_uploader("Sube el Archivo 1 (Documentos de Salida)", type=['xlsx', 'xls', 'csv'])
with col2:
    archivos_2 = st.file_uploader("Sube los Archivos 2 (.xlsm / .xlsx / .csv)", type=['xlsm', 'xlsx', 'csv'], accept_multiple_files=True)

# Función auxiliar robusta para leer CSV sin errores de formato o codificación
def leer_csv_robusto(archivo):
    try:
        archivo.seek(0)
        return pd.read_csv(archivo, sep=None, engine='python')
    except UnicodeDecodeError:
        archivo.seek(0)
        return pd.read_csv(archivo, encoding='latin-1', sep=None, engine='python')

if archivo_1 and archivos_2:
    try:
        # 2. Carga y consolidación de datos
        if archivo_1.name.endswith('.csv'):
            df1 = leer_csv_robusto(archivo_1)
        else:
            df1 = pd.read_excel(archivo_1)

        lista_df2 = []
        for file in archivos_2:
            if file.name.endswith('.csv'):
                df_temp = leer_csv_robusto(file)
            else:
                df_temp = pd.read_excel(file, engine='openpyxl')
            lista_df2.append(df_temp)
        
        df2 = pd.concat(lista_df2, ignore_index=True)

        if 'documento_salida' not in df1.columns:
            st.error("El Archivo 1 no contiene la columna 'documento_salida'.")
        elif 'Reexpediciones' not in df2.columns:
            st.error("Los Archivos 2 consolidados no contienen la columna 'Reexpediciones'.")
        else:
            # 3. Limpieza y Comparación
            df1['doc_clean'] = df1['documento_salida'].astype(str).str.replace('-', '', regex=False).str.strip()
            df2['reexp_clean'] = df2['Reexpediciones'].astype(str).str.replace('-', '', regex=False).str.strip()

            df_match = pd.merge(df1, df2, left_on='doc_clean', right_on='reexp_clean', how='inner')

            # 4. Dashboard de visualización
            st.divider()
            st.subheader("Resumen de Registros")
            
            m1, m2, m3 = st.columns(3)
            m1.metric("Registros en Archivo 1", len(df1))
            m2.metric("Registros en Archivo 2 (Consolidado)", len(df2))
            m3.metric("Coincidencias (Incluye repetidos)", len(df_match))

            st.write("### Detalle de Registros")
            tab1, tab2, tab3 = st.tabs([
                "Coincidencias (Duplicados en ambos)", 
                "Registros Archivo 1", 
                "Registros Archivo 2 (Consolidado)"
            ])
            
            with tab1:
                st.dataframe(df_match, use_container_width=True)
            with tab2:
                st.dataframe(df1[['documento_salida']], use_container_width=True)
            with tab3:
                st.dataframe(df2, use_container_width=True)

            # 5. Generación de Excel
            st.divider()
            st.subheader("Exportar Resultados")
            
            # --- NUEVO: Checkbox para filtrar fechas vacías ---
            excluir_sin_fecha = st.checkbox("Excluir registros que no tengan datos en 'Fecha Cierre'", value=False)
            
            columnas_reporte = ['N° MIC', 'Aduana Destino', 'Fecha Cierre', 'Patente Tracto', 'Reexpediciones']
            columnas_finales = [col for col in columnas_reporte if col in df_match.columns]
            df_final = df_match[columnas_finales]

            # Aplicar el filtro si el checkbox está marcado
            if excluir_sin_fecha and 'Fecha Cierre' in df_final.columns:
                # Eliminar nulos reales (NaN)
                df_final = df_final.dropna(subset=['Fecha Cierre'])
                # Eliminar celdas que parecen vacías pero tienen espacios o la palabra "nan"
                df_final = df_final[df_final['Fecha Cierre'].astype(str).str.strip() != '']
                df_final = df_final[df_final['Fecha Cierre'].astype(str).str.strip().str.lower() != 'nan']
                
                st.info(f"Filtro aplicado: Se exportarán {len(df_final)} registros (se excluyeron los que no tenían Fecha de Cierre).")

            # Crear archivo en memoria
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Coincidencias')
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Generar Excel (coincidencias.xlsx)",
                data=excel_data,
                file_name="coincidencias.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

    except Exception as e:
        st.error(f"Error procesando los datos: {e}")
else:
    st.info("Sube el Archivo 1 y selecciona uno o varios Archivos 2 para iniciar el análisis.")
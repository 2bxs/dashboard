import streamlit as st
import pandas as pd
import io

# Configuración de página
st.set_page_config(page_title="Reexpedicion Por Cumplir", layout="wide")

# Estilos visuales (estilo ArtCraft: oscuro, minimalista, tipografía técnica)
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@400;500;600;620;700&family=Inter:wght@400;500;600&family=Geist+Mono:wght@400;500;700&display=swap');

:root {
    --bg: #060607;
    --bg-raised: #121316;
    --bg-sunken: #0a0a0b;
    --ink: #f2f1ee;
    --ink-strong: #ffffff;
    --muted: rgba(242, 241, 238, 0.62);
    --faint: rgba(242, 241, 238, 0.4);
    --line: rgba(242, 241, 238, 0.16);
    --line-strong: rgba(242, 241, 238, 0.4);
    --accent: #2d81ff;
    --accent-ink: #1659c4;
    --danger: #b3261e;
}

.stApp, [data-testid="stAppViewContainer"], .main, section.main {
    background-color: var(--bg) !important;
    color: var(--ink) !important;
    font-family: 'Inter', system-ui, sans-serif !important;
}
.block-container {
    padding-top: 2.5rem !important;
    padding-bottom: 3rem !important;
    max-width: 1240px !important;
}

h1, h2, h3, h4, .stTitle, [data-testid="stHeader"] {
    font-family: 'Archivo', system-ui, sans-serif !important;
    font-weight: 620 !important;
    letter-spacing: -0.02em !important;
    color: var(--ink-strong) !important;
}
h1, .stTitle { font-size: 2.4rem !important; line-height: 1.05 !important; }
h2, .stSubheader { font-size: 1.6rem !important; line-height: 1.1 !important; }
h3 { font-size: 1.2rem !important; }
p, span, label, div, li { color: var(--ink); }

div.stButton > button, div.stDownloadButton > button {
    background-color: #000000 !important;
    color: var(--ink-strong) !important;
    border: 1px solid var(--line-strong) !important;
    border-radius: 0 !important;
    font-family: 'Geist Mono', ui-monospace, monospace !important;
    font-size: 11px !important;
    font-weight: 700 !important;
    letter-spacing: 1.32px !important;
    text-transform: uppercase !important;
    padding: 10px 24px !important;
    min-width: auto !important;
    width: auto !important;
    box-shadow: none !important;
    transition: background-color 0.15s ease, border-color 0.15s ease !important;
    cursor: pointer !important;
}
div.stButton > button:hover, div.stDownloadButton > button:hover {
    background-color: #1a1a1a !important;
    color: var(--ink-strong) !important;
    border-color: var(--ink) !important;
}
div.stButton > button:focus, div.stDownloadButton > button:focus {
    outline: 1px solid var(--line-strong) !important;
    outline-offset: 2px !important;
    box-shadow: none !important;
}
/* Botón primario (resaltado) */
div.stButton > button[kind="primary"], div.stDownloadButton > button[kind="primary"] {
    background-color: #000000 !important;
    color: var(--ink-strong) !important;
    border: 1px solid var(--ink-strong) !important;
}
div.stButton > button[kind="primary"]:hover, div.stDownloadButton > button[kind="primary"]:hover {
    background-color: #1a1a1a !important;
    color: var(--ink-strong) !important;
    border-color: var(--ink) !important;
    opacity: 1 !important;
}
div.stButton > button[kind="secondary"] {
    background-color: #000000 !important;
    color: var(--ink) !important;
    border: 1px solid var(--line-strong) !important;
}
div.stButton > button[kind="secondary"]:hover {
    background-color: #1a1a1a !important;
    color: var(--ink-strong) !important;
    opacity: 1 !important;
}

.stTextInput input, .stNumberInput input, .stTextInput > div > div > input {
    background-color: var(--bg-raised) !important;
    color: var(--ink) !important;
    border: 1px solid var(--line) !important;
    border-radius: 0 !important;
    font-family: 'Inter', sans-serif !important;
    font-size: 14px !important;
}
.stTextInput input:focus, .stNumberInput input:focus {
    border-color: var(--line-strong) !important;
    box-shadow: none !important;
    outline: none !important;
}
.stTextInput label, .stNumberInput label {
    font-family: 'Geist Mono', monospace !important;
    font-size: 11px !important;
    text-transform: uppercase !important;
    letter-spacing: 1.32px !important;
    color: var(--muted) !important;
}

[data-testid="stFileUploader"] {
    background-color: var(--bg-raised) !important;
    border: 1px dashed var(--line-strong) !important;
    border-radius: 0 !important;
    padding: 1.25rem !important;
}
[data-testid="stFileUploader"] section {
    background-color: transparent !important;
}
[data-testid="stFileUploader"] button {
    background-color: #000000 !important;
    color: var(--ink-strong) !important;
    border: 1px solid var(--line-strong) !important;
    border-radius: 0 !important;
    font-family: 'Geist Mono', monospace !important;
    font-size: 11px !important;
    text-transform: uppercase !important;
    letter-spacing: 1.32px !important;
}
[data-testid="stFileUploader"] small {
    color: var(--muted) !important;
    font-family: 'Geist Mono', monospace !important;
    font-size: 10px !important;
    text-transform: uppercase !important;
    letter-spacing: 1.2px !important;
}

[data-testid="metric-container"] {
    background-color: var(--bg-raised) !important;
    border: 1px solid var(--line) !important;
    border-radius: 0 !important;
    padding: 18px 20px !important;
}
[data-testid="stMetricLabel"] {
    font-family: 'Geist Mono', monospace !important;
    font-size: 10px !important;
    text-transform: uppercase !important;
    letter-spacing: 1.4px !important;
    color: var(--muted) !important;
}
[data-testid="stMetricValue"] {
    font-family: 'Archivo', sans-serif !important;
    font-weight: 620 !important;
    font-size: 1.9rem !important;
    color: var(--ink-strong) !important;
    letter-spacing: -0.02em !important;
}

hr {
    border: none !important;
    border-top: 1px solid var(--line) !important;
    margin: 2.5rem 0 !important;
}

[data-testid="stVerticalBlockBorderWrapper"] {
    border: 1px solid var(--line) !important;
    border-radius: 0 !important;
    background-color: var(--bg-raised) !important;
    padding: 1rem !important;
}

.stCaption, [data-testid="stCaptionContainer"] p {
    color: var(--muted) !important;
    font-family: 'Geist Mono', monospace !important;
    font-size: 11px !important;
    text-transform: uppercase !important;
    letter-spacing: 1.32px !important;
}

[data-testid="stDataFrame"], [data-testid="stDataFrame"] > div {
    border: 1px solid var(--line) !important;
    border-radius: 0 !important;
    background-color: var(--bg-raised) !important;
}

.stSuccess, .stError, .stInfo, .stWarning, [data-testid="stNotification"] {
    border-radius: 0 !important;
    border: 1px solid var(--line-strong) !important;
    background-color: var(--bg-raised) !important;
    color: var(--ink) !important;
    font-family: 'Inter', sans-serif !important;
}
.stSuccess { border-left: 2px solid #ffffff !important; }
.stError { border-left: 2px solid var(--danger) !important; }

.stSpinner > div {
    border-color: var(--ink) transparent transparent transparent !important;
}

::-webkit-scrollbar { width: 8px; height: 8px; }
::-webkit-scrollbar-track { background: var(--bg); }
::-webkit-scrollbar-thumb { background: var(--line-strong); border-radius: 0; }

a { color: var(--accent) !important; text-decoration: none !important; }
a:hover { text-decoration: underline !important; }
</style>
""", unsafe_allow_html=True)

st.title("Reexpedicion Por Cumplir")
st.caption("Free & open source · No subscription required")

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

                df_temp['Archivo Origen'] = file.name
                lista_df2.append(df_temp)

            # Procesar el resto de archivos normalmente
            else:
                if nombre_archivo.endswith('.csv'):
                    df_temp = leer_csv_robusto(file)
                else:
                    df_temp = pd.read_excel(file, engine='openpyxl')
                df_temp['Archivo Origen'] = file.name
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
            columnas_reporte = ['N° MIC', 'Aduana Destino', 'Fecha Cierre', 'Patente Tracto', 'Reexpediciones', 'Archivo Origen']
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

            # Máscara booleana: ¿la fila tiene Fecha Cierre válida?
            if 'Fecha Cierre' in df_match.columns:
                s_str = df_match['Fecha Cierre'].astype(str).str.strip()
                s_low = s_str.str.lower()
                mask_fecha = df_match['Fecha Cierre'].notna() & (s_str != '') & (~s_low.isin(['nan', 'nat', 'none']))
            else:
                mask_fecha = pd.Series(False, index=df_match.index)

            n_con_fecha = int(mask_fecha.sum())

            # 3.1 Resumen: cuántas reexpediciones por cumplir corresponden a cada archivo
            df_resumen = None
            if 'Archivo Origen' in df_match.columns:
                df_resumen = (
                    df_match.assign(_con_fecha=mask_fecha)
                    .groupby('Archivo Origen')
                    .agg(Coincidencias=('_con_fecha', 'size'),
                         Por_Cumplir=('_con_fecha', 'sum'))
                    .reset_index()
                    .rename(columns={'Archivo Origen': 'Archivo', 'Por_Cumplir': 'Por Cumplir'})
                    .sort_values('Por Cumplir', ascending=False)
                )

                # Fila de total
                total = pd.DataFrame([{
                    'Archivo': 'TOTAL',
                    'Coincidencias': int(df_resumen['Coincidencias'].sum()),
                    'Por Cumplir': int(df_resumen['Por Cumplir'].sum())
                }])
                df_resumen = pd.concat([df_resumen, total], ignore_index=True)

            # 4. Tabla resumen por archivo
            if df_resumen is not None:
                st.divider()
                st.subheader("📑 Reexpediciones por cumplir por archivo")
                st.dataframe(df_resumen, use_container_width=True, hide_index=True)

                st.download_button(
                    label="📥 Descargar resumen por archivo",
                    data=generar_excel(df_resumen.drop(index=df_resumen.index[-1]), nombre_hoja='Por archivo'),
                    file_name="Resumen_por_archivo.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # 5. Métricas y Botones de Descarga
            st.divider()
            m1, m2, m3 = st.columns(3)

            with m1:
                st.metric("Universo de Comparación", len(df2))

            with m2:
                st.metric("Coincidencias", len(df_match))
                st.write("")  # Pequeño espacio para separar la métrica del botón

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
                st.write("")  # Pequeño espacio para separar la métrica del botón

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
    st.info("Sube el Archivo 1 y los Archivos 2 de comparación para iniciar el análisis solo sera procesados por lo que no quedara registro en alguna base de datos.")

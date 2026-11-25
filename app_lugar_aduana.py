import streamlit as st
import pandas as pd
import io

# Configuración de página
st.set_page_config(page_title="Lugar Presentación vs Aduana Destino", layout="wide")

# Estilos visuales (mismo estilo que main.py)
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

st.title("Lugar Presentación vs Aduana Destino")
st.caption("Registros donde 'Lugar Presentación' y 'Aduana Destino' coinciden")

# Carga de archivos
archivos = st.file_uploader("Archivos a analizar", type=['xlsx', 'xls', 'csv', 'xlsm'], accept_multiple_files=True)

# Función auxiliar robusta para leer CSV sin errores
def leer_csv_robusto(archivo):
    try:
        archivo.seek(0)
        return pd.read_csv(archivo, sep=None, engine='python')
    except UnicodeDecodeError:
        archivo.seek(0)
        return pd.read_csv(archivo, encoding='latin-1', sep=None, engine='python')

# Función auxiliar para convertir DataFrame a Excel en memoria
def generar_excel(df, nombre_hoja='Coincidentes'):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=nombre_hoja)
    return output.getvalue()

if archivos:
    try:
        lista_df = []
        for file in archivos:
            nombre_archivo = file.name.lower()
            if nombre_archivo.endswith('.csv'):
                df_temp = leer_csv_robusto(file)
            else:
                df_temp = pd.read_excel(file, engine='openpyxl')
            df_temp['Archivo Origen'] = file.name
            lista_df.append(df_temp)

        df_all = pd.concat(lista_df, ignore_index=True)

        col_lp = 'Lugar Presentación'
        col_ad = 'Aduana Destino'

        if col_lp not in df_all.columns or col_ad not in df_all.columns:
            st.error(f"Los archivos deben contener las columnas '{col_lp}' y '{col_ad}'.")
        else:
            # Normalizar para comparar (strip + minúsculas), conservando valores originales
            lp_norm = df_all[col_lp].astype(str).str.strip().str.lower()
            ad_norm = df_all[col_ad].astype(str).str.strip().str.lower()

            # Excluir nulos/vacíos
            validos = df_all[col_lp].notna() & df_all[col_ad].notna()
            mask_igual = validos & (lp_norm != 'nan') & (ad_norm != 'nan') & (lp_norm != '') & (ad_norm != '') & (lp_norm == ad_norm)

            df_coincidentes = df_all[mask_igual].copy()

            # Métricas en la página principal
            st.divider()
            m1, m2 = st.columns(2)
            with m1:
                st.metric("Total de registros", len(df_all))
            with m2:
                st.metric("Registros coincidentes", len(df_coincidentes))

            if len(df_coincidentes) > 0:
                st.download_button(
                    label="📥 Descargar coincidentes",
                    data=generar_excel(df_coincidentes),
                    file_name="LugarPresentacion_AduanaDestino_coincidentes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            else:
                st.info("No se encontraron registros coincidentes.")

            # Tabla mostrada abajo en la página principal
            st.divider()
            st.subheader("📋 Registros donde 'Lugar Presentación' = 'Aduana Destino'")
            st.dataframe(df_coincidentes, use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f"Error procesando los datos: {e}")
else:
    st.info("Sube uno o varios archivos para iniciar el análisis.")

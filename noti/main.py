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

# --- FUNCION REUTILIZABLE PARA GENERAR EL DOCUMENTO ---
def procesar_documento(rut, nrm, fecha_gen, gene, razon_social, df_completo, template_file):
    grupo_rut = df_completo[df_completo['RUT'] == rut]
    correo = grupo_rut['CRREO 1'].iloc[0] if 'CRREO 1' in grupo_rut.columns else ""
    tele = grupo_rut['TELE'].iloc[0]

    template_file.seek(0)
    doc = Document(template_file)
    
    replacements = {
        "[NOM_IMPOR]": str(razon_social),
        "[RUT_IMP]": str(rut),
        "[TELE]": str(tele),
        "[CORREO]": str(correo),
        "[NRM]": nrm,
        "[FECHA_GEN]": fecha_gen,
        "[GENE]": gene
    }
    
    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                p.text = p.text.replace(key, val)
                
    if doc.tables: 
        for table in doc.tables:
            for _, reexp_row in grupo_rut.iterrows():
                cells = table.add_row().cells
                if len(cells) >= 6: 
                    cells[0].text = str(reexp_row.get('REEXP.', ''))
                    cells[1].text = str(reexp_row.get('FECHA DOCUMENTO/VISACION', ''))
                    cells[2].text = str(reexp_row.get('FECHA CONTROL SALIDA', ''))
                    cells[3].text = str(reexp_row.get('AVANZADA ADUANA', ''))
                    cells[4].text = nrm
                    cells[5].text = fecha_gen

    os.makedirs("temp_docs", exist_ok=True)
    docx_path = f"temp_docs/{nrm}.docx"
    pdf_path = f"temp_docs/{nrm}.pdf"
    doc.save(docx_path)
    
    try:
        convert(docx_path, pdf_path)
        return pdf_path, "pdf"
    except Exception:
        return docx_path, "docx"

# 1. Carga de Archivos
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("Sube el Archivo 1 (Excel de registros)", type=["xlsx"])
with col2:
    template_file = st.file_uploader("Sube el Archivo 2 (Plantilla base en .docx)", type=["docx"])

if excel_file and template_file:
    # 2. Procesamiento de Datos con Pandas
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip()
    
    if 'FONO 1' in df.columns and 'FONO 2' in df.columns and 'FONO 3' in df.columns:
        df['TELE'] = df['FONO 1'].fillna(df['FONO 2']).fillna(df['FONO 3'])
    else:
        df['TELE'] = "SIN FONO"

    resumen = df.groupby('RUT').agg(
        Razon_Social=('Razon Social', 'first'),
        Cantidad=('RUT', 'count')
    ).reset_index()

    resumen = resumen.rename(columns={'Razon_Social': 'Nombre de Empresa', 'Cantidad': 'Cantidad de Reexpediciones'})

    # --- CALLBACKS PARA ACTUALIZAR TODA LA TABLA AL MISMO TIEMPO ---
    def actualizar_iniciales():
        nuevo_valor = st.session_state.global_iniciales
        for r in resumen['RUT']:
            st.session_state[f"gen_{r}"] = nuevo_valor

    def actualizar_nrm():
        inicio = st.session_state.global_nrm
        if inicio > 0:
            for i, r in enumerate(resumen['RUT']):
                st.session_state[f"nrm_{r}"] = f"MR-{(inicio + i):03d}"
        else:
            for r in resumen['RUT']:
                st.session_state[f"nrm_{r}"] = ""

    st.divider()
    st.subheader("Configuración de Datos Automáticos")
    
    col_a, col_b = st.columns(2)
    with col_a:
        st.number_input("Inicio de MR (NRM Inicial)", min_value=0, value=0, step=1, key="global_nrm", on

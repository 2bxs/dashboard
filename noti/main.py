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

    # Reiniciar el puntero del archivo de Word por si se lee múltiples veces
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

    st.divider()
    st.subheader("Configuración de Datos Automáticos")
    
    col_a, col_b = st.columns(2)
    with col_a:
        inicio_mr = st.number_input("Inicio de MR (NRM Inicial)", min_value=0, value=0, step=1)
    with col_b:
        iniciales = st.text_input("Iniciales de quien genera")

    meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    hoy = datetime.datetime.now()
    fecha_actual_texto = f"{hoy.day:02d} de {meses[hoy.month - 1]} de {hoy.year}"

    # Formatear a MR-000 usando f-strings y padding de ceros (:03d)
    if inicio_mr > 0:
        resumen['NRM'] = [f"MR-{(inicio_mr + i):03d}" for i in range(len(resumen))]
    else:
        resumen['NRM'] = ""
        
    resumen['FECHA_GEN'] = fecha_actual_texto
    resumen['GENE'] = iniciales

    st.divider()
    st.subheader("Registros Pendientes (Tabla Editable)")

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
    )

    st.divider()

    # --- SECCIÓN DE GENERACIÓN ---
    col_izq, col_der = st.columns([1, 1])

    with col_izq:
        st.subheader("Generación Masiva")
        st.caption("Crea un archivo ZIP con todos los documentos válidos de la tabla.")
        
        if st.button("Generar Todos en ZIP", type="primary"):
            with st.spinner("Generando documentos masivos..."):
                generated_files = []
                
                for index, row in edited_df.iterrows():
                    rut = row['RUT']
                    nrm = str(row['NRM']).strip() if pd.notna(row['NRM']) else ""
                    
                    if not nrm: 
                        continue
                    
                    fecha_gen = str(row['FECHA_GEN']).strip() if pd.notna(row['FECHA_GEN']) else ""
                    gene = str(row['GENE']).strip() if pd.notna(row['GENE']) else ""
                    razon_social = row['Nombre de Empresa']

                    file_path, _ = procesar_documento(rut, nrm, fecha_gen, gene, razon_social, df, template_file)
                    generated_files.append(file_path)
                
                if generated_files:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                        for file_path in generated_files:
                            zip_file.write(file_path, os.path.basename(file_path))
                    
                    st.success("¡Documentos masivos generados!")
                    st.download_button(
                        label="📥 Descargar ZIP",
                        data=zip_buffer.getvalue(),
                        file_name="Notificaciones_Generadas.zip",
                        mime="application/zip"
                    )

    with col_der:
        st.subheader("Generación Individual")
        st.caption("Genera y descarga un registro específico.")
        
        for index, row in edited_df.iterrows():
            rut = row['RUT']
            nrm = str(row['NRM']).strip() if pd.notna(row['NRM']) else ""
            
            if not nrm: 
                continue
            
            c1, c2, c3 = st.columns([5, 3, 3])
            c1.write(f"🏢 {row['Nombre de Empresa']}")
            c2.write(f"📄 {nrm}")
            
            # Usamos session_state para mantener habilitado el botón de descarga en Streamlit
            key_estado = f"file_data_{rut}_{nrm}"
            
            with c3:
                if st.button(f"Generar", key=f"btn_gen_{rut}"):
                    with st.spinner("⏳"):
                        fecha_gen = str(row['FECHA_GEN']).strip() if pd.notna(row['FECHA_GEN']) else ""
                        gene = str(row['GENE']).strip() if pd.notna(row['GENE']) else ""
                        
                        file_path, ext = procesar_documento(rut, nrm, fecha_gen, gene, row['Nombre de Empresa'], df, template_file)
                        
                        with open(file_path, "rb") as f:
                            st.session_state[key_estado] = {
                                "bytes": f.read(),
                                "name": os.path.basename(file_path),
                                "mime": "application/pdf" if ext == "pdf" else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            }
                        st.rerun() # Recarga la página para mostrar el botón de descarga
                        
                if key_estado in st.session_state:
                    file_info = st.session_state[key_estado]
                    st.download_button(
                        label="⬇️ Descargar",
                        data=file_info["bytes"],
                        file_name=file_info["name"],
                        mime=file_info["mime"],
                        key=f"btn_dl_{rut}"
                    )

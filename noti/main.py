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

    # Formatear a MR-000
    if inicio_mr > 0:
        resumen['NRM'] = [f"MR-{(inicio_mr + i):03d}" for i in range(len(resumen))]
    else:
        resumen['NRM'] = ""
        
    resumen['FECHA_GEN'] = fecha_actual_texto
    resumen['GENE'] = iniciales

    st.divider()
    st.subheader("Registros Pendientes")
    
    # --- CONSTRUCCIÓN DE TABLA PERSONALIZADA ---
    # Proporción del ancho de las columnas
    col_widths = [2.5, 1.2, 0.8, 1.2, 1.5, 1.0, 1.5]
    
    # Cabeceras
    h1, h2, h3, h4, h5, h6, h7 = st.columns(col_widths)
    h1.markdown("**Nombre de Empresa**")
    h2.markdown("**RUT**")
    h3.markdown("**Cant.**")
    h4.markdown("**NRM**")
    h5.markdown("**Fecha Gen.**")
    h6.markdown("**Iniciales**")
    h7.markdown("**Acción**")
    
    st.markdown("---") # Línea divisoria visual

    # Lista para capturar los datos y enviarlos al ZIP masivo
    archivos_para_masivo = []

    # Iterar cada registro y crear su propia fila
    for index, row in resumen.iterrows():
        c1, c2, c3, c4, c5, c6, c7 = st.columns(col_widths)
        rut = row['RUT']
        
        c1.write(row['Nombre de Empresa'])
        c2.write(rut)
        c3.write(str(row['Cantidad de Reexpediciones']))
        
        # Cuadros de texto integrados en la tabla (con label oculto)
        nrm = c4.text_input("NRM", value=row['NRM'], key=f"nrm_{rut}", label_visibility="collapsed")
        fecha_gen = c5.text_input("Fecha", value=row['FECHA_GEN'], key=f"fec_{rut}", label_visibility="collapsed")
        gene = c6.text_input("Gene", value=row['GENE'], key=f"gen_{rut}", label_visibility="collapsed")
        
        # Guardamos el estado actual por si presionan el botón Masivo al final
        if nrm:
            archivos_para_masivo.append({
                "rut": rut,
                "nrm": nrm,
                "fecha_gen": fecha_gen,
                "gene": gene,
                "razon_social": row['Nombre de Empresa']
            })

        # Columna de Botones (Generar y Descargar)
        key_estado = f"file_data_{rut}"
        
        with c7:
            if st.button("Generar", key=f"btn_gen_{rut}"):
                if nrm: # Solo genera si la celda NRM tiene datos
                    with st.spinner("⏳"):
                        file_path, ext = procesar_documento(rut, nrm, fecha_gen, gene, row['Nombre de Empresa'], df, template_file)
                        
                        with open(file_path, "rb") as f:
                            st.session_state[key_estado] = {
                                "bytes": f.read(),
                                "name": os.path.basename(file_path),
                                "mime": "application/pdf" if ext == "pdf" else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            }
            
            # Si el documento ya fue generado, mostramos el botón de descarga debajo
            if key_estado in st.session_state:
                file_info = st.session_state[key_estado]
                st.download_button(
                    label="⬇️ Descargar",
                    data=file_info["bytes"],
                    file_name=file_info["name"],
                    mime=file_info["mime"],
                    key=f"btn_dl_{rut}"
                )

    st.divider()

    # --- SECCIÓN DE GENERACIÓN MASIVA (Opcional) ---
    st.subheader("Generación Masiva en ZIP")
    if st.button("Generar Todos los Documentos Listados", type="primary"):
        with st.spinner("Generando documentos masivos..."):
            generated_files = []
            
            for datos in archivos_para_masivo:
                file_path, _ = procesar_documento(
                    datos['rut'], datos['nrm'], datos['fecha_gen'], 
                    datos['gene'], datos['razon_social'], df, template_file
                )
                generated_files.append(file_path)
            
            if generated_files:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                    for file_path in generated_files:
                        zip_file.write(file_path, os.path.basename(file_path))
                
                st.success("¡Documentos masivos generados con éxito!")
                st.download_button(
                    label="📥 Descargar ZIP Completo",
                    data=zip_buffer.getvalue(),
                    file_name="Notificaciones_Generadas.zip",
                    mime="application/zip"
                )

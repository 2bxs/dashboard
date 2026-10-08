import streamlit as st
import pandas as pd
from docx import Document
from docx2pdf import convert
import os
import zipfile
import io

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
    
    # Limpieza y lógica para el Teléfono (Fono 1 -> Fono 2 -> Fono 3)
    # Asumimos que si no hay dato, es NaN. coalesce equivalente:
    if 'FONO 1' in df.columns and 'FONO 2' in df.columns and 'FONO 3' in df.columns:
        df['TELE'] = df['FONO 1'].fillna(df['FONO 2']).fillna(df['FONO 3'])
    else:
        df['TELE'] = "SIN FONO"

    # Agrupar por RUT
    rut_groups = df.groupby('RUT')
    
    st.subheader("Registros por RUT")
    
    # Diccionario para guardar los inputs manuales del usuario
    manual_inputs = {}

    # 3. Interfaz Frontend Dinámica (Lista de cada RUT individual)
    for rut, group in rut_groups:
        razon_social = group['Razon Social'].iloc[0]
        
        with st.expander(f"RUT: {rut} - {razon_social} ({len(group)} reexpediciones)"):
            st.dataframe(group[['REEXP.', 'Fecha Documento/ Visación', 'Fecha Control Salida', 'Avanzada Aduana']])
            
            # Entradas manuales por cada RUT
            c1, c2, c3 = st.columns(3)
            with c1:
                nrm = st.text_input("NRM", key=f"nrm_{rut}")
            with c2:
                fecha_gen = st.date_input("FECHA_GEN", key=f"fecha_{rut}")
            with c3:
                gene = st.text_input("GENE", key=f"gene_{rut}")
                
            manual_inputs[rut] = {
                "nrm": nrm, 
                "fecha_gen": fecha_gen.strftime("%d-%m-%Y"), 
                "gene": gene,
                "data": group,
                "razon_social": razon_social,
                "correo": group['CRREO 1'].iloc[0] if 'CRREO 1' in group.columns else "",
                "tele": group['TELE'].iloc[0]
            }

    # 4. Botón de Generación Masiva
    if st.button("Generar Todos los Documentos (.PDF)", type="primary"):
        with st.spinner("Generando documentos y convirtiendo a PDF..."):
            
            # Crear directorio temporal para guardar archivos
            os.makedirs("temp_docs", exist_ok=True)
            generated_files = []
            
            for rut, info in manual_inputs.items():
                if not info['nrm']: # Saltar si no le han puesto NRM
                    continue
                    
                # Cargar la plantilla original
                doc = Document(template_file)
                
                # REEMPLAZO DE VARIABLES SIMPLES (Párrafos)
                replacements = {
                    "[NOM_IMPOR]": str(info['razon_social']),
                    "[RUT_IMP]": str(rut),
                    "[TELE]": str(info['tele']),
                    "[CORREO]": str(info['correo']),
                    "[NRM]": str(info['nrm']),
                    "[FECHA_GEN]": str(info['fecha_gen']),
                    "[GENE]": str(info['gene'])
                }
                
                for p in doc.paragraphs:
                    for key, val in replacements.items():
                        if key in p.text:
                            p.text = p.text.replace(key, val)
                            
                # LLENADO DE LA TABLA
                # Buscamos la tabla en el documento que contenga las cabeceras objetivo
                for table in doc.tables:
                    # Suponiendo que la tabla que queremos es la primera o tiene un texto específico
                    # Añadimos las filas correspondientes a cada reexpedición
                    for index, row in info['data'].iterrows():
                        cells = table.add_row().cells
                        # Ajusta los índices de 'cells' según la cantidad de columnas de tu Word
                        cells[0].text = str(row['REEXP.'])
                        cells[1].text = str(row['FECHA DOCUMENTO/VISACION'])
                        cells[2].text = str(row['FECHA CONTROL SALIDA'])
                        cells[3].text = str(row['AVANZADA ADUANA'])
                        # Agregar las columnas manuales al cuadro
                        cells[4].text = str(info['nrm'])        # Columna OF. NOTIFICACION
                        cells[5].text = str(info['fecha_gen'])  # Columna FECHA OF. NOTIFIC.

                # Guardar DOCX temporal
                docx_path = f"temp_docs/{info['nrm']}.docx"
                pdf_path = f"temp_docs/{info['nrm']}.pdf"
                doc.save(docx_path)
                
                # Convertir a PDF (Requiere MS Word en el entorno local/Windows)
                try:
                    convert(docx_path, pdf_path)
                    generated_files.append(pdf_path)
                except Exception as e:
                    st.error(f"Error convirtiendo {docx_path} a PDF. ¿Estás en Linux? Necesitarás LibreOffice.")
            
            # Crear un archivo ZIP con todos los PDFs
            if generated_files:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                    for file_path in generated_files:
                        zip_file.write(file_path, os.path.basename(file_path))
                
                st.success("¡Documentos generados con éxito!")
                st.download_button(
                    label="📥 Descargar ZIP con PDFs",
                    data=zip_buffer.getvalue(),
                    file_name="Notificaciones_Generadas.zip",
                    mime="application/zip"
                )

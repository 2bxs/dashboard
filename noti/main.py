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
    )

    # 4. Botón de Generación Masiva
    if st.button("Generar Todos los Documentos (.PDF)", type="primary"):
        with st.spinner("Generando documentos..."):
            os.makedirs("temp_docs", exist_ok=True)
            generated_files = []
            
            # Recorrer la tabla editada por el usuario
            for index, row in edited_df.iterrows():
                rut = row['RUT']
                nrm = str(row['NRM']).strip() if pd.notna(row['NRM']) else ""
                
                # Saltar si la celda NRM está vacía
                if not nrm: 
                    continue
                
                # Rescatar los textos listos de la tabla
                fecha_gen = str(row['FECHA_GEN']).strip() if pd.notna(row['FECHA_GEN']) else ""
                gene = str(row['GENE']).strip() if pd.notna(row['GENE']) else ""
                
                # Obtener los datos originales de ese RUT específico desde el Excel principal
                grupo_rut = df[df['RUT'] == rut]
                razon_social = row['Nombre de Empresa']
                correo = grupo_rut['CRREO 1'].iloc[0] if 'CRREO 1' in grupo_rut.columns else ""
                tele = grupo_rut['TELE'].iloc[0]

                # Cargar la plantilla original
                doc = Document(template_file)
                
                # REEMPLAZO DE VARIABLES SIMPLES EN EL TEXTO
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
                            
                # LLENADO DE LA TABLA DEL WORD
                if doc.tables: 
                    for table in doc.tables:
                        for _, reexp_row in grupo_rut.iterrows():
                            cells = table.add_row().cells
                            if len(cells) >= 6: 
                                cells[0].text = str(reexp_row.get('REEXP.', ''))
                                cells[1].text = str(reexp_row.get('FECHA DOCUMENTO/VISACION', ''))
                                cells[2].text = str(reexp_row.get('FECHA CONTROL SALIDA', ''))
                                cells[3].text = str(reexp_row.get('AVANZADA ADUANA', ''))
                                cells[4].text = nrm        # OF. NOTIFICACION
                                cells[5].text = fecha_gen  # FECHA OF. NOTIFIC.

                # Guardar archivos temporales
                docx_path = f"temp_docs/{nrm}.docx"
                pdf_path = f"temp_docs/{nrm}.pdf"
                doc.save(docx_path)
                
                # Intentar convertir a PDF
                try:
                    convert(docx_path, pdf_path)
                    generated_files.append(pdf_path)
                except Exception as e:
                    st.warning(f"No se pudo convertir {nrm} a PDF. Se adjuntará en formato Word.")
                    generated_files.append(docx_path)
            
            # Crear y descargar ZIP con todos los archivos procesados
            if generated_files:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                    for file_path in generated_files:
                        zip_file.write(file_path, os.path.basename(file_path))
                
                st.success("¡Documentos generados con éxito!")
                st.download_button(
                    label="📥 Descargar ZIP con Documentos",
                    data=zip_buffer.getvalue(),
                    file_name="Notificaciones_Generadas.zip",
                    mime="application/zip"
                )

import streamlit as st
import pandas as pd
from docx import Document
from docx2pdf import convert
import os
import zipfile
import io

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

    # Agregar las columnas vacías para el ingreso manual
    resumen['NRM'] = ""
    resumen['FECHA_GEN'] = None  # None permite el selector de calendario vacío
    resumen['GENE'] = ""

    st.subheader("Registros Pendientes (Haz clic en NRM, FECHA_GEN y GENE para editar)")

    # 3. Interfaz Frontend Dinámica (Tabla de datos editable)
    edited_df = st.data_editor(
        resumen,
        column_config={
            "Nombre de Empresa": st.column_config.TextColumn(disabled=True),
            "RUT": st.column_config.TextColumn(disabled=True),
            "Cantidad de Reexpediciones": st.column_config.NumberColumn(disabled=True),
            "NRM": st.column_config.TextColumn("NRM ✎"),
            "FECHA_GEN": st.column_config.DateColumn(
                "FECHA_GEN ✎",
                format="DD-MM-YYYY", # Formato estricto Día-Mes-Año sin hora
            ),
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
                
                # Saltar si el usuario no ingresó un NRM en la tabla
                if not nrm: 
                    continue
                
                # Formatear la fecha a texto "DD-MM-YYYY" sin hora
                fecha_gen = row['FECHA_GEN'].strftime("%d-%m-%Y") if pd.notna(row['FECHA_GEN']) else ""
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
                if doc.tables: # Verifica que exista al menos una tabla en el Word
                    for table in doc.tables:
                        for _, reexp_row in grupo_rut.iterrows():
                            cells = table.add_row().cells
                            # Validación simple de que la tabla tenga suficientes columnas
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
                    # Si falla la conversión a PDF (ej: en Linux sin LibreOffice), adjuntamos el Word (.docx)
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

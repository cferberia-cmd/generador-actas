import streamlit as st
import datetime
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
import io
import os

st.set_page_config(page_title="Generador de Actas de Obra", layout="wide")

st.title("🏗️ Gestor de Actas de Visita de Obra")
st.write("Rellena los campos para generar automáticamente el documento en formato DOCX.")

# Asegurarse de que la plantilla existe en el directorio
TEMPLATE_PATH = "Plantilla_Acta_Visita.docx"

if not os.path.exists(TEMPLATE_PATH):
    st.error(f"Por favor, asegúrate de colocar el archivo '{TEMPLATE_PATH}' en la misma carpeta que este script.")
else:
    # --- FORMULARIO WEB ---
    col1, col2, col3 = st.columns(3)
    with col1:
        actaNum = st.text_input("Acta Nº:", placeholder="Ej. 01/2026")
        fecha = st.date_input("Fecha de la visita:", datetime.date.today()).strftime("%d/%m/%Y")
    with col2:
        hora = st.text_input("Hora:", placeholder="Ej. 10:00 AM")
        proyecto = st.text_input("Identificación del Proyecto:")
    with col3:
        direccion = st.text_input("Dirección de la obra:")

    st.subheader("👥 Asistentes")
    col_as1, col_as2, col_as3 = st.columns(3)
    with col_as1:
        repreEstudio = st.text_input("Representante del Estudio de Arquitectura:")
    with col_as2:
        EmpConst = st.text_input("Empresa Constructora:")
    with col_as3:
        propiedad = st.text_input("Propiedad:")

    st.subheader("📝 Detalles de la Visita")
    estadoTrabajos = st.text_area("Estado de los Trabajos:")
    instrucciones = st.text_area("Instrucciones Técnicas de Dirección:")
    incidencias = st.text_area("Incidencias y No Conformidades:")
    tareas = st.text_area("Tareas Asignadas con Responsable y Plazo:")

    st.subheader("📸 Fotografías de la Visita")
    # Inicializar el estado para las fotos si no existe
    if 'num_fotos' not in st.session_state:
        st.session_state.num_fotos = 1

    def añadir_foto():
        st.session_state.num_fotos += 1

    fotos_data = []
    for i in range(st.session_state.num_fotos):
        st.markdown(f"**Fotografía {i+1}**")
        col_f1, col_f2 = st.columns([1, 2])
        with col_f1:
            uploaded_foto = st.file_uploader(f"Subir imagen {i+1}", type=["jpg", "jpeg", "png"], key=f"foto_{i}")
        with col_f2:
            desc_foto = st.text_input(f"Descripción Foto {i+1}", key=f"desc_{i}")
        
        if uploaded_foto is not None:
            fotos_data.append({
                "imagen": uploaded_foto,
                "descripcion": desc_foto
            })

    st.button("➕ Añadir otra fotografía", on_click=añadir_foto)

    st.subheader("✍️ Firmas y Datos de Cierre")
    col_f_c1, col_f_c2 = st.columns(2)
    with col_f_c1:
        st.markdown("**La Constructora**")
        firmaConstructora = st.text_input("Nombre (Constructora):")
        dniConstructora = st.text_input("DNI (Constructora):")
        
        st.markdown("**La Propiedad**")
        firmaPropiedad = st.text_input("Nombre (Propiedad):")
        dniPropiedad = st.text_input("DNI (Propiedad):")
    
    with col_f_c2:
        st.markdown("**Dirección de Obra**")
        firmaDireccion = st.text_input("Nombre (D. Obra):")
        dniDireccion = st.text_input("DNI (D. Obra):")
        
        st.markdown("**Dirección de Ejecución**")
        firmaEjecucion = st.text_input("Nombre (D. Ejecución):")
        dniEjecucion = st.text_input("DNI (D. Ejecución):")

    # --- PROCESAMIENTO Y GENERACIÓN ---
    st.markdown("---")
    if st.button("🚀 Procesar Acta y Preparar Descargas"):
        
        # Cargar la plantilla con docxtpl
        doc = DocxTemplate(TEMPLATE_PATH)
        
        # Preparar las imágenes para meterlas en el contexto de docxtpl
        fotos_contexto = []
        for f in fotos_data:
            # Guardamos la imagen temporalmente en memoria para pasársela a InlineImage
            img_stream = io.BytesIO(f["imagen"].read())
            inline_img = InlineImage(doc, img_stream, width=Inches(3.5)) # Ajusta el tamaño de la foto en el Word
            fotos_contexto.append({
                "FOTO_IMAGEN": inline_img,
                "FOTO_DESCRIPCION": f["descripcion"]
            })
        
        # Mapear todas las variables que tiene tu archivo de Word
        context = {
            "actaNum": actaNum,
            "fecha": fecha,
            "hora": hora,
            "proyecto": proyecto,
            "direccion": direccion,
            "repreEstudio": repreEstudio,
            "EmpConst": EmpConst,
            "propiedad": propiedad,
            "estadoTrabajos": estadoTrabajos,
            "instrucciones": instrucciones,
            "incidencias": incidencias,
            "tareas": tareas,
            "fotos": fotos_contexto,
            "firmaConstructora": firmaConstructora,
            "dniConstructora": dniConstructora,
            "firmaPropiedad": firmaPropiedad,
            "dniPropiedad": dniPropiedad,
            "firmaDireccion": firmaDireccion,
            "dniDireccion": dniDireccion,
            "firmaEjecucion": firmaEjecucion,
            "dniEjecucion": dniEjecucion
        }
        
        # Renderizar el documento (inyectar los datos)
        doc.render(context)
        
        # Guardar el documento resultante en un buffer de memoria
        docx_buffer = io.BytesIO()
        doc.save(docx_buffer)
        docx_buffer.seek(0)
        
        st.success("✨ ¡Acta procesada con éxito!")
        
        # Botón de descarga en formato Word (DOCX)
        st.download_button(
            label="📥 Descargar Acta en formato DOCX",
            data=docx_buffer,
            file_name=f"Acta_{actaNum.replace('/', '-') if actaNum else 'visita'}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
        # Nota sobre el PDF
        st.info("ℹ️ *Nota sobre la descarga en PDF:* Para realizar la conversión a PDF directamente en la web de forma idéntica al diseño del Word, se requiere tener instalado un entorno de conversión en el servidor (como LibreOffice o una pasarela Cloud). Si ejecutas esto de forma local en Windows, puedes añadir la librería `docx2pdf` para exportarlo en un clic.")
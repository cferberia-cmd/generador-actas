import streamlit as st
import datetime
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from st_signature_pad import st_signature_pad
from PIL import Image
import io
import os
import base64

# Configuración de pantalla centrada ideal para smartphones
st.set_page_config(page_title="Generador de Actas", layout="centered")

st.title("🏗️ Gestor de Actas de Obra Táctil")
st.write("Rellena las secciones consecutivamente y firma directamente con el dedo desde tu móvil.")

TEMPLATE_PATH = "Plantilla_Acta_Visita.docx"

if not os.path.exists(TEMPLATE_PATH):
    st.error(f"⚠️ No se encuentra el archivo '{TEMPLATE_PATH}'. Asegúrate de subirlo a tu repositorio de GitHub.")
else:
    # --- BLOQUE 1: DATOS GENERALES ---
    st.markdown("### 📝 1. Información del Acta y Asistentes")
    with st.expander("Abrir / Editar Datos Generales", expanded=True):
        actaNum = st.text_input("Acta Nº:", placeholder="Ej. 10")
        fecha = st.date_input("Fecha de la visita:", datetime.date.today()).strftime("%d/%m/%Y")
        hora = st.text_input("Hora:", placeholder="Ej. 12:00")
        proyecto = st.text_input("Identificación del Proyecto:")
        direccion = st.text_input("Dirección:")
        
        st.markdown("---")
        st.markdown("**Asistentes a la visita**")
        repreEstudio = st.text_input("Representante del Estudio de Arquitectura:")
        EmpConst = st.text_input("Empresa Constructora:")
        propiedad = st.text_input("Propiedad:")

    st.markdown("---")

    # --- BLOQUE 2: REPORTES EXTENSOS ---
    st.markdown("### 📋 2. Texto del Reporte Técnico")
    with st.expander("Abrir / Editar Texto del Acta", expanded=False):
        st.caption("🎙️ Tip móvil: ¡Puedes usar el micrófono de tu teclado para dictar el texto!")
        estadoTrabajos = st.text_area("Estado de los Trabajos:", height=120)
        instrucciones = st.text_area("Instrucciones Técnicas de Dirección:", height=120)
        incidencias = st.text_area("Incidencias y No Conformidades:", height=120)
        tareas = st.text_area("Tareas Asignadas con Responsable y Plazo:", height=120)

    st.markdown("---")

    # --- BLOQUE 3: REPORTAJE FOTOGRÁFICO ---
    st.markdown("### 📸 3. Fotografías de Obra")
    with st.expander("Abrir / Añadir Fotografías", expanded=False):
        st.info("📱 Si pulsas 'Browse files' desde el móvil, podrás abrir la cámara directamente.")
        
        if 'num_fotos' not in st.session_state:
            st.session_state.num_fotos = 1

        def añadir_foto():
            st.session_state.num_fotos += 1

        fotos_data = []
        for i in range(st.session_state.num_fotos):
            st.markdown(f"**Fotografía {i+1}**")
            uploaded_foto = st.file_uploader(f"Capturar imagen {i+1}", type=["jpg", "jpeg", "png"], key=f"foto_{i}")
            desc_foto = st.text_input(f"Descripción de la Foto {i+1}", key=f"desc_{i}", placeholder="Ej. Detalle de ejecución...")
            
            if uploaded_foto is not None:
                st.image(uploaded_foto, caption=f"Vista previa {i+1}", width=250)
                fotos_data.append({
                    "imagen": uploaded_foto,
                    "descripcion": desc_foto
                })
            st.markdown("---")

        st.button("➕ Añadir otra foto", on_click=añadir_foto)

    st.markdown("---")

    # --- BLOQUE 4: PANEL DE FIRMAS TÁCTILES ---
    st.markdown("### ✍️ 4. Panel de Firmas Digitales")
    st.caption("Escribe los datos de los intervinientes y firma con el dedo en el recuadro blanco.")

    # --- CONSTRUCTORA ---
    st.markdown("#### 🛠️ La Constructora")
    firmaConstructora = st.text_input("Nombre Responsable Constructora:")
    dniConstructora = st.text_input("DNI Constructora:")
    st.caption("Firma la Constructora aquí debajo:")
    signature_const = st_signature_pad(key="sig_const")

    st.markdown("---")
    
    # --- PROPIEDAD ---
    st.markdown("#### 🏢 La Propiedad")
    firmaPropiedad = st.text_input("Nombre / Razón Social Propiedad:")
    dniPropiedad = st.text_input("DNI / CIF Propiedad:")
    st.caption("Firma la Propiedad aquí debajo:")
    signature_prop = st_signature_pad(key="sig_prop")

    st.markdown("---")
    
    # --- DIRECCIÓN DE OBRA ---
    st.markdown("#### 📐 Dirección de Obra")
    firmaDireccion = st.text_input("Nombre Director de Obra:")
    dniDireccion = st.text_input("DNI Director de Obra:")
    st.caption("Firma el D. de Obra aquí debajo:")
    signature_dir = st_signature_pad(key="sig_dir")

    st.markdown("---")
    
    # --- DIRECCIÓN DE EJECUCIÓN ---
    st.markdown("#### 🦺 Dirección de Ejecución")
    firmaEjecucion = st.text_input("Nombre Director de Ejecución:")
    dniEjecucion = st.text_input("DNI Director de Ejecución:")
    st.caption("Firma el D. de Ejecución aquí debajo:")
    signature_ejec = st_signature_pad(key="sig_ejec")


    # Función para convertir la firma en formato Base64/DataURL a imagen válida para Word
    def procesar_firma_pad(signature_data, doc_obj):
        if signature_data and signature_data.startswith("data:image/png;base64,"):
            try:
                base64_data = signature_data.split(",")[1]
                img_bytes = base64.b64decode(base64_data)
                img = Image.open(io.BytesIO(img_bytes))
                
                img_buffer = io.BytesIO()
                img.save(img_buffer, format='PNG')
                img_buffer.seek(0)
                return InlineImage(doc_obj, img_buffer, width=Inches(1.8))
            except Exception:
                return ""
        return ""

    # --- BOTÓN DE PROCESADO GENERAL ---
    st.markdown("---")
    if st.button("🚀 Procesar y Generar Acta Final", use_container_width=True):
        with st.spinner("Compilando datos, imágenes y firmas en la plantilla..."):
            doc = DocxTemplate(TEMPLATE_PATH)
            
            # Formatear el listado de imágenes del reportaje fotográfico
            fotos_contexto = []
            for f in fotos_data:
                img_stream = io.BytesIO(f["imagen"].read())
                inline_img = InlineImage(doc, img_stream, width=Inches(3.5))
                fotos_contexto.append({
                    "FOTO_IMAGEN": inline_img,
                    "FOTO_DESCRIPCION": f["descripcion"]
                })
            
            # Extraer firmas de las pizarras táctiles
            img_f_const = procesar_firma_pad(signature_const, doc)
            img_f_prop = procesar_firma_pad(signature_prop, doc)
            img_f_dir = procesar_firma_pad(signature_dir, doc)
            img_f_ejec = procesar_firma_pad(signature_ejec, doc)

            # Construir el diccionario con los nombres exactos de las llaves de tu Word
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
                # Textos por defecto si no dibujan nada
                "firmaConstructora": f"\n{firmaConstructora}" if firmaConstructora else "",
                "dniConstructora": dniConstructora,
                "firmaPropiedad": f"\n{firmaPropiedad}" if firmaPropiedad else "",
                "dniPropiedad": dniPropiedad,
                "firmaDireccion": f"\n{firmaDireccion}" if firmaDireccion else "",
                "dniDireccion": dniDireccion,
                "firmaEjecucion": f"\n{firmaEjecucion}" if firmaEjecucion else "",
                "dniEjecucion": dniEjecucion,
            }
            
            # Si hay trazos táctiles, inyectamos la firma encima del texto
            if img_f_const: context["firmaConstructora"] = img_f_const
            if img_f_prop: context["firmaPropiedad"] = img_f_prop
            if img_f_dir: context["firmaDireccion"] = img_f_dir
            if img_f_ejec: context["firmaEjecucion"] = img_f_ejec

            # Renderizar los datos en la plantilla
            doc.render(context)
            
            # Guardar el resultado en memoria
            docx_buffer = io.BytesIO()
            doc.save(docx_buffer)
            docx_buffer.seek(0)
            
            st.success("✨ ¡Acta generada y firmada con éxito!")
            
            # Botón de descarga
            st.download_button(
                label="📥 Descargar Documento Final (.DOCX)",
                data=docx_buffer,
                file_name=f"Acta_Visita_{actaNum}.docx" if actaNum else "Acta_Visita.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
import streamlit as st
import datetime
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from st_signature_pad import st_signature_pad
from PIL import Image
import io
import os
import base64

st.set_page_config(page_title="Acta de Obra", layout="centered")

st.title("🏗️ Gestor de Actas de Obra")
st.write("Versión estable. Sube todas tus fotos a la vez y firma al final.")

TEMPLATE_PATH = "Plantilla_Acta_Visita.docx"

if not os.path.exists(TEMPLATE_PATH):
    st.error(f"⚠️ No se encuentra el archivo '{TEMPLATE_PATH}'.")
else:
    # --- 1. DATOS GENERALES ---
    st.markdown("### 📝 1. Información del Acta")
    actaNum = st.text_input("Acta Nº:", placeholder="Ej. 10")
    fecha = st.date_input("Fecha:", datetime.date.today()).strftime("%d/%m/%Y")
    hora = st.text_input("Hora:", placeholder="Ej. 12:00")
    proyecto = st.text_input("Proyecto:")
    direccion = st.text_input("Dirección:")
    
    st.markdown("**Asistentes**")
    repreEstudio = st.text_input("Representante Dirección Facultativa:")
    EmpConst = st.text_input("Empresa Constructora:")
    propiedad = st.text_input("Propiedad:")

    st.markdown("---")

    # --- 2. REPORTES ---
    st.markdown("### 📋 2. Reporte Técnico")
    estadoTrabajos = st.text_area("Estado de los Trabajos:")
    instrucciones = st.text_area("Instrucciones Técnicas:")
    incidencias = st.text_area("Incidencias y No Conformidades:")
    tareas = st.text_area("Tareas Asignadas:")

    st.markdown("---")

    # --- 3. FOTOS (NUEVO SISTEMA MÚLTIPLE ESTABLE) ---
    st.markdown("### 📸 3. Reportaje Fotográfico")
    st.info("💡 Pulsa el botón y selecciona TODAS las fotos de la visita de golpe desde tu galería.")
    
    # Este componente permite seleccionar varias fotos a la vez sin romper la pantalla
    uploaded_fotos = st.file_uploader("Seleccionar fotografías", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
    
    fotos_data = []
    
    # Si se han subido fotos, muestra un campo de texto para cada una
    if uploaded_fotos:
        for i, foto in enumerate(uploaded_fotos):
            st.markdown(f"**Foto {i+1}: {foto.name}**")
            st.image(foto, width=200)
            desc_foto = st.text_input(f"Descripción de la foto {i+1}:", key=f"desc_{i}")
            fotos_data.append({
                "imagen": foto,
                "descripcion": desc_foto
            })
            st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("---")

    # --- 4. FIRMAS DIGITALES ---
    st.markdown("### ✍️ 4. Firmas en Pantalla")

    st.markdown("#### 🛠️ La Constructora")
    firmaConstructora = st.text_input("Nombre (Constructora):")
    dniConstructora = st.text_input("DNI (Constructora):")
    signature_const = st_signature_pad(key="sig_const")

    st.markdown("#### 🏢 La Propiedad")
    firmaPropiedad = st.text_input("Nombre (Propiedad):")
    dniPropiedad = st.text_input("DNI (Propiedad):")
    signature_prop = st_signature_pad(key="sig_prop")

    st.markdown("#### 📐 Dirección de Obra")
    firmaDireccion = st.text_input("Nombre (D. Obra):")
    dniDireccion = st.text_input("DNI (D. Obra):")
    signature_dir = st_signature_pad(key="sig_dir")

    st.markdown("#### 🦺 Dirección de Ejecución")
    firmaEjecucion = st.text_input("Nombre (D. Ejecución):")
    dniEjecucion = st.text_input("DNI (D. Ejecución):")
    signature_ejec = st_signature_pad(key="sig_ejec")

    # Función para procesar la firma
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

    # --- BOTÓN DE GENERACIÓN ---
    st.markdown("---")
    if st.button("🚀 Generar Acta", use_container_width=True):
        with st.spinner("Creando documento Word..."):
            doc = DocxTemplate(TEMPLATE_PATH)
            
            # Procesar las fotos
            fotos_contexto = []
            for f in fotos_data:
                f["imagen"].seek(0)
                img_stream = io.BytesIO(f["imagen"].read())
                inline_img = InlineImage(doc, img_stream, width=Inches(3.5))
                fotos_contexto.append({
                    "FOTO_IMAGEN": inline_img,
                    "FOTO_DESCRIPCION": f["descripcion"]
                })
            
            # Extraer las firmas
            img_f_const = procesar_firma_pad(signature_const, doc)
            img_f_prop = procesar_firma_pad(signature_prop, doc)
            img_f_dir = procesar_firma_pad(signature_dir, doc)
            img_f_ejec = procesar_firma_pad(signature_ejec, doc)

            context = {
                "actaNum": actaNum, "fecha": fecha, "hora": hora,
                "proyecto": proyecto, "direccion": direccion,
                "repreEstudio": repreEstudio, "EmpConst": EmpConst, "propiedad": propiedad,
                "estadoTrabajos": estadoTrabajos, "instrucciones": instrucciones,
                "incidencias": incidencias, "tareas": tareas,
                "fotos": fotos_contexto,
                "firmaConstructora": f"\n{firmaConstructora}" if firmaConstructora else "", "dniConstructora": dniConstructora,
                "firmaPropiedad": f"\n{firmaPropiedad}" if firmaPropiedad else "", "dniPropiedad": dniPropiedad,
                "firmaDireccion": f"\n{firmaDireccion}" if firmaDireccion else "", "dniDireccion": dniDireccion,
                "firmaEjecucion": f"\n{firmaEjecucion}" if firmaEjecucion else "", "dniEjecucion": dniEjecucion,
            }
            
            if img_f_const: context["firmaConstructora"] = img_f_const
            if img_f_prop: context["firmaPropiedad"] = img_f_prop
            if img_f_dir: context["firmaDireccion"] = img_f_dir
            if img_f_ejec: context["firmaEjecucion"] = img_f_ejec

            doc.render(context)
            
            docx_buffer = io.BytesIO()
            doc.save(docx_buffer)
            docx_buffer.seek(0)
            
            st.success("✨ ¡Acta lista para descargar!")
            
            st.download_button(
                label="📥 Descargar Acta DOCX",
                data=docx_buffer,
                file_name=f"Acta_{actaNum}.docx" if actaNum else "Acta_Visita.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
import streamlit as st
import datetime
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from st_signature_pad import st_signature_pad
from PIL import Image
import io
import os
import base64

# Configuración estricta de página
st.set_page_config(page_title="Generador de Actas de Obra", layout="centered")

st.title("🏗️ Gestor de Actas de Obra (Versión Enterprise)")
st.write("Estructura lineal optimizada para mitigar errores de renderizado en dispositivos móviles.")

TEMPLATE_PATH = "Plantilla_Acta_Visita.docx"

if not os.path.exists(TEMPLATE_PATH):
    st.error(f"⚠️ Archivo crítico ausente: '{TEMPLATE_PATH}'. Por favor, verifícalo en tu repositorio.")
else:
    # --- BLOQUE 1: DATOS GENERALES ---
    st.markdown("### 📝 1. Información del Acta y Asistentes")
    with st.container(border=True):
        actaNum = st.text_input("Acta Nº:", placeholder="Ej. 10", key="main_acta_num")
        fecha = st.date_input("Fecha de la visita:", datetime.date.today(), key="main_date").strftime("%d/%m/%Y")
        hora = st.text_input("Hora:", placeholder="Ej. 12:00", key="main_hour")
        proyecto = st.text_input("Identificación del Proyecto:", key="main_project")
        direccion = st.text_input("Dirección:", key="main_address")
        
        st.markdown("**Asistentes Intervinientes**")
        repreEstudio = st.text_input("Representante del Estudio de Arquitectura:", key="main_repre")
        EmpConst = st.text_input("Empresa Constructora:", key="main_emp")
        propiedad = st.text_input("Propiedad:", key="main_prop")

    # --- BLOQUE 2: REPORTES TÉCNICOS ---
    st.markdown("### 📋 2. Bloques de Reporte Técnico")
    with st.container(border=True):
        estadoTrabajos = st.text_area("Estado de los Trabajos:", height=120, key="txt_estado")
        instrucciones = st.text_area("Instrucciones Técnicas de Dirección:", height=120, key="txt_inst")
        incidencias = st.text_area("Incidencias y No Conformidades:", height=120, key="txt_inc")
        tareas = st.text_area("Tareas Asignadas con Responsable y Plazo:", height=120, key="txt_tareas")

    # --- BLOQUE 3: REPORTAJE FOTOGRÁFICO CONTROLADO ---
    st.markdown("### 📸 3. Reportaje Fotográfico de Obra")
    
    # Inicialización segura del estado de la sesión
    if 'num_fotos' not in st.session_state:
        st.session_state.num_fotos = 1

    fotos_data = []

    # Contenedor estático para aislar la manipulación dinámica del DOM
    with st.container(border=True):
        for i in range(st.session_state.num_fotos):
            # Forzamos una caja interna por cada foto para que React no mezcle componentes
            with st.container(border=True):
                st.markdown(f"**Elemento Fotográfico #{i+1}**")
                
                # Llaves compuestas e inmutables para blindar el Virtual DOM de React
                uploaded_foto = st.file_uploader(
                    f"Seleccionar/Capturar Imagen {i+1}", 
                    type=["jpg", "jpeg", "png"], 
                    key=f"uploader_foto_stable_id_{i}"
                )
                desc_foto = st.text_input(
                    f"Descripción de la Captura {i+1}", 
                    key=f"desc_foto_stable_id_{i}", 
                    placeholder="Describa el detalle observado..."
                )
                
                if uploaded_foto is not None:
                    st.image(uploaded_foto, caption=f"Previsualización #{i+1}", width=240)
                    fotos_data.append({
                        "imagen": uploaded_foto,
                        "descripcion": desc_foto
                    })

        # Callback limpio para evitar re-renders destructivos del DOM
        def callback_incrementar_fotos():
            st.session_state.num_fotos += 1

        st.button("➕ Añadir Casilla para Nueva Fotografía", on_click=callback_incrementar_fotos, key="btn_add_foto")

    # --- BLOQUE 4: PANELES DE FIRMA INDEPENDIENTES ---
    st.markdown("### ✍️ 4. Validación de Firmas Digitales")

    # Aislamos cada panel de firma de manera estricta para evitar fallos de 'removeChild'
    with st.container(border=True):
        st.markdown("#### 🛠️ Representación: La Constructora")
        firmaConstructora = st.text_input("Nombre Responsable:", key="input_name_const")
        dniConstructora = st.text_input("DNI / NIE:", key="input_dni_const")
        signature_const = st_signature_pad(key="canvas_signature_constructora")

    with st.container(border=True):
        st.markdown("#### 🏢 Representación: La Propiedad")
        firmaPropiedad = st.text_input("Nombre / Razón Social:", key="input_name_prop")
        dniPropiedad = st.text_input("DNI / CIF:", key="input_dni_prop")
        signature_prop = st_signature_pad(key="canvas_signature_propiedad")

    with st.container(border=True):
        st.markdown("#### 📐 Dirección Facultativa: Dirección de Obra")
        firmaDireccion = st.text_input("Nombre Director de Obra:", key="input_name_dir")
        dniDireccion = st.text_input("DNI Técnico:", key="input_dni_dir")
        signature_dir = st_signature_pad(key="canvas_signature_direccion")

    with st.container(border=True):
        st.markdown("#### 🦺 Dirección Facultativa: Dirección de Ejecución")
        firmaEjecucion = st.text_input("Nombre Director de Ejecución:", key="input_name_ejec")
        dniEjecucion = st.text_input("DNI Técnico:", key="input_dni_ejec")
        signature_ejec = st_signature_pad(key="canvas_signature_ejecucion")

    # --- PIPELINE DE PROCESAMIENTO Y COMPILACIÓN ---
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

    st.markdown("---")
    
    if st.button("🚀 Compilar y Generar Acta Oficial", use_container_width=True, key="btn_submit_global"):
        with st.spinner("Ejecutando pipeline de renderizado sobre la plantilla Word..."):
            try:
                doc = DocxTemplate(TEMPLATE_PATH)
                
                # Procesar vector de imágenes
                fotos_contexto = []
                for f in fotos_data:
                    f["imagen"].seek(0)
                    img_stream = io.BytesIO(f["imagen"].read())
                    inline_img = InlineImage(doc, img_stream, width=Inches(3.5))
                    fotos_contexto.append({
                        "FOTO_IMAGEN": inline_img,
                        "FOTO_DESCRIPCION": f["descripcion"]
                    })
                
                # Mapear firmas electrónicas desde los buffers canvas
                img_f_const = procesar_firma_pad(signature_const, doc)
                img_f_prop = procesar_firma_pad(signature_prop, doc)
                img_f_dir = procesar_firma_pad(signature_dir, doc)
                img_f_ejec = procesar_firma_pad(signature_ejec, doc)

                # Diccionario de contexto estructurado para docxtpl
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
                    "firmaConstructora": f"\n{firmaConstructora}" if firmaConstructora else "",
                    "dniConstructora": dniConstructora,
                    "firmaPropiedad": f"\n{firmaPropiedad}" if firmaPropiedad else "",
                    "dniPropiedad": dniPropiedad,
                    "firmaDireccion": f"\n{firmaDireccion}" if firmaDireccion else "",
                    "dniDireccion": dniDireccion,
                    "firmaEjecucion": f"\n{firmaEjecucion}" if firmaEjecucion else "",
                    "dniEjecucion": dniEjecucion,
                }
                
                if img_f_const: context["firmaConstructora"] = img_f_const
                if img_f_prop: context["firmaPropiedad"] = img_f_prop
                if img_f_dir: context["firmaDireccion"] = img_f_dir
                if img_f_ejec: context["firmaEjecucion"] = img_f_ejec

                # Renderizar e inyectar datos en caliente
                doc.render(context)
                
                # Serialización en memoria
                docx_buffer = io.BytesIO()
                doc.save(docx_buffer)
                docx_buffer.seek(0)
                
                st.success("✨ ¡El acta ha sido procesada y compilada correctamente sin errores de ejecución!")
                
                st.download_button(
                    label="📥 Descargar Acta Generada (.DOCX)",
                    data=docx_buffer,
                    file_name=f"Acta_Visita_{actaNum}.docx" if actaNum else "Acta_Visita.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="btn_download_final"
                )
            except Exception as e:
                st.error(f"❌ Error crítico en el backend durante el renderizado del Word: {str(e)}")
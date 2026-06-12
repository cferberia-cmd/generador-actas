import streamlit as st
import datetime
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from streamlit_canvas import st_canvas
from PIL import Image
import io
import os

st.set_page_config(page_title="Actas de Obra Móvil", layout="centered")

st.title("🏗️ Acta de Visita de Obra")
st.write("Versión táctil con captura de firmas digitales.")

TEMPLATE_PATH = "Plantilla_Acta_Visita.docx"

if not os.path.exists(TEMPLATE_PATH):
    st.error(f"Falta el archivo de plantilla '{TEMPLATE_PATH}' en el servidor.")
else:
    # --- PESTAÑAS MÓVILES ---
    tab1, tab2, tab3, tab4 = st.tabs(["📝 Datos", "📋 Bloques", "📸 Fotos", "✍️ Firmas Digitales"])

    # PESTAÑA 1: DATOS
    with tab1:
        st.subheader("Datos de la Visita")
        actaNum = st.text_input("Acta Nº:", placeholder="Ej. 01/2026")
        fecha = st.date_input("Fecha:", datetime.date.today()).strftime("%d/%m/%Y")
        hora = st.text_input("Hora:", placeholder="Ej. 10:30")
        proyecto = st.text_input("Proyecto:")
        direccion = st.text_input("Dirección:")

        st.subheader("Asistentes")
        repreEstudio = st.text_input("Rep. Estudio Arquitectura:")
        EmpConst = st.text_input("Empresa Constructora:")
        propiedad = st.text_input("Propiedad:")

    # PESTAÑA 2: CONTENIDO
    with tab2:
        st.subheader("Reporte de la Visita")
        estadoTrabajos = st.text_area("Estado de los Trabajos:", height=120)
        instrucciones = st.text_area("Instrucciones Técnicas:", height=120)
        incidencias = st.text_area("Incidencias / No Conformidades:", height=120)
        tareas = st.text_area("Tareas Asignadas:", height=120)

    # PESTAÑA 3: FOTOS
    with tab3:
        st.subheader("Reporte Fotográfico")
        if 'num_fotos' not in st.session_state:
            st.session_state.num_fotos = 1

        def añadir_foto():
            st.session_state.num_fotos += 1

        fotos_data = []
        for i in range(st.session_state.num_fotos):
            st.markdown(f"---")
            uploaded_foto = st.file_uploader(f"Capturar imagen {i+1}", type=["jpg", "jpeg", "png"], key=f"foto_{i}")
            desc_foto = st.text_input(f"Descripción Foto {i+1}", key=f"desc_{i}")
            
            if uploaded_foto is not None:
                st.image(uploaded_foto, width=200)
                fotos_data.append({"imagen": uploaded_foto, "descripcion": desc_foto})

        st.button("➕ Añadir nueva foto", on_click=añadir_foto)

    # PESTAÑA 4: CAJAS DE FIRMA TÁCTILES
    with tab4:
        st.subheader("Cierre y Firmas en Pantalla")
        st.caption("Usa el dedo o un lápiz óptico para firmar en los recuadros negros.")

        # Configuración común para los lienzos de firma
        canvas_kwargs = {
            "fill_color": "rgba(255, 255, 255, 0)",
            "stroke_width": 3,
            "stroke_color": "#000000", # Tinta negra para el documento
            "background_color": "#EEEEEE", # Fondo gris claro en la web para ver el límite
            "height": 150,
            "width": 320,
            "drawing_mode": "freedraw",
            "key": None
        }

        # --- CONSTRUCTORA ---
        st.markdown("### 🛠️ La Constructora")
        firmaConstructora = st.text_input("Nombre Responsable (Constructora):")
        dniConstructora = st.text_input("DNI (Constructora):")
        st.caption("Firma aquí:")
        canvas_const = st_canvas(**{**canvas_kwargs, "key": "canvas_const"})

        # --- PROPIEDAD ---
        st.markdown("### 🏢 La Propiedad")
        firmaPropiedad = st.text_input("Nombre / Razón Social (Propiedad):")
        dniPropiedad = st.text_input("DNI / CIF (Propiedad):")
        st.caption("Firma aquí:")
        canvas_prop = st_canvas(**{**canvas_kwargs, "key": "canvas_prop"})

        # --- DIRECCIÓN DE OBRA ---
        st.markdown("### 📐 Dirección de Obra")
        firmaDireccion = st.text_input("Nombre (D. Obra):")
        dniDireccion = st.text_input("DNI (D. Obra):")
        st.caption("Firma aquí:")
        canvas_dir = st_canvas(**{**canvas_kwargs, "key": "canvas_dir"})

        # --- DIRECCIÓN DE EJECUCIÓN ---
        st.markdown("### 🦺 Dirección de Ejecución")
        firmaEjecucion = st.text_input("Nombre (D. Ejecución):")
        dniEjecucion = st.text_input("DNI (D. Ejecución):")
        st.caption("Firma aquí:")
        canvas_ejec = st_canvas(**{**canvas_kwargs, "key": "canvas_ejec"})


        # --- FUNCIÓN INTERNA PARA CONVERTIR EL TRAZO A IMAGEN PARA WORD ---
        def procesar_firma_tactic(canvas_data, doc_obj):
            if canvas_data.image_data is not None:
                # El canvas devuelve una matriz de píxeles. Convertimos a imagen PIL.
                img = Image.fromarray(canvas_data.image_data.astype('uint8'), 'RGBA')
                # Verificar si el usuario ha dibujado algo (para no meter cuadros vacíos)
                # Si está vacío o dibujado, lo pasamos a un buffer de memoria
                img_byte_arr = io.BytesIO()
                img.save(img_byte_arr, format='PNG')
                img_byte_arr.seek(0)
                return InlineImage(doc_obj, img_byte_arr, width=Inches(1.8)) # Ajuste de tamaño en el Word
            return ""

        # --- PROCESADO FINAL ---
        st.markdown("---")
        if st.button("🚀 Procesar Documento con Firmas", use_container_width=True):
            with st.spinner("Incrustando datos y firmas..."):
                doc = DocxTemplate(TEMPLATE_PATH)
                
                # Procesar Bloque de Fotos
                fotos_contexto = []
                for f in fotos_data:
                    img_stream = io.BytesIO(f["imagen"].read())
                    inline_img = InlineImage(doc, img_stream, width=Inches(3.5))
                    fotos_contexto.append({"FOTO_IMAGEN": inline_img, "FOTO_DESCRIPCION": f["descripcion"]})
                
                # Transformar los lienzos táctiles en imágenes para el Word
                img_f_const = procesar_firma_tactic(canvas_const, doc)
                img_f_prop = procesar_firma_tactic(canvas_prop, doc)
                img_f_dir = procesar_firma_tactic(canvas_dir, doc)
                img_f_ejec = procesar_firma_tactic(canvas_ejec, doc)

                # Unificar todo el contexto
                context = {
                    "actaNum": actaNum, "fecha": fecha, "hora": hora, "proyecto": proyecto, "direccion": direccion,
                    "repreEstudio": repreEstudio, "EmpConst": EmpConst, "propiedad": propiedad,
                    "estadoTrabajos": estadoTrabajos, "instrucciones": instrucciones, "incidencias": incidencias, "tareas": tareas,
                    "fotos": fotos_contexto,
                    "firmaConstructora": firmaConstructora, "dniConstructora": dniConstructora,
                    "firmaPropiedad": firmaPropiedad, "dniPropiedad": dniPropiedad,
                    "firmaDireccion": firmaDireccion, "dniDireccion": dniDireccion,
                    "firmaEjecucion": firmaEjecucion, "dniEjecucion": dniEjecucion,
                    # Variables de imagen de las firmas:
                    "firma_const_img": img_f_const,
                    "firma_prop_img": img_f_prop,
                    "firma_dir_img": img_f_dir,
                    "firma_ejec_img": img_f_ejec
                }
                
                doc.render(context)
                
                docx_buffer = io.BytesIO()
                doc.save(docx_buffer)
                docx_buffer.seek(0)
                
                st.success("✨ ¡Acta firmada y cerrada con éxito!")
                
                st.download_button(
                    label="📥 Descargar Acta Firmada (.DOCX)",
                    data=docx_buffer,
                    file_name=f"Acta_Firmada_{actaNum.replace('/', '-') if actaNum else 'final'}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
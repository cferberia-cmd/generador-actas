import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
import io

st.set_page_config(page_title="ACTA ROSA - Generador", layout="wide")

st.title("📝 Generador de Actas de Visita de Obra")
st.markdown("---")

# 1. Creamos pestañas para aislar los elementos y congelar el DOM en móviles
tab_datos, tab_fotos, tab_firmas = st.tabs([
    "📋 1. Datos y Textos", 
    "📸 2. Fotografías", 
    "✍️ 3. Firmas y Guardar"
])

# --- PESTAÑA 1: DATOS GENERALES Y TEXTOS ---
with tab_datos:
    st.subheader("Información General del Proyecto")
    col1, col2 = st.columns(2)
    with col1:
        actaNum = st.text_input("Acta Nº", key="actaNum")
        fecha = st.text_input("Fecha de la visita", key="fecha", placeholder="Ej. 12 de junio de 2026")
        hora = st.text_input("Hora", key="hora", placeholder="Ej. 17:30")
    with col2:
        proyecto = st.text_input("Identificación del Proyecto", key="proyecto")
        direccion = st.text_input("Dirección de la Obra", key="direccion")

    st.subheader("Asistentes a la Reunión")
    col3, col4, col5 = st.columns(3)
    with col3:
        repreEstudio = st.text_input("Representante del Estudio", key="repreEstudio")
    with col4:
        EmpConst = st.text_input("Empresa Constructora", key="EmpConst")
    with col5:
        propiedad = st.text_input("Propiedad", key="propiedad")

    st.subheader("Desarrollo de la Visita")
    estadoTrabajos = st.text_area("Estado de los Trabajos", key="estadoTrabajos", height=150)
    instrucciones = st.text_area("Instrucciones Técnicas de Dirección", key="instrucciones", height=150)
    incidencias = st.text_area("Incidencias y No Conformidades", key="incidencias", height=150)
    tareas = st.text_area("Tareas Asignadas (Responsable y Plazo)", key="tareas", height=150)


# --- PESTAÑA 2: FOTOGRAFÍAS (Subidor múltiple nativo) ---
with tab_fotos:
    st.subheader("Registro Fotográfico")
    st.info("💡 Consejo: Selecciona todas las fotos de golpe desde tu galería. Al estar en esta pestaña aislada, no romperá la aplicación.")
    
    uploaded_files = st.file_uploader(
        "Subir Fotografías", 
        accept_multiple_files=True, 
        type=["png", "jpg", "jpeg"], 
        key="fotos_uploader"
    )
    
    # Lista para almacenar las fotos y sus descripciones
    fotos_contexto_lista = []
    
    if uploaded_files:
        st.markdown("---")
        for i, file in enumerate(uploaded_files):
            col_img, col_txt = st.columns([1, 2])
            with col_img:
                st.image(file, caption=f"Imagen {i+1}", width=200)
            with col_txt:
                # Usamos una clave estática única basada en el índice para evitar conflictos
                desc_foto = st.text_input(f"Descripción para la foto {i+1}", key=f"input_desc_{i}")
                fotos_contexto_lista.append({
                    "file_raw": file,
                    "descripcion": desc_foto
                })


# --- PESTAÑA 3: FIRMAS Y EXPORTACIÓN ---
with tab_firmas:
    st.subheader("Datos de Cierre y Firmantes")
    
    col6, col7 = st.columns(2)
    with col6:
        st.markdown("**La Constructora**")
        firmaConstructora = st.text_input("Nombre Firmante Constructora", key="firmaConstructora")
        dniConstructora = st.text_input("DNI Constructora", key="dniConstructora")
        
        st.markdown("**La Propiedad**")
        firmaPropiedad = st.text_input("Nombre Firmante Propiedad", key="firmaPropiedad")
        dniPropiedad = st.text_input("DNI Propiedad", key="dniPropiedad")
        
    with col7:
        st.markdown("**Dirección Facultativa**")
        firmaDireccion = st.text_input("Dirección de Obra (Nombre)", key="firmaDireccion")
        dniDireccion = st.text_input("DNI Dirección de Obra", key="dniDireccion")
        firmaEjecucion = st.text_input("Dirección de Ejecución (Nombre)", key="firmaEjecucion")
        dniEjecucion = st.text_input("DNI Dirección de Ejecución", key="dniEjecucion")

    st.markdown("---")
    
    # Botón único de generación de documento
    if st.button("🚀 Generar Acta de Visita (DOCX)", type="primary"):
        if not actaNum:
            st.warning("⚠️ Es recomendable introducir al menos el Número de Acta.")
            
        try:
            # 1. Cargamos la plantilla Word original
            doc = DocxTemplate("Plantilla_Acta_Visita.docx")
            
            # 2. Procesamos las imágenes para docxtpl
            fotos_finales = []
            for f in fotos_contexto_lista:
                # Convertimos el archivo subido en un objeto de imagen compatible con Word
                img_stream = io.BytesIO(f["file_raw"].read())
                inline_img = InlineImage(doc, img_stream, width=Inches(4.5)) 
                
                fotos_finales.append({
                    "FOTO_IMAGEN": inline_img,
                    "FOTO_DESCRIPCION": f["descripcion"]
                })
            
            # 3. Empaquetamos todo el diccionario de datos
            context = {
                "actaNum": actaNum, "fecha": fecha, "hora": hora, "proyecto": proyecto, "direccion": direccion,
                "repreEstudio": repreEstudio, "EmpConst": EmpConst, "propiedad": propiedad,
                "estadoTrabajos": estadoTrabajos, "instrucciones": instrucciones, "incidencias": incidencias, "tareas": tareas,
                "fotos": fotos_finales,
                "firmaConstructora": firmaConstructora, "dniConstructora": dniConstructora,
                "firmaPropiedad": firmaPropiedad, "dniPropiedad": dniPropiedad,
                "firmaDireccion": firmaDireccion, "dniDireccion": dniDireccion,
                "firmaEjecucion": firmaEjecucion, "dniEjecucion": dniEjecucion
            }
            
            # 4. Inyectamos los datos en la plantilla
            doc.render(context)
            
            # 5. Guardamos el resultado en memoria para la descarga
            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            
            st.success("🎉 ¡Acta procesada correctamente sin errores de sistema!")
            st.download_button(
                label="📥 Descargar Acta de Visita",
                data=output,
                file_name=f"Acta_Visita_{actaNum}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except FileNotFoundError:
            st.error("❌ No se encuentra el archivo 'Plantilla_Acta_Visita.docx' en el servidor. Asegúrate de subirlo junto a app.py.")
        except Exception as e:
            st.error(f"❌ Error inesperado al procesar la plantilla: {str(e)}")
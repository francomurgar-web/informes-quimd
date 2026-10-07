from datetime import datetime
import io
import os
import random
import tempfile
import openpyxl
from PIL import Image as PILImage
import streamlit as st

# Configuración de la página
st.set_page_config(
    page_title="Generador de Informes - QUIMDES",
    page_icon="📋",
    layout="centered",
)

st.title("📋 Generador de Informes QUIMDES S.A.C.")
st.markdown(
    "Sistema multi-plantilla unificado con persistencia de datos segura y optimizada."
)

# --- RUTA DE PLANTILLAS UNIFICADAS ---
current_dir = os.path.dirname(os.path.abspath(__file__))

dict_plantillas = {
    "3D": os.path.join(current_dir, "plantilla_3d.xlsx"),
    "CAMPANA, FILTROS, DUCTO Y MOTOR": os.path.join(
        current_dir, "plantilla_campana_filtros_ducto_motor.xlsx"
    ),
    "DESINSECTACION": os.path.join(current_dir, "plantilla_desinsectacion.xlsx"),
    "DESINSECTACION Y DESRATIZACION": os.path.join(
        current_dir, "plantilla_desinsectacion_desratizacion.xlsx"
    ),
    "PRESERVANTE DE MADERA": os.path.join(
        current_dir, "plantilla_preservante_madera.xlsx"
    ),
    "TRAMPA DE GRASA": os.path.join(current_dir, "plantilla_trampa_grasa.xlsx"),
}

# --- BARRA LATERAL: Configuración del Servicio ---
st.sidebar.header("⚙️ Configuración del Servicio")
tipo_servicio = st.sidebar.selectbox(
    "Tipo de Servicio:", list(dict_plantillas.keys())
)
ruta_plantilla = dict_plantillas[tipo_servicio]

if "n_informe" not in st.session_state:
    st.session_state.n_informe = f"{random.randint(1000, 9999)}-10-2026"

n_informe = st.sidebar.text_input(
    "N° de Informe", value=st.session_state.n_informe
)
fecha_servicio = st.sidebar.date_input("Fecha del Servicio", value=datetime.now())

# --- FORMULARIO PRINCIPAL BLINDADO (st.form) ---
with st.form(key="form_informe"):
    st.subheader("1. 🏢 Datos del Cliente")
    col_c1, col_c2 = st.columns(2)

    with col_c1:
        razon_social = st.text_input("Razón Social")
        direccion = st.text_input("Dirección")
        ruc = st.text_input("RUC (Opcional)")

    with col_c2:
        razon_comercial = st.text_input("Razón Comercial / Local")
        telefono = st.text_input("Teléfono / Contacto")
        correo = st.text_input("Correo Electrónico")

    st.subheader("2. 🛠️ Datos Operativos")
    col_o1, col_o2 = st.columns(2)

    with col_o1:
        encargado = st.text_input("Técnico Encargado", value="KEVIN")
        producto = st.text_input("Productos / Materiales")

    with col_o2:
        area_tratada = st.text_input("Áreas Tratadas")
        metodo = st.text_input("Método de Aplicación")

    grado_accion = st.selectbox(
        "Grado de Acción",
        ["OK / POSTERIOR", "URGENTE", "INMEDIATO", "PRONTO", "PREVENTIVO"],
    )

    observaciones = st.text_area("Descripción de Actividades", height=120)
    recomendaciones = st.text_area("Recomendaciones", height=120)

    st.markdown("---")
    st.subheader("3. 📷 Evidencia Fotográfica")
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        foto_antes = st.file_uploader(
            "Subir Foto ANTES", type=["png", "jpg", "jpeg"]
        )
    with col_f2:
        foto_despues = st.file_uploader(
            "Subir Foto DESPUÉS", type=["png", "jpg", "jpeg"]
        )

    submit_button = st.form_submit_button(
        "🚀 Generar Informe con su Plantilla Exclusiva", type="primary"
    )

# --- PROCESAMIENTO Y MAPEO AL HACER CLIC ---
if submit_button:
    if not razon_social.strip():
        st.warning(
            "⚠️ Por favor, ingresa la Razón Social del cliente antes de"
            " generar el informe."
        )
    elif not os.path.exists(ruta_plantilla):
        st.error(
            f"⚠ No se encontró la plantilla para este servicio en la ruta:"
            f" {ruta_plantilla}. Asegúrate de tener los archivos Excel en la"
            f" carpeta."
        )
    else:
        try:
            wb = openpyxl.load_workbook(ruta_plantilla)
            ws = wb.active

            def escribir_celda(hoja, coordenada, valor):
                celda = hoja[coordenada]
                if type(celda).__name__ == "MergedCell":
                    for rango in hoja.merged_cells.ranges:
                        if coordenada in rango:
                            hoja.cell(
                                row=rango.min_row,
                                column=rango.min_col,
                                value=valor,
                            )
                            return
                else:
                    celda.value = valor

            # Mapeo unificado estricto para todas las plantillas
            escribir_celda(ws, "A1", f"INFORME {n_informe} | ADM-QUIMDES")

            # Datos del Cliente
            escribir_celda(ws, "B5", razon_social)
            escribir_celda(ws, "N5", razon_comercial)
            escribir_celda(ws, "B6", direccion)
            escribir_celda(ws, "N6", telefono)
            escribir_celda(ws, "B7", ruc)
            escribir_celda(ws, "N7", correo)

            # Datos Operativos
            escribir_celda(ws, "B11", tipo_servicio)
            escribir_celda(ws, "N11", fecha_servicio.strftime("%d/%m/%Y"))
            escribir_celda(ws, "B12", producto)
            escribir_celda(ws, "N12", encargado)
            escribir_celda(ws, "B13", metodo)
            escribir_celda(ws, "N13", area_tratada)

            escribir_celda(ws, "L20", observaciones)
            escribir_celda(ws, "O20", recomendaciones)
            escribir_celda(ws, "P20", grado_accion)

            # Uso de directorio temporal seguro para procesar imágenes
            with tempfile.TemporaryDirectory() as temp_dir:

                def procesar_y_guardar_imagen(uploaded_file, nombre_archivo):
                    path_img = os.path.join(temp_dir, nombre_archivo)
                    with PILImage.open(uploaded_file) as img:
                        if img.mode in ("RGBA", "P"):
                            img = img.convert("RGB")
                        img.thumbnail((1000, 1000))  # Optimización de tamaño
                        img.save(path_img, "JPEG", quality=85)
                    return path_img

                if foto_antes is not None:
                    path_antes = procesar_y_guardar_imagen(
                        foto_antes, "antes.jpg"
                    )
                    img_a = openpyxl.drawing.image.Image(path_antes)
                    img_a.width = 300
                    img_a.height = 220
                    ws.add_image(img_a, "J49")

                if foto_despues is not None:
                    path_despues = procesar_y_guardar_imagen(
                        foto_despues, "despues.jpg"
                    )
                    img_d = openpyxl.drawing.image.Image(path_despues)
                    img_d.width = 300
                    img_d.height = 220
                    ws.add_image(img_d, "J66")

                # Guardar archivo generado directamente en un buffer de memoria (BytesIO)
                buffer_excel = io.BytesIO()
                wb.save(buffer_excel)
                buffer_excel.seek(0)

            nombre_salida = (
                f"INFORME {n_informe} ({razon_social or 'Cliente'}).xlsx"
            )

            st.success(
                "¡Informe generado con éxito bajo el estándar unificado y datos"
                " seguros!"
            )

            st.download_button(
                label="📥 Descargar Informe Excel",
                data=buffer_excel,
                file_name=nombre_salida,
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )

        except Exception as e:
            st.error(f"Error procesando la plantilla: {e}")

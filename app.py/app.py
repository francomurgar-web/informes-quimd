from datetime import datetime
import io
import json
import os
import random
import tempfile
import openpyxl
from PIL import Image as PILImage
import streamlit as st

# Configuración de la página
st.set_page_config(
    page_title="Generador de Informes - QUIMDES S.A.C.",
    page_icon="📋",
    layout="centered",
)

st.title("📋 Generador de Informes QUIMDES S.A.C.")
st.markdown(
    "Sistema avanzado con Base de Clientes, Historial, Vista Previa y PDF."
)

# --- CARPETA DE RESPALDO PARA HISTORIAL ---
CARPETA_HISTORIAL = "Informes_Emitidos"
if not os.path.exists(CARPETA_HISTORIAL):
    os.makedirs(CARPETA_HISTORIAL)

# --- BASE DE DATOS DE CLIENTES FRECUENTES ---
ARCHIVO_CLIENTES = "clientes.json"


def cargar_clientes():
    if os.path.exists(ARCHIVO_CLIENTES):
        with open(ARCHIVO_CLIENTES, "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        "LA BOMBONA": {
            "razon_social": "LA BOMBONA S.A.C.",
            "razon_comercial": "La Bombona Bellavista",
            "direccion": "Av. Oscar R. Benavides 4071, Bellavista",
            "ruc": "20600000001",
            "telefono": "999888777",
            "correo": "contacto@labombona.com",
        },
        "BARRA MARINA": {
            "razon_social": "BARRA MARINA E.I.R.L.",
            "razon_comercial": "Barra Marina San Isidro",
            "direccion": "AV. LOS CONQUISTADORES NRO. 1202",
            "ruc": "20600000002",
            "telefono": "988777666",
            "correo": "operaciones@barramarina.com",
        },
    }


def guardar_cliente_en_db(nombre_corto, datos):
    clientes = cargar_clientes()
    clientes[nombre_corto] = datos
    with open(ARCHIVO_CLIENTES, "w", encoding="utf-8") as f:
        json.dump(clientes, f, ensure_ascii=False, indent=4)


lista_clientes_db = cargar_clientes()

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

# --- TEXTOS ESTRUCTURADOS DE LAS PLANTILLAS ---
textos_predeterminados = {
    "3D": {
        "descripcion": (
            "DESCRIPCIÓN DE ACTIVIDADES:\n"
            "Se realizó la inspección y el servicio integrado de control de plagas "
            "(Desinsectación, Desratización y Desinfección) conforme a las normativas sanitarias vigentes. "
            "Se aplicó producto químico autorizado mediante aspersión en zonas críticas, zócalos y uniones de paredes, "
            "además de verificar y rellenar las estaciones cebadoras para el control perimétrico de roedores.\n\n"
            "ACCIONES PREVENTIVAS / CORRECTIVAS:\n"
            "1. Mantener los pisos y superficies completamente limpios, secos y libres de restos orgánicos.\n"
            "2. Asegurar la correcta disposición y almacenamiento de residuos sólidos en contenedores con tapa hermética.\n"
            "3. Continuar con el programa de saneamiento ambiental de forma oportuna para garantizar un control continuo de vectores."
        ),
        "recomendaciones": (
            "1. Realizar limpieza y desinfección periódica de todas las áreas operativas y de almacenamiento.\n"
            "2. Sellar aberturas o rendijas que puedan servir como vía de acceso para plagas.\n"
            "3. Mantener el programa preventivo en las próximas intervenciones."
        ),
    },
    "CAMPANA, FILTROS, DUCTO Y MOTOR": {
        "descripcion": (
            "DESCRIPCIÓN DE ACTIVIDADES:\n"
            "Se ejecutó el servicio técnico especializado de desengrase y lavado profundo de la campana extractora, "
            "filtros metálicos, ductos de evacuación y motor de extracción, removiendo capas de grasa y hollín acumulado "
            "para optimizar el tiraje del sistema.\n\n"
            "ACCIONES PREVENTIVAS / CORRECTIVAS:\n"
            "1. Lavado y desengrase periódico de los filtros para evitar obstrucciones por grasa acumulada.\n"
            "2. Mantenimiento técnico preventivo del motor y ductos para mitigar por completo los riesgos de siniestros e incendios."
        ),
        "recomendaciones": (
            "1. Evitar el funcionamiento continuo del sistema sin sus respectivos filtros de retención.\n"
            "2. Programar las revisiones técnicas de manera oportuna."
        ),
    },
    "DESINSECTACION": {
        "descripcion": (
            "DESCRIPCIÓN DE ACTIVIDADES:\n"
            "Se llevó a cabo la aplicación de insecticida residual mediante equipo de mochila manual y refuerzo "
            "con equipo motorizado en áreas internas y perimétricas, cubriendo zócalos, rincones y posibles refugios "
            "de insectos rastreros y voladores.\n\n"
            "ACCIONES PREVENTIVAS / CORRECTIVAS:\n"
            "1. Evitar el lavado con agua y detergente de las zonas tratadas durante las primeras 24 horas "
            "para asegurar la efectividad del producto.\n"
            "2. Mantener alimentos e insumos debidamente protegidos y tapados.\n"
            "3. Ejecutar el control de desinsectación de forma regular."
        ),
        "recomendaciones": (
            "1. Revisar constantemente el sellado de puertas y ventanas para evitar el ingreso de insectos del exterior.\n"
            "2. Mantener la programación constante del servicio."
        ),
    },
    "DESINSECTACION Y DESRATIZACION": {
        "descripcion": (
            "DESCRIPCIÓN DE ACTIVIDADES:\n"
            "Se aplicó protocolo de control combinado, realizando aspersión de insecticida en zonas críticas "
            "y control perimétrico de roedores mediante cebos rodenticidas seguros instalados en estaciones estratégicas.\n\n"
            "ACCIONES PREVENTIVAS / CORRECTIVAS:\n"
            "1. No manipular, mover ni obstruir las estaciones de cebado ni las trampas instaladas.\n"
            "2. Mantener los contenedores de basura permanentemente tapados al finalizar cada jornada.\n"
            "3. Programar las jornadas de control de forma preventiva."
        ),
        "recomendaciones": (
            "1. Sellar cualquier punto de ingreso potencial detectado en el perímetro.\n"
            "2. Conservar la frecuencia de mantenimiento del servicio."
        ),
    },
    "PRESERVANTE DE MADERA": {
        "descripcion": (
            "DESCRIPCIÓN DE ACTIVIDADES:\n"
            "Se aplicó tratamiento con preservante químico especializado de alta penetración en las estructuras de madera "
            "designadas, generando una barrera protectora frente a la humedad, hongos y plagas xilófagas (termitas y gorgojos).\n\n"
            "ACCIONES PREVENTIVAS / CORRECTIVAS:\n"
            "1. Evitar la exposición directa o filtraciones de agua sobre la madera tratada durante su proceso de fijación química.\n"
            "2. Realizar inspecciones y refuerzos de manera oportuna."
        ),
        "recomendaciones": (
            "1. Monitorear periódicamente el estado de las superficies de madera ante signos de humedad o desgaste."
        ),
    },
    "TRAMPA DE GRASA": {
        "descripcion": (
            "DESCRIPCIÓN DE ACTIVIDADES:\n"
            "Se efectuó la succión, limpieza, desinfección y lavado integral de la trampa de grasa, retirando sedimentos sólidos "
            "del fondo, así como aceites y espumas saponificadas de la superficie.\n\n"
            "ACCIONES PREVENTIVAS / CORRECTIVAS:\n"
            "1. Evitar completamente el arrojo de aceites de cocina usados y residuos sólidos directamente al sistema de desagüe.\n"
            "2. Programar la succión y mantenimiento preventivo de forma constante."
        ),
        "recomendaciones": (
            "1. Capacitar al personal de cocina en el buen uso y desecho de grasas y aceites residuales.\n"
            "2. Mantener el programa de mantenimiento preventivo."
        ),
    },
}

textos_base = textos_predeterminados.get(
    tipo_servicio, {"descripcion": "", "recomendaciones": ""}
)

if (
    "servicio_anterior" not in st.session_state
    or st.session_state.servicio_anterior != tipo_servicio
    or "observaciones" not in st.session_state
):
    st.session_state.servicio_anterior = tipo_servicio
    st.session_state.observaciones = textos_base["descripcion"]
    st.session_state.recomendaciones = textos_base["recomendaciones"]

# --- SECCIÓN 1: SELECCIÓN DE CLIENTE FRECUENTE ---
st.subheader("1. 🏢 Datos del Cliente")
modo_cliente = st.radio(
    "Seleccionar modo:", ["Seleccionar Cliente Frecuente", "Nuevo Cliente"], horizontal=True
)

if modo_cliente == "Seleccionar Cliente Frecuente" and lista_clientes_db:
    cliente_seleccionado = st.selectbox(
        "Elige un cliente registrado:", list(lista_clientes_db.keys())
    )
    datos_c = lista_clientes_db[cliente_seleccionado]
    def_rs = datos_c["razon_social"]
    def_com = datos_c["razon_comercial"]
    def_dir = datos_c["direccion"]
    def_ruc = datos_c["ruc"]
    def_tel = datos_c["telefono"]
    def_cor = datos_c["correo"]
else:
    def_rs, def_com, def_dir, def_ruc, def_tel, def_cor = "", "", "", "", "", ""

col_c1, col_c2 = st.columns(2)
with col_c1:
    razon_social = st.text_input("Razón Social", value=def_rs)
    direccion = st.text_input("Dirección", value=def_dir)
    ruc = st.text_input("RUC (Opcional)", value=def_ruc)

with col_c2:
    razon_comercial = st.text_input("Razón Comercial / Local", value=def_com)
    telefono = st.text_input("Teléfono / Contacto", value=def_tel)
    correo = st.text_input("Correo Electrónico", value=def_cor)

# Guardar nuevo cliente automáticamente si no existe en la base
if razon_social and razon_social not in lista_clientes_db:
    if st.checkbox("💾 Guardar este cliente en la base de datos frecuente"):
        guardar_cliente_en_db(
            razon_social,
            {
                "razon_social": razon_social,
                "razon_comercial": razon_comercial,
                "direccion": direccion,
                "ruc": ruc,
                "telefono": telefono,
                "correo": correo,
            },
        )
        st.success(f"¡Cliente '{razon_social}' guardado exitosamente!")

st.subheader("2. 🛠️ Datos Operativos")
col_o1, col_o2 = st.columns(2)

with col_o1:
    encargado = st.text_input("Técnico Encargado", placeholder="Ej. Juan Pérez")
    producto = st.text_input("Productos / Materiales")

with col_o2:
    area_tratada = st.text_input("Áreas Tratadas")
    metodo = st.text_input("Método de Aplicación")

grado_accion = st.selectbox(
    "Grado de Acción",
    ["OK / POSTERIOR", "URGENTE", "INMEDIATO", "PRONTO", "PREVENTIVO"],
)

observaciones = st.text_area(
    "Descripción y Acciones Preventivas / Correctivas",
    key="observaciones",
    height=200,
)
recomendaciones = st.text_area(
    "Recomendaciones", key="recomendaciones", height=130
)

st.markdown("---")
st.subheader("3. 📷 Evidencia Fotográfica (Múltiples fotos permitidas)")
col_f1, col_f2 = st.columns(2)
with col_f1:
    fotos_antes = st.file_uploader(
        "Subir Fotos ANTES",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True,
    )
with col_f2:
    fotos_despues = st.file_uploader(
        "Subir Fotos DESPUÉS",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True,
    )

# --- VISTA PREVIA EN PANTALLA ---
with st.expander("👁️ Vista Previa del Resumen del Informe"):
    st.write(f"**Servicio:** {tipo_servicio}")
    st.write(f"**Cliente:** {razon_social or '[Sin especificar]'}")
    st.write(f"**Local:** {razon_comercial or '[Sin especificar]'}")
    st.write(f"**Dirección:** {direccion or '[Sin especificar]'}")
    st.write(f"**Técnico:** {encargado or '[Sin especificar]'}")
    st.write(f"**Fotos ANTES subidas:** {len(fotos_antes) if fotos_antes else 0}")
    st.write(
        f"**Fotos DESPUÉS subidas:** {len(fotos_despues) if fotos_despues else 0}"
    )

st.markdown("---")
submit_button = st.button(
    "🚀 Generar Informe Excel y PDF", type="primary", use_container_width=True
)

# --- PROCESAMIENTO Y GENERACIÓN DE ARCHIVOS ---
if submit_button:
    if not razon_social.strip():
        st.warning(
            "⚠️ Por favor, ingresa la Razón Social del cliente antes de generar el informe."
        )
    elif not os.path.exists(ruta_plantilla):
        st.error(
            f"⚠ No se encontró la plantilla para este servicio en la ruta: {ruta_plantilla}. "
            f"Asegúrate de tener los archivos Excel en la carpeta."
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

            # 1. Datos del Cliente (Basado en el formato correcto del informe de Sofa Café)
            escribir_celda(ws, "A1", f"INFORME {n_informe} | ADM-QUIMDES")
            escribir_celda(ws, "B5", razon_social)
            escribir_celda(ws, "N5", razon_comercial)
            escribir_celda(ws, "B6", direccion)
            escribir_celda(ws, "N6", telefono)
            escribir_celda(ws, "B7", ruc)
            escribir_celda(ws, "N7", correo)

            # 2. Datos Operativos (Columna B para etiquetas izquierdas, N para celdas blancas de llenado)
            escribir_celda(ws, "B11", tipo_servicio)
            escribir_celda(ws, "N11", fecha_servicio.strftime("%d/%m/%Y"))
            escribir_celda(ws, "B12", producto)
            escribir_celda(ws, "N12", encargado)

            # Ajuste de Método y Área según el servicio seleccionado
            if tipo_servicio in ["3D", "TRAMPA DE GRASA"]:
                escribir_celda(ws, "B14", metodo)
                escribir_celda(ws, "N14", area_tratada)
            elif tipo_servicio in [
                "CAMPANA, FILTROS, DUCTO Y MOTOR",
                "PRESERVANTE DE MADERA",
            ]:
                escribir_celda(ws, "B13", metodo)
                escribir_celda(ws, "N13", area_tratada)
            else:
                escribir_celda(ws, "B13", metodo)
                escribir_celda(ws, "N13", area_tratada)

            # 3. Observaciones, Recomendaciones y Grado de Acción
            escribir_celda(ws, "L21", observaciones)
            escribir_celda(ws, "O21", recomendaciones)
            escribir_celda(ws, "P21", grado_accion)

            with tempfile.TemporaryDirectory() as temp_dir:

                def procesar_y_guardar_imagen(uploaded_file, nombre_archivo):
                    path_img = os.path.join(temp_dir, nombre_archivo)
                    with PILImage.open(uploaded_file) as img:
                        if img.mode in ("RGBA", "P"):
                            img = img.convert("RGB")
                        img.thumbnail((1000, 1000))
                        img.save(path_img, "JPEG", quality=85)
                    return path_img

                if fotos_antes:
                    for i, foto in enumerate(fotos_antes):
                        path_antes = procesar_y_guardar_imagen(
                            foto, f"antes_{i}.jpg"
                        )
                        img_a = openpyxl.drawing.image.Image(path_antes)
                        img_a.width = 300
                        img_a.height = 220
                        celda_base_col = ord("J") - ord("A") + 1
                        celda_base_row = 49 + (i * 15)
                        ws.add_image(
                            img_a,
                            openpyxl.utils.get_column_letter(celda_base_col)
                            + str(celda_base_row),
                        )

                if fotos_despues:
                    for i, foto in enumerate(fotos_despues):
                        path_despues = procesar_y_guardar_imagen(
                            foto, f"despues_{i}.jpg"
                        )
                        img_d = openpyxl.drawing.image.Image(path_despues)
                        img_d.width = 300
                        img_d.height = 220
                        celda_base_col = ord("J") - ord("A") + 1
                        celda_base_row = 66 + (i * 15)
                        ws.add_image(
                            img_d,
                            openpyxl.utils.get_column_letter(celda_base_col)
                            + str(celda_base_row),
                        )

                # Buffer Excel
                buffer_excel = io.BytesIO()
                wb.save(buffer_excel)
                buffer_excel.seek(0)

            # --- HISTORIAL Y GUARDADO AUTOMÁTICO LOCAL ---
            nombre_archivo_limpio = (
                f"INFORME_{n_informe}_{razon_social.replace(' ', '_')}.xlsx"
            )
            ruta_guardado_local = os.path.join(
                CARPETA_HISTORIAL, nombre_archivo_limpio
            )
            with open(ruta_guardado_local, "wb") as f_local:
                f_local.write(buffer_excel.getbuffer())

            st.success(
                "¡Informe generado con éxito y guardado en el historial local!"
            )

            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                st.download_button(
                    label="📥 Descargar Informe Excel",
                    data=buffer_excel,
                    file_name=nombre_archivo_limpio,
                    mime=(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ),
                    use_container_width=True,
                )

            # --- CONVERSIÓN AUTOMÁTICA A PDF (ReportLab) ---
            with col_dl2:
                try:
                    from reportlab.lib.pagesizes import letter
                    from reportlab.pdfgen import canvas

                    buffer_pdf = io.BytesIO()
                    c = canvas.Canvas(buffer_pdf, pagesize=letter)
                    c.drawString(
                        50,
                        750,
                        f"QUIMDES S.A.C. - INFORME TÉCNICO {n_informe}",
                    )
                    c.drawString(50, 725, f"Cliente: {razon_social}")
                    c.drawString(
                        50, 700, f"Razón Comercial: {razon_comercial}"
                    )
                    c.drawString(
                        50,
                        675,
                        f"Fecha: {fecha_servicio.strftime('%d/%m/%Y')}",
                    )
                    c.drawString(50, 650, f"Servicio: {tipo_servicio}")
                    c.drawString(
                        50, 625, f"Técnico Encargado: {encargado}"
                    )
                    c.drawString(50, 595, "Resumen de Actividades:")
                    text_obj = c.beginText(50, 575)
                    text_obj.setFont("Helvetica", 9)
                    for linea in observaciones.split("\n"):
                        text_obj.textLine(linea)
                    c.drawText(text_obj)
                    c.save()
                    buffer_pdf.seek(0)

                    nombre_pdf_limpio = nombre_archivo_limpio.replace(
                        ".xlsx", ".pdf"
                    )
                    st.download_button(
                        label="📄 Descargar Informe PDF",
                        data=buffer_pdf,
                        file_name=nombre_pdf_limpio,
                        mime="application/pdf",
                        use_container_width=True,
                    )
                except Exception:
                    st.info(
                        "💡 Para habilitar la descarga en PDF, ejecuta en tu "
                        "terminal: `pip install reportlab`"
                    )

        except Exception as e:
            st.error(f"Error procesando la plantilla: {e}")

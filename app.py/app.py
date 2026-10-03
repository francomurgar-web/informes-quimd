import random
import streamlit as st
import pandas as pd
from datetime import datetime
import openpyxl
import os

st.set_page_config(page_title="Generador de Informes - QUIMDES", page_icon="📋", layout="centered")
st.title("🧪 Generador de Informes QUIMDES S.A.C.")
st.write("Selecciona el servicio, completa los datos del cliente y genera tu Excel corporativo al instante.")

st.sidebar.header("Configuración del Servicio")
tipo_servicio = st.sidebar.selectbox("Tipo de Servicio:", [
    "LIMPIEZA DE TRAMPA DE GRASA",
    "DESINSECTACION Y DESRATIZACION",
    "LIMPIEZA Y DESINFECCION DE RESERVORIO",
    "LIMPIEZA DE REDES DE DESAGUE"
])

import random  # <--- (Asegúrate de tener esta importación arriba junto con los demás 'import')

# Generar un número de informe aleatorio (ejemplo: un número entre 1000 y 9999 seguido del año actual)
numero_aleatorio = random.randint(1000, 9999)
n_informe_defecto = f"{numero_aleatorio}-08-2026"

n_informe = st.sidebar.text_input("N° de Informe", value=n_informe_defecto)
fecha_servicio = st.sidebar.date_input("Fecha del Servicio", value=datetime.now())

st.subheader("1. Datos del Cliente")
razon_social = st.text_input("Razón Social", value="SOFA CAFÉ")
razon_comercial = st.text_input("Razón Comercial / Local", value="SOFA CAFÉ - MIRAFLORES")
direccion = st.text_input("Dirección", value="JR. ROCA VERGALLO 308 - MAGDALENA")
telefono = st.text_input("Teléfono / Contacto", value="")
correo = st.text_input("Correo Electrónico", value="")

st.subheader("2. Datos Operativos")
col1, col2 = st.columns(2)
with col1:
    encargado = st.text_input("Técnico Encargado", value="KEVIN")
    supervisor = st.text_input("Supervisor(a)", value="JIMENA")
with col2:
    area_tratada = st.text_input("Áreas Tratadas", value="COCINA")
    grado_accion = st.selectbox("Grado de Acción", ["OK/POSTERIOR", "URGENTE", "PREVENTIVO"])

producto = st.text_input("Productos Utilizados", value="DESENGRASANTE / BIODET D3")
metodo = st.text_input("Método de Aplicación", value="LIMPIEZA MECANICA Y LAVADO A PRESION")

observaciones = st.text_area("Descripción de Actividades / Observaciones", value="Se realizó la inspección inicial, limpieza profunda, retiro de residuos y desinfección total de la zona, dejándola operativa.")
recomendaciones = st.text_area("Recomendaciones", value="1. Realizar el mantenimiento de forma MENSUAL.\n2. Evitar el vertido directo de residuos sólidos y grasas en los desagües.")

st.markdown("---")
if st.button("🚀 Generar Informe en Excel", type="primary"):
    plantilla_base = "INFORME   0456-08-2026 (SOFA CAFÉ - MIRAFLORES ) - TRAMPA DE GRASA.xlsx"
    if os.path.exists(plantilla_base):
        try:
            wb = openpyxl.load_workbook(plantilla_base)
            ws = wb.active
            
            # Función segura para escribir en celdas (evita errores si alguna celda está combinada)
            def escribir_celda(hoja, coordenada, valor):
                celda = hoja[coordenada]
                if type(celda).__name__ == 'MergedCell':
                    # Si es una celda combinada, buscamos la celda principal del rango
                    for rango in hoja.merged_cells.ranges:
                        if coordenada in rango:
                            hoja.cell(row=rango.min_row, column=rango.min_col, value=valor)
                            return
                else:
                    celda.value = valor

            # Asignación de datos utilizando la función segura
            escribir_celda(ws, 'B4', razon_social)
            escribir_celda(ws, 'L3', razon_comercial)
            escribir_celda(ws, 'B5', direccion)
            escribir_celda(ws, 'L4', telefono)
            escribir_celda(ws, 'L5', correo)
            escribir_celda(ws, 'B10', tipo_servicio)
            escribir_celda(ws, 'M10', fecha_servicio.strftime('%d/%m/%Y'))
            escribir_celda(ws, 'B11', producto)
            escribir_celda(ws, 'B13', metodo)
            escribir_celda(ws, 'L13', area_tratada)
            escribir_celda(ws, 'M19', recomendaciones)
            escribir_celda(ws, 'L19', grado_accion)
            
            nombre_salida = f"INFORME {n_informe} ({razon_social}).xlsx"
            wb.save(nombre_salida)
            
            st.success(f"¡Informe generado correctamente!")
            
            with open(nombre_salida, "rb") as f:
                st.download_button(
                    label="📥 Descargar Archivo Excel",
                    data=f,
                    file_name=nombre_salida,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Ocurrió un error al procesar el Excel: {e}")
    else:
        st.error(f"No se encontró el archivo plantilla base en la carpeta.")
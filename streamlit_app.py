import streamlit as st
import openpyxl
from io import BytesIO
import traceback

st.set_page_config(page_title="Consolidador con Plantilla", layout="wide")
st.title("Consolidador REM con Plantilla Dedicada")

st.info(
    "**Instrucciones de Uso:**\n"
    "1. **Carga la Plantilla:** Sube el archivo Excel que tiene el formato y los textos finales.\n"
    "2. **Carga los Archivos de Datos:** Sube uno o más archivos REM de los cuales se sumarán los valores numéricos.\n"
    "3. **Consolida:** Presiona el botón para generar el archivo final."
)

# --- PASO 1: Cargar la Plantilla ---
st.header("1. Sube tu Archivo de Plantilla")
template_file = st.file_uploader(
    "Este archivo definirá la estructura y el formato del resultado. Solo se modificarán sus celdas numéricas.",
    type=["xlsx", "xlsm"]
)

# --- PASO 2: Cargar los Archivos de Datos ---
st.header("2. Sube los Archivos con Datos a Sumar")
data_files = st.file_uploader(
    "De estos archivos solo se tomarán los valores numéricos para ser sumados.",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True
)

# --- Inicialización del Estado de Sesión ---
if 'processed_file' not in st.session_state:
    st.session_state.processed_file = None
    st.session_state.file_name = None

# --- PASO 3: Botón para Procesar ---
if st.button("✨ Consolidar Datos en la Plantilla"):

    # Validaciones
    if not template_file:
        st.error("❌ Por favor, sube un archivo de plantilla para continuar.")
    elif not data_files:
        st.error("❌ Por favor, sube al menos un archivo de datos para consolidar.")
    else:
        with st.spinner("Procesando archivos..."):
            try:
                # --- Lógica de Suma de Datos ---
                sumas_consolidadas = {}
                progress_bar = st.progress(0, text="Paso 1/2: Sumando datos de los archivos...")
                for i, data_file in enumerate(data_files):
                    data_file.seek(0)
                    wb_data = openpyxl.load_workbook(data_file, data_only=True)
                    for hoja_nombre in wb_data.sheetnames:
                        if hoja_nombre not in sumas_consolidadas:
                            sumas_consolidadas[hoja_nombre] = {}
                        ws_data = wb_data[hoja_nombre]
                        for fila in ws_data.iter_rows():
                            for celda in fila:
                                if isinstance(celda.value, (int, float)):
                                    ref = celda.coordinate
                                    sumas_consolidadas[hoja_nombre][ref] = sumas_consolidadas[hoja_nombre].get(ref, 0) + celda.value
                    progress_bar.progress((i + 1) / len(data_files), text=f"Paso 1/2: Leyendo {data_file.name}")

                # --- Lógica de Escritura en la Plantilla ---
                progress_bar.progress(0, text="Paso 2/2: Escribiendo resultados en la plantilla...")
                template_file.seek(0)
                # Cargamos la plantilla, descartando las macros para una salida .xlsx limpia
                wb_final = openpyxl.load_workbook(template_file, keep_vba=False)

                for hoja_nombre, celdas in sumas_consolidadas.items():
                    if hoja_nombre in wb_final.sheetnames:
                        ws_final = wb_final[hoja_nombre]
                        for celda_ref, valor_sumado in celdas.items():
                            try:
                                # Modificamos solo el valor, el formato se mantiene
                                ws_final[celda_ref].value = valor_sumado
                            except Exception:
                                # Ignora si no se puede escribir en la celda
                                pass
                    else:
                        st.warning(f"La hoja '{hoja_nombre}' existe en los datos pero no en la plantilla y fue ignorada.")

                # --- Guardar y preparar para la descarga ---
                progress_bar.progress(1.0, text="Generando archivo final...")
                output = BytesIO()
                wb_final.save(output)
                output.seek(0)
                
                # Guardar en la sesión para que el botón de descarga funcione
                st.session_state.processed_file = output
                st.session_state.file_name = "Rem_consolidados.xlsx"
                progress_bar.empty()

            except Exception as e:
                st.error(f"Ocurrió un error crítico durante el proceso: {e}")
                st.error(traceback.format_exc())
                st.session_state.processed_file = None

# --- Lógica para mostrar el botón de descarga ---
if st.session_state.processed_file is not None:
    st.success("✅ ¡Consolidación completada! Ya puedes descargar el archivo.")
    
    st.download_button(
        label="📥 Descargar Rem_consolidados.xlsx",
        data=st.session_state.processed_file,
        file_name=st.session_state.file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if st.button("Limpiar y empezar de nuevo"):
        st.session_state.processed_file = None
        st.session_state.file_name = None
        st.rerun()
